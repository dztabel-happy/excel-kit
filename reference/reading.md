# 读单元格与取值

三个读取入口的差别、返回的格长什么样、怎样在代码里只取当前值。来源标记 `s` 的四个取值见 `SKILL.md` 第 3 条，字段查 `help read` 与 `help read.cells`，一页读不完见 [paging.md](paging.md)。

## 三个入口

- **`query` 请求的 `read`**：一份请求列出同一个文件的多块区域，一次读完；结果在顶层 `ranges`，顺序与请求一致。一次性 `run` 的 `read` 相同。
- **`wb.read`、`session.read`**：读取项写在 `ranges`，不是 `read`；结果同样在 `ranges`。
- **命令行 `read`**：只接受命令行参数，不接受请求文件，一次一块。每次调用都重新解析整本，不要用它逐页读大表。

分页参数 `offsetRows`、`limitRows` 写在各个范围里；`expectedInputVersion` 与 `timeoutMs` 写在请求外层，不写在范围里，也不写在 `calculation` 里。

## `inspect` 给什么

`inspect` 返回工作簿概况，不计算：`workbook.sheets`（每项有 `name`）、`definedNames`、`tables`、`pivots`、`externalLinks`、`fileType`、`diagnostics`，以及文件版本 `workbook.version`。它不是完整的对象盘点：图表、图片、批注、保护的完整权限等用 `query` 的对应种类查。`wb.inspect()`、`session.inspect` 返回同样的形状。

## 两种形状

- **不写 `form`**：格在 `cells`（按行的二维数组），每格是带 `kind`、`formula`、`result` 的对象。没有本次结果时 `result` 是 `null`，文件里的存值在 `cache`、状态在 `cacheState`；直接取 `result.value` 会抛类型错误。`display.text` 可能由存值格式化而来，先看 `display.source`。
- **`"form":"compact"`**：格在 `rows`。常量就是它的 JSON 值，没有来源标记：真正的空白是 `null`，空文本是 `""`，错误值是 `{"e":"#N/A"}`。公式是 `{"f","v","s"}`，数组结果的成员是 `{"a":锚点,"v","s"}`，模拟运算表的结果另带 `d`，透视结果格是 `{"v","pivot":"current"|"stored"}`，格内图片是 `{"image"}`。它不含格式、显示文本与行可见性，不能与 `include` 同用；体量能小一个数量级。

## 要当前值

- 把读取写进做计算的那一步的 `read`：与这一步的计算是同一次，之后再单独读就是 `cache`。
- 不计算的读取里，有存值的公式标 `cache` 是正常的；报告它时要说明这是文件里保存的上次计算值，没有经过本次验证。文件里本来就没有存值的公式是 `unavailable`，不一定是出错。
- `s` 是 `result` 的值仍可能是错误值 `{"e":…}`。

## 取值片段

下面的函数适用于一次性 `run`、`query`、`read` 的响应，`session.run`、`wb.run` 的步骤响应，以及 `report:"full"` 里的步骤结果。它只返回常量或本次结果；失败、计算不完整、存值、不可得与错误值都抛出原因。

```js
function currentCell(response, rangeIndex = 0, row = 0, column = 0) {
  if (response?.ok !== true) {
    throw new Error(JSON.stringify(response?.error ?? response?.calculation ?? response));
  }
  const report = response.status === 'succeeded' ? response.result : response;
  const status = report?.calculation?.status;
  if (status !== 'complete' && status !== 'not_requested') {
    throw new Error(JSON.stringify(report?.calculation ?? 'Missing full step report'));
  }
  const range = report.ranges?.[rangeIndex];
  if (!range) throw new Error('Missing requested range');
  if (range.form === 'compact') {
    const cell = range.rows?.[row]?.[column];
    if (cell === undefined) throw new Error('Cell outside returned page');
    if (cell !== null && typeof cell === 'object') {
      if ('image' in cell) throw new Error('Image is not a scalar');
      if ('s' in cell && cell.s !== 'result') throw new Error('Formula is not current: ' + cell.s);
      if ('pivot' in cell && cell.pivot !== 'current') throw new Error('Pivot is not current');
      const value = 'v' in cell ? cell.v : cell;
      if (value !== null && typeof value === 'object') throw new Error(JSON.stringify(value));
      return value;
    }
    return cell;
  }
  const cell = range.cells?.[row]?.[column];
  if (!cell || cell.result === null || cell.result === undefined) throw new Error('No current scalar');
  if (cell.pivotResult && cell.pivotResult.source !== 'current_pivot_result') throw new Error('Pivot is not current');
  if (cell.result.type === 'error') throw new Error(String(cell.result.value));
  return cell.result.value;
}
```

传什么进去：

- 一次性命令：整个响应。脚本里：`currentCell(await wb.run(step))`。
- 会话：先处理应答顶层的协议 `error`，再传 `reply.result.response`。
- 脚本的报告：一次性脚本先看外层 `ok` 与 `script.ok`，再传 `response.script.steps[i]`；会话脚本传 `response.steps[i]`。默认的报告没有每步的 `result`，要加 `report:"full"`。
- `script.result` 是脚本自己返回的 JSON，不要假定它有 `ranges`。
