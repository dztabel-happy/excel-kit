# 脚本

发布方式见入口第 2 条；请求字段与成员查 `help script`、`help wb.<成员>`。可运行脚本见包内 `recipes/README.md`。

## 请求与脚本体

- 脚本是 JavaScript **异步函数体**，不是模块：可以用顶层 `await` 和 `return`，不能 `import`、`export`。参数是 `wb` 和 `args`（请求里的 `arguments`）。长脚本放进文件用 `scriptPath`。
- 只读的脚本也必须写 `output`，给一个不存在的路径；没有成功的 `wb.run` 就不写文件，`output.saved:false`。
- `session.script` 在未保存状态上执行。默认保留成功修改，`session.save` 才写盘；`trial:true` 只带回结果并恢复原状态。见[会话](sessions.md)。
- 修改只有 `wb.run` 一条路，用的是与 `run` 相同的操作。`wb` 没有通用文件接口：CSV 用 `data` 操作导入工作表再读回，图片与附件用 `image-extract`、`attachment-extract` 命令提取。

## 默认写 `strict:true`

像普通 JavaScript 的异常：任何一次 `wb.run`、`wb.read`、`wb.query`、`wb.inspect` 失败都抛出 `ExcelKitError`，不用逐个判 `ok`。

```sh
cat > flag.js <<'EOF'
const page = await wb.read({ ranges: [{ sheet: 'Sheet1', range: 'A2:C200', form: 'compact' }] });
let flagged = 0;
for (const [i, row] of page.ranges[0].rows.entries()) {
  if (typeof row[1] !== 'number' || row[1] <= args.limit) continue;
  await wb.run({ op: 'format', sheet: 'Sheet1', range: `A${i + 2}:C${i + 2}`, style: { fill: 'FFC7CE' } });
  flagged++;
}
return { flagged };
EOF
cat > flag.json <<'EOF'
{"input":"input.xlsx","output":"flagged.xlsx","strict":true,"calculation":{"mode":"none"},
 "arguments":{"limit":1900},"scriptPath":"flag.js"}
EOF
./excel-kit script flag.json > flag.out.json
```

- 抛出的对象有 `method`（`run`、`read`、`query`、`inspect`）、`code`、`message`、`details` 和原响应 `response`。
- 不捕获：脚本以 `SCRIPT_FAILED` 失败并整体回滚，`error.message` 形如 `wb.read failed: SHEET_NOT_FOUND: …`，原错误在 `error.details.cause`（`{method, code, message, details}`，过大的 `details` 会省去）。
- 要自己处理某一次失败，就只捕获那一次：`const step = await wb.run(…).catch(error => error.response);`，拿回的是原响应，再判 `step.ok`。被捕获的失败 `wb.run` 仍计入 `stepFailures`。
- `wb.run` 默认计算不完整也 `ok:true`，`strict:true` 不抛；要拒绝可在打开参数或分组步骤写 `requireCompleteCalculation:true`。独立读取/查询的两种计算不完整：**查询**仍是 `ok:true`、不抛，自己看 `calculation.status`；**读取**抛 `CALCULATION_INCOMPLETE`，`details` 带 `scope` 与 `diagnostics`。未知动态数组的范围可能影响读取或查询位置时，两者仍可能报 `ARRAY_EXTENT_UNAVAILABLE`；`strict:true` 会将该失败抛出。

## 不带 `strict` 时

失败的调用不抛错，返回 `ok:false`：`wb.run` 是 `{ok:false, status:"failed", inputVersion, error}`，其余是 `{ok:false, error}`；计算不完整的读取有 `calculation` 而**没有 `error`**。所以每次取值之前都要判：

```js
const need = (r, what) => {
  if (r.ok) return r;
  throw new Error(`${what}: ${r.error ? r.error.code : 'calculation ' + r.calculation?.status}`);
};
const page = need(await wb.read({ ranges: [{ sheet: 'Sheet1', range: 'A2:C200' }] }), 'read');
```

- 没判就取，通常是一个 `TypeError`：脚本以 `SCRIPT_FAILED` 失败，原错误码只在 `error.details.lastCallFailure` 里。
- 失败之后继续执行并正常返回，成功的步骤照样发布。只有失败的 `wb.run` 计入 `stepFailures`；被忽略的读取、查询失败在报告里看不出来。要整段作废就 `throw`。

## 整体回滚的情形

脚本抛错或有语法错误、线程在返回前退出（没人处理的 Promise 拒绝、定时器回调里抛错）是 `SCRIPT_FAILED`；超过 `scriptTimeoutMs` 是 `SCRIPT_TIMEOUT`；被取消是 `CANCELLED`；`wb.run` 超过次数上限是 `SCRIPT_STEP_LIMIT`（失败的调用也计数）；返回值不能序列化（BigInt、循环引用）是 `SCRIPT_RESULT_NOT_SERIALIZABLE`；某一步出内部错误也整段回滚。返回之前已经发出、没有 `await` 的调用照常执行并计入步骤；返回之后才发生的拒绝不影响结果。

## 响应与发布

一次性 `script` 的外层与 `run` 相同，但**没有 `calculation`**；报告在 `script` 下：

- `script.status` 是 `completed`、`completed_with_failures` 或 `failed`；`script.result` 是返回值（`undefined`、函数、Symbol 记为 `null`）。
- `script.steps` 每次 `wb.run` 一项，默认只有状态、版本和错误；要每步的结果加 `"report":"full"`，体积随步数线性增长。回滚后仍列出已执行的步骤：报告的 `outputVersion` 等于 `inputVersion`，步骤记录保留执行时的版本。
- 一步都没成功：`ok:true`、`output.saved:false`、退出码 0。脚本成功而发布被拒（例如运行期间输入文件被改）：外层 `ok:false`，`error` 是发布的错误，`script.ok` 仍为真。请求在脚本开始前就被拒：只有 `ok` 与 `error`，没有 `script`。
- 所以检查一次性脚本看三样：`ok`、`script.stepFailures`、`output.saved`。退出码 0 包含 `completed_with_failures`。
- 全程持有目标文件的锁，其间别的写入者得到 `FILE_BUSY`。

## 环境

- `console` 的输出进 `script.logs`，不进标准输出；有条数与总量上限，超出时 `logsTruncated:true`。
- `Bun`、`process`、`require`、`fetch` 被遮蔽，只为防止误用，不是安全沙箱：脚本仍能经 `globalThis` 访问宿主，那些文件或网络访问不受事务回滚保护。
- 脚本之间不共享 JavaScript 状态。每次 `wb.read` 都重新解析工作簿：只读需要的范围，页开到上限。
- 没有脚本的预检；要预演就在会话里运行脚本而不保存。
- 步数与整段期限有上限（`limits`）；长循环的办法见 [cost.md](cost.md)。
