# 透视表与切片器

透视表的来源变了怎样刷新、依赖透视结果的公式怎样读、只保留而不支持编辑的原生透视怎样处理。参数查 `help pivot.refresh`、`help slicer.select`、`help query.pivots`。

## 两类透视

- **程序支持刷新的透视**：来源变化了，必须在改动它的**同一步**里执行 `pivot.refresh`，否则这一步以 `PIVOT_REFRESH_REQUIRED` 失败。追加明细、改透视来源和刷新透视要写进同一个 `operations` 数组。
- **程序只保留、不支持编辑的原生透视**：来源变化时缓存被标为失效，报告 `PIVOT_NOT_REFRESHED`，结果状态为需要刷新，不一定报 `PIVOT_REFRESH_REQUIRED`。这类保留下来的结果不是本次刷新的结果；需要时请用户在 Excel 里刷新。

透视不会自动刷新。追加明细后刷新的写法是在同一步里写 `[{"op":"table.append",…},{"op":"pivot.refresh","sheet":…,"name":…}]`。`inspect` 的 `workbook.pivots` 列出透视。

## 依赖透视结果的公式

依赖透视结果的公式按工作簿当前保存的透视结果计算，与 Excel 重算而不刷新透视相同：公式本次算出了值，但透视没有因此刷新，其与当前来源的一致性未获确认。要让这些公式反映刷新后的透视，就在执行 `pivot.refresh` 或 `slicer.select` 的同一步里，用该步自带的 `read` 读取它们；结果在该步的 `result.ranges`。之后单独读取（`session.read`、`wb.read`）默认不计算，读到的是存值（`s:"cache"`）；让那次读取计算，公式按保存的透视结果重算，透视仍是 `stored`。

例外：透视缓存设了打开时刷新（`refreshOnLoad="1"`），或同一步里先刷新、之后又改了它的来源而没有再刷新时，引用它的公式不可得，诊断是 `PIVOT_RESULT_UNAVAILABLE`。以透视结果为来源再建透视（直接或经过公式），上游透视也要先在同一步刷新。

下面是配方 `recipes/pivot-by-slicer` 的做法（片段：它要求 `Pivot` 表上有名为 `RegionSlicer` 的切片器、`Report` 表上有依赖透视结果的公式）：

```js
// 片段：脚本函数体内，每次选择都用这一步自带的 read
const step = await wb.run({
  operations: [ { op: 'slicer.select', sheet: 'Pivot', name: 'RegionSlicer', items: ['East'] } ],
  read: [ { sheet: 'Report', range: 'B1:B2' } ],
});
if (!step.ok) throw new Error(`${step.error.code}: ${step.error.message}`);
const values = step.result.ranges[0].cells;
```

读取透视结果区域时，紧凑形式里透视结果的格是 `{"v":值,"pivot":"current"|"stored"}`：`current` 是这一步刷新出来的结果，`stored` 是文件里存着的结果，不是本次的。

## 新增工作表之后的透视状态

只新增工作表时，只有数据源公式的保存值**确实保留下来**、来源没有变化的透视，才不会因为这次新增工作表而被标成需要刷新。数据源里有会被作废的公式时（哪几类见 [deliverable.md](deliverable.md) 的"哪些公式保留存值"），不能这样保证；原来就已经失效的状态，也不会因为新增工作表而恢复。

满足条件、这一步也没有刷新透视时，做了修改的那一步的编辑响应里，`pivots[]` 的对应项为 `sourceChange: unchanged`，`resultStatus` 保留原来的状态，原生缓存不置为失效：原来是 `stored_unverified` 的仍是它（`diagnostic` 为 `PIVOT_RESULT_UNVERIFIED`）；原来是 `refresh_required` 的仍是 `refresh_required`，不会恢复。`query` 的 `pivots` 查询不返回 `sourceChange` 和单数的 `diagnostic`，它的状态在 `pivots[i].items[j].resultStatus`。`stored_unverified` 的意思是来源没变，但并没有重新证明原透视结果正确；它仍是文件里保存的结果，不是本次结果。

## 报告时

依赖未刷新透视的公式有值，但只与透视保存的结果一致。报告里要写明：公式已按工作簿当前保存的透视结果重新计算；透视未刷新，其与当前来源的一致性未获确认。列出哪些值依赖透视；来源变过（`PIVOT_NOT_REFRESHED`、`refresh_required`）的，请用户在 Excel 里刷新后再采信。打开时刷新的透视，引用它的公式不可得，`all` 默认仍发布其他可靠结果并给清单；`requireCompleteCalculation:true` 才拒绝。
