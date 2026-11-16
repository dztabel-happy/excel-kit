# 交付件

保存之后怎样核对、文件里每个公式存的是什么、要打开时重算该怎么办。选哪种计算方式见 [calculation.md](calculation.md)。

## 核对交付件

1. 响应：`ok`、`output.saved`、`calculation.status`，逐项看 `warnings`。脚本另看 `script.stepFailures`。
2. `fileVerification` 说的是写出的文件与内存一致（`source:"serialized_in_memory"`）；数值与 Excel 是否一致不在核对之内，`numericalVerification.status` 恒为 `not_performed`。
3. `output.version` 是写出文件的 SHA-256。会话里它与会话版本相等，就确认了写出的内容，不必重新打开。
4. 结构性结果（工作表次序、可见性、保护、名称）对交付件 `inspect` 或 `query`。
5. 关键数值用紧凑读取，确认 `s` 是 `result`。对交付件不计算地读，读到的是存值（`cache`）；要当前值就让这次读取计算。
6. 要证明"其他内容没变"，用同一份读取请求（`none`、紧凑形式）分别读原件和交付件，在本地逐格比较，只比任务可能影响到的区域。

## 一步之后，每个公式存的是什么

| 这个公式 | 文件里 | 下一次不计算的读取 |
|---|---|---|
| 这一步算了 | 当前值 | `cache` |
| 没算，也没受这次修改影响 | 原来的存值 | `cache` |
| 没算，受了影响的普通公式 | 没有值 | `unavailable` |
| 受了影响的模拟运算表 | 旧的存值原样留着 | `cache`，看不出已经失效 |

## 哪些公式保留存值

不在本次计算范围内的公式，没受影响的保留存值。保留下来的是上次的结果，不是本次的；要本次结果仍须计算。

- 只新增工作表（可以同时移动、隐藏工作表）时，绑定没有变化的公式及其下游保留存值。照旧作废的：引用了新表名的公式；依赖表数或表序的公式；动态引用；绑定不明的引用（例如保留下来的三维引用）及其下游。
- 复制、删除、改名工作表与插删行列，保守作废。
- 只引用外部工作簿固定地址的公式，在与它无关的本地编辑之后保留存值；同一公式若还引用了被改的本地格，照常作废。
- 动态引用、绑定不明与结构变化会保守地扩大作废范围：不能凭静态的依赖查询断言别的存值都没受影响。

## `FORMULAS_WITHOUT_VALUES`

发布的文件里有公式没有值时，`warnings` 里有这一项，附数量与示例位置；每步的报告另有 `uncachedFormulas`。原因是编辑作废了存值而本次没有算到/算不出，或输入文件里本来就没有值；用 `targets` 或 `none` 编辑之后常见。

- 它只数没有值的公式。保留着存值的公式不触发它，所以没有这条警告不等于都是本次结果，仍要读回看来源 `s`。
- 处理：工作簿能整本算，就改用 `all` 重新生成；不能，就把需要显示值的区域纳入 `targets`，仍须检查不可得清单，其余在报告里说明。

## 模拟运算表

- 程序原样保留模拟运算表，不计算它。表内的结果与依赖它的本次结果不可得，默认 all 仍保存其他可靠结果。
- 改了它的输入：它的存值已经失效，但仍留在文件里，不计入上面的警告；下一次读取里标成 `cache`。不要把它当作当前值报告，告诉用户运算表要在 Excel 里重算。
- 这时工作簿存的计算方式若是自动（或没有声明），程序替你设 `fullCalcOnLoad`；若是手动或 `autoNoTable`，运算表本身不触发该标志（普通公式缺值仍可能触发），响应的 `cacheInvalidation.dataTables[i].consumerRecalculation` 是 `manual_required`。

## 打开时重算

普通公式不带值地保存时，程序在 auto/autoNoTable 自动设 fullCalcOnLoad:true，覆盖原来的 false/缺省；manual 不改并给 MANUAL_RECALCULATION_REQUIRED。该标志只是请求，不保证消费者支持未知函数或能取得外部材料。可显式选择消费者计算方式：

```json
{"op":"workbook.settings","calculation":{"mode":"auto","fullCalcOnLoad":true}}
```

- 这是一个操作，写进 `operations`。它里面的 `calculation` 是存进文件的设置，与这一步算不算无关。之后的保存保留这项设置。
- 没有普通公式缺值时保留原设置；手动模式即使已经为 true 仍提示人工重算。autoNoTable 的运算表仍须人工重算，与普通公式的自动重算请求并存。
- 查它现在是什么：`query` 的 `"workbookSettings":true`，看 `workbookSettings.calculation.values`。


## 完整不可得清单与同次保存证据

`run` 的 `calculation.unavailable` 给 `count/formulaCount/byReason/examples(≤10)/stepId/report`；计算不可得与保存缺值是两项不同统计，多个原因的计数也不能相加当格数。报告 `{path,sha256,records}` 指向本次发布的独立 JSONL 文件；批次提交、session.save、一次性脚本在顶层 `calculationReport` 给聚合文件。会话步骤只给统计与 stepId，清单先留在内存；预检不写报告文件。

文件第一行是头：schemaVersion、最终输出版本、各步骤 inputVersion/outputVersion、计算条件及摘要、不可得格数 count、压缩记录数 records、recordsSha256 与同次保存核验。之后每行是一条区域记录，带 stepId/sheetId/sheet/range/count/reasons/owner/savedValue。每步按表序、行、列排列，同原因/同保存状态/同所有者的横向连续格压成一段，不跨空隙。

先核对文件的 SHA-256 与响应一致，再从第二行开始按行分批读取（例如每 1000 行一页）；保留每行原始 UTF-8 字节及换行。按 stepId 分组核对记录条数 records、展开格数 count 与原始记录行的 SHA-256（recordsSha256），并检查区域无重叠。这样取完整份清单，不需调用 query 或重新计算。发布文件与报告应一起保留；报告本身存在不能证明工作簿已经发布，仍看 output.saved/version 与批次 publication。

`fileVerification.calculation` 给本次 computedFormulas/computedResults、missingFormulas/missingResults、historicalResults、computedSha256 和 outputVersion；computedSha256 是表序/行/列顺序的 `[sheetId,cell,type,value]` JSON UTF-8 行摘要（每行末尾 LF）。保存核验逐格确认写出值与本次模型一致；这不等于与 Excel 真值一致。不可把保留的 historical 值当本次计算结果，也不能用后一次 query 或 Excel 重算替代同次保存依据。
