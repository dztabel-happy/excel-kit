# 按条件批量修改记录并读回结果

按列条件数出匹配的记录，只对这一批记录做修改，只重算关心的结果并读回。

## 运行

在本目录下依次执行：

1. `excel-kit script request.json`：生成 `trading-table-updated.xlsx`。

`request.json` 里的路径相对本目录。脚本本体是 `recipe.js`，开头的注释列出了全部参数。

## 参数

`target` 与 `where` 是选择条件；`set` 是要写入的列值；`read` 是要重算并返回的范围；`expectMatches` 可选，用来防止条件写错。

## 样本与许可

`sample/trading-table.xlsx` 是本仓库第三阶段用 XlsxWriter 合成的 Table 样本，可再分发。

## 样本上的结果

数量不小于 2 的两条记录单价改为 9.9：1 月汇总 49.5，2 月汇总 20。

## 要点

- 编辑步骤带 `expectedMatches`，匹配数量与刚才数到的不一致就拒绝执行（`SELECTION_COUNT_MISMATCH`）。它只核对数量，不核对是哪几条；先数和修改在同一个脚本事务里，中间不会有别的写入。
- 修改、只算 `read` 列出的结果、读取在同一步。交付件里只保证这些结果是当前值；其余受这次修改影响的公式没有值，保存时有 `FORMULAS_WITHOUT_VALUES` 警告。
- 没有匹配是正常结果：不写文件，`output.saved:false`。

## 结果算不完整时

依次运行 `excel-kit run prepare-incomplete.json`、`excel-kit script incomplete.json`：前者在样本的月汇总 D2 写入未知函数，后者仍修改两条记录并发布 `incomplete-updated.xlsx`。D2 不写占位错误值，返回结果为 null；这表示不可得，不能当成计算出的空值。`script.result.unavailable` 给出计数、示例、stepId，完整清单由顶层 `calculationReport.path` 指向。按行读取 JSONL：第一行是步骤版本、计算条件、数量和摘要，后续是逐格压缩记录；核对记录数、展开格数及摘要。

需要旧行为时运行 `excel-kit script incomplete-strict.json`，它设置 `requireCompleteCalculation:true`，被分组 wb.run 继承；本例会回滚且不写文件。`strict:true` 只让失败调用抛错，不会把默认成功的 incomplete 改成失败。正常配方也可按需在请求顶层加此发布选项。
