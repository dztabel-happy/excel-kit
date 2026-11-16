# 巡检公式问题并应用选定的修正

对整本工作簿做计算审计，找出失效引用、未知函数、循环引用和公式错误，区分问题的源头和只是受累的单元格；给出修正后，只有第二次审计干净且声明的业务规则全部成立才发布。

## 运行

在本目录下依次执行：

1. `excel-kit script inspect.json`：只巡检，返回候选清单，不写文件。
2. `excel-kit script request.json`：应用 `fixes`，生成 `model-repaired.xlsx`。

`request.json` 里的路径相对本目录。脚本本体是 `recipe.js`，开头的注释列出了全部参数。

## 参数

`fixes` 是选定的修正（`sheet`、`cell`、`formula`）；针对审计没有标记的单元格的修正要标 `businessRule:true`；`assertions` 是修复后必须成立的业务断言。

## 样本与许可

`sample/model-to-check.xlsx` 与 `sample/clean-control.xlsx` 来自本仓库第二阶段模型检查证据的归档 `docs/goals/phase-03-workbook-maintenance/evidence/model-audit/raw.tar.gz`（成员 `model-audit/old-task-02/待检查 模型.xlsx` 与 `正常 对照.xlsx`），由 `scripts/validation/audit.py` 的独立写入器合成，属仓库自生成，可再分发。SHA-256 分别为 `31d49d5a…eda2d` 与 `5b56b3a6…aa4f`。`sample/expected.json` 是从该脚本的常量手工抄出的注入清单与修正值。

## 样本上的结果

巡检指出 5 个源头（B3 失效引用、B4 与 B5 循环、B6 未知函数、B10 除零）和 5 个受累单元格；修复后审计为零，三条业务断言通过，公式与正常对照逐格相同。

## 要点

- 全部修正放在同一步：循环或未知函数还在时整本无法完成计算，逐个修会每一步都失败。
- B1、B2 两处业务错误能正常计算，任何审计都看不出来，只能靠声明的业务规则发现。
- 少修一处循环，配方拒绝发布。
