# 导入 CSV、清洗、去重并追加到 Table

把多个 CSV 按类型导入暂存表，在脚本里清洗与映射字段，剔除不完整记录，对文件内和 Table 内已有的记录去重，追加到带活动筛选的 Table，重应用筛选，最后用独立求和核对可见小计。

## 运行

在本目录下依次执行：

1. `excel-kit run prepare.json`：给样本 Table 加上筛选条件（数量不小于 2）和一个 SUBTOTAL 单元格，生成 `sample/filtered-table.xlsx`。
2. `excel-kit script request.json`：生成 `table-with-imports.xlsx`。

`request.json` 里的路径相对本目录。脚本本体是 `recipe.js`，开头的注释列出了全部参数。

## 参数

`sources` 是 CSV 文件；`columns` 是导入时的列类型；`map` 把 CSV 列映射到 Table 列，可带 `trim`、`pattern`、`replacement`；`key` 是去重键；`required` 是清洗后不能为空的列；`check` 指向 SUBTOTAL 单元格和它汇总的列。

## 样本与许可

三个样本文件都是本仓库第三阶段合成的，可再分发。样本 Table 原本没有筛选和 SUBTOTAL，由 `prepare.json` 预置。

## 样本上的结果

5 条 CSV 记录里追加 0004 和 0005；文件内重复 1 条、Table 内已有 1 条、缺单价被拒 1 条；追加后可见 4 行，可见小计 95.75。

## 要点

- 脚本不能直接读文件。CSV 先用 `data` 操作按类型导入暂存表，脚本读回后在代码里清洗，最后在同一步里删除暂存表。
- 日期和数字由导入按声明的格式解析，脚本不猜格式。
- 向带活动筛选的 Table 追加，必须在同一步里 `filter.reapply`。
