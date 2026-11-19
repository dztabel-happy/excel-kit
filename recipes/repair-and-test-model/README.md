# 修复模型并选择一个情景

在空目录生成基准、工作副本和来源模板，最终交付 `delivery.xlsx` 与 `report.json`。示例需要 Python 3；所有工作簿读取、修改和计算均由 ExcelKit 执行。

从程序包根目录运行：

```sh
python3 recipes/repair-and-test-model/workflow.py ./model-output --cli ./excel-kit
```

Windows 使用 `py -3` 和 `.\excel-kit.exe`。安装在 PATH 时可省略 `--cli`。

先比较同一矩形中的公式与输入。示例在目标会话已有未保存编辑的情况下，验证替换范围、来源绑定、当前版本，然后迁入含合并标题、费率公式和图表的区块。两个试算分别得到 36 和 54，仅提交 36。脚本再刷新条件透视指标，调高显示不全的标题行，核对汇总值后保存。报告记录了每次选择所依据的内容差异，以及调用者需要重新读取的范围和图表。
