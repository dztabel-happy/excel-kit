# 填写模板、校验、加打印设置与保护

用数据填充 `{{键}}` 占位符并写入原生数字，让模板自带的数据验证规则在发布前判定结果，通过后再加打印设置和工作表保护。

## 运行

在本目录下依次执行：

1. `excel-kit script request.json`：生成 `filled-form.xlsx`。

`request.json` 里的路径相对本目录。脚本本体是 `recipe.js`，开头的注释列出了全部参数。

## 参数

`data` 是占位符数据（文本）；`numbers` 是要写成原生数字的单元格；`validate` 是需要通过验证规则的单元格；`print` 是打印设置；`protect` 是保护选项，可带 `password`。

## 样本与许可

`sample/template.xlsx` 由 PowerAI Desktop 的测试样本 `cf-hyperlink-image-protection.xlsx`（Apache-2.0，见 `sample/original/LICENSE`）派生：原件不含占位符和验证规则，且带无口令的工作表保护；派生时用公开操作解除保护、新增"申请单"工作表、写入占位符和四条数据验证规则，派生请求是 `sample/derive-template.request.json`。原件里一张块长度不一致的 2×2 PNG 原样保留，配方不触碰它。

## 样本上的结果

8 个占位符分布在 7 个单元格，全部替换；四个被校验的单元格都通过；金额为 12 × 1350.5 = 16206；工作表受保护并允许调整列宽，打印为 A4 纵向、缩放到一页。

## 要点

- 占位符填充只写文本；数字用 `write`，否则数据验证和公式把它当文本。
- 步骤的 `validate` 让工作簿自己的规则判定：不通过则该步失败，什么都不留下。
- 数据里缺少某个键按错误处理（`unresolved:"error"`），不会发布还留着 `{{...}}` 的表单。
- 保护放在最后，避免保护挡住前面的填写。
