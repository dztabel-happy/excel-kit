# 清点图片与附件并提取核对

列出工作簿里每张图片和每个嵌入附件的大小与 SHA-256，并为每一项返回一份可直接执行的提取请求；提取由 `image-extract` 与 `attachment-extract` 命令完成，提取结果的 `output.version` 应等于清单里的 `sha256`。

## 运行

在本目录下依次执行：

1. `excel-kit run prepare.json`：向附件样本插入两张图片，生成 `sample/with-media.xlsx`。
2. `excel-kit script request.json`：返回清单；脚本只读，不写文件。
3. 先建 `extracted` 目录，再把清单里每一项的 `extract.request` 存成文件，用 `extract.command` 指明的命令执行。

`request.json` 里的路径相对本目录。脚本本体是 `recipe.js`，开头的注释列出了全部参数。

## 参数

`workbook` 是提取命令要读的工作簿路径；`folder` 是提取文件的目录。

## 样本与许可

`sample/WithEmbeded.xlsx` 原样取自 Apache POI 的 `test-data/spreadsheet`（Apache-2.0，见 `sample/LICENSE` 与 `sample/NOTICE`）。它不含浮动图片，两张图片由 `prepare.json` 插入：`sample/green.jpg` 是本仓库用 Pillow 生成的纯色 JPEG，`sample/blue.png` 是本仓库用 Python 标准库生成的纯色 PNG，均无外部素材。

## 样本上的结果

2 张图片（JPEG 637 字节、PNG 78 字节）和 2 个附件（一个 32,768 字节的 XLS 复合文档、一个 1,693 字节的 Ole10Native 压缩包）；提取出的文件摘要与清单逐项相等。

## 要点

- 脚本除工作簿外不能写文件，所以清单与提取分两段：脚本给出请求，一次性命令执行。
- 附件里存的文件名只是数据，不当作路径：输出名由工作表序号和对象 id 生成，只保留短的字母数字扩展名。
- 只读脚本仍要在请求里写 `output`，但不会写出任何文件。
