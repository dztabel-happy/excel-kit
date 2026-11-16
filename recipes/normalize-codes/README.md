# 用正则在整本工作簿里规范化编号

用一个正则表达式在整本工作簿（含隐藏表）里把不规范的编号改成规范形式，再用同一个搜索证明没有残留；有残留就不发布。

## 运行

在本目录下依次执行：

1. `excel-kit run prepare.json`：向样本注入不规范编号，生成 `sample/irregular-codes.xlsx`。
2. `excel-kit script request.json`：生成 `normalized-codes.xlsx`。

`request.json` 里的路径相对本目录。脚本本体是 `recipe.js`，开头的注释列出了全部参数。

## 参数

`pattern` 与 `replacement` 是正则和替换式（支持 `$1`）；`caseSensitive` 默认 false；`sheets` 可限定工作表。

## 样本与许可

`sample/trading-table.xlsx` 是本仓库第三阶段合成的 Table 样本，可再分发。样本里的编号本来是规范的，不规范的写法（CSV 里 `SKU-0004-X` 那种形态）由 `prepare.json` 注入。

## 样本上的结果

4 个单元格、5 处替换（有一格的句子里嵌了两个编号）；大小写不同的写法也命中；`SKU-12-X` 位数不符合模式，按模式不动；公式拼出的结果不是存储文本，不在替换范围。

## 要点

- 替换只作用于存储的文本，不改公式结果。
- 规范后的编号仍是文本，前导零保留。
- 再跑一次找不到任何匹配，不写文件。
