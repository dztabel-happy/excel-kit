# 编辑受保护的模型并恢复原有权限

更新无口令保护模型的值或公式、新增汇总表，保存后恢复原有的工作表权限与工作簿结构保护。原本未保护的工作表保持未保护，新增工作表也不自动加保护。

例如交给 agent：“把这份模型的输入改为 12，新增汇总页引用结果；修改后保留原来的保护权限。”agent 将明确获准的修改放进本配方的 `operations`，程序在同一个事务中保存保护状态、解除所需保护、编辑、恢复并核对，最后只发布一次。

## 运行示例

在本目录执行，`excel-kit` 换成当前程序路径：

```sh
excel-kit run prepare.json
excel-kit script request.json
```

得到 `updated-model.xlsx`：`Inputs!A1=12`，`Inputs!B1` 与 `Summary!B1` 的本次结果均为 24；`Inputs` 仍受保护，原来的选择、列格式、筛选及对象权限保持；结构保护仍生效。`Notes` 与新建的 `Summary` 不受工作表保护。示例源文件由 `prepare.json` 生成，不依赖外部材料。

用于自己的文件时，修改 `request.json` 的 `input`、`output`、`arguments.operations`，保留 `scriptPath`。已知原件摘要时可带 `expectedInputVersion`；已有目标文件沿用公开的 `overwrite.expectedVersion` 规则。

## 参数与失败处理

- `operations` 只接受 `write` 和 `sheet.add`。已有工作表身份、单元格锁定及公式隐藏格式保持；改名、删表、格式维护应使用另外制定的完整任务。
- `calculation` 可指定目标计算，计算不完整默认也接受步骤，须看返回计算状态与诊断文件；要汇总结果全部可得才交付，可在脚本的 wb.run 分组参数上设 requireCompleteCalculation:true；省略时使用 `mode:"none"` 做保存维护，受影响的旧缓存按正常规则清理，不代表整本已计算。
- 配方只解除本次写入涉及的工作表保护，以及新增工作表所需的结构保护。其他工作表保护不变。
- 需要解除的保护带口令时，脚本在编辑前失败，说明“请用户在 Excel 中自行解除相应保护并另存副本，再重试”。本配方不接收或猜测口令；现有公开保护操作仍可由有合法口令的用户明确调用。
- 遇到不支持的扩展保护、编辑失败、整次计算故障或恢复核对失败，配方抛错，源文件及已有目标保持不变，不发布中途解除保护的文件。

恢复时传回查询得到的**完整 `allow`**，不能只写一个无参数 `sheet.protect`，否则会换成新保护的默认权限。工作簿操作仅管理结构锁，原有窗口／修订限制留在原生记录中。`protectedRanges` 及单元格锁定、公式隐藏也由现有保存链保留；配方不修改它们。恢复按权限语义比较，不要求 XML 中 `true/false` 与 `1/0` 的拼写完全一致。

保护操作见 `excel-kit help protection`；脚本什么时候发布见 `SKILL.md` 第 2 条。
