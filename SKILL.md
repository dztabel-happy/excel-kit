---
name: excel-kit
description: 用 ExcelKit 处理本地 XLSX／XLSM。适用于 agent 需要了解陌生工作簿、核对两个版本、编辑公式与原生对象、试算方案或交付可继续编辑的 Excel 文件时。
---

# ExcelKit

先用 `excel-kit inspect <文件>` 找工作表、名称、Table、对象及文件版本。把任务拆成要查明的事实、要改的对象和验收用的格或绑定。用 `excel-kit help find <英文关键词>` 找入口，再用精确帮助查字段。一次能说清的修改写进 `run` 的 `operations`，把验收用的 `read` 放在同一次计算里。独立程序包在当前目录用 `./excel-kit`，Windows 用 `.\excel-kit.exe`。

### 1. 确定输入与输出

`preflight` 执行并核对一次拟议的修改，返回已绑定版本的执行请求。需要确认覆盖、迁移或影响范围时使用。一次 `run` 原子发布到 `output`。替换已有文件必须给出 `overwrite.expectedVersion`。[编辑、预检与保存](reference/editing-and-saving.md)

### 2. 选择执行方式

步骤已知，用一次 run。多个分别计成败的步骤用 `batch`。需要根据中间值判断、循环或断言，用 `script` 和 `strict:true`。跨调用积累未保存编辑用 `session`；在当前状态探索多个情景，用 `session.script` 的 `trial:true` 取回结果。以 `output.saved` 确认文件实际发布。[脚本](reference/scripts.md) · [会话、试算与重读](reference/sessions.md)

### 3. 读值时确认来源

紧凑读取的公式格看 `s`。本次计算为 `result`，历史保存值为 `cache`，已失效为 `invalidated`，没有可用值为 `unavailable`。需要解释差异时先看公式和输入，再查上下游。[读取与取值](reference/reading.md) · [双文件比较与续页](reference/paging.md)

### 4. 决定要算多少

中间编辑用 `none`，核对指定指标用 `targets`，交付要求整本当前结果时用 `all` 并检查完整性。缺普通公式值时自动模式可以设置 `fullCalcOnLoad`，手动模式要读警告。[计算范围](reference/calculation.md) · [交付核对](reference/deliverable.md)

### 5. 用响应决定下一步

先看业务响应的 `ok` 和 `error`，会话从 `result.response` 取。再看计算状态、目标格是否为 Excel 错误值，以及步骤是否实际保存。根据错误提供的位置和恢复建议重选范围、补绑定或修改请求。[错误恢复](reference/errors.md) · [审计与依赖](reference/audit.md)

### 6. 控制读取和计算成本

先取一个有代表性的范围。续页绑定版本，将同一轮需要的编辑和读数组合起来。遇到超时先检查计算范围、步骤数与 `timeoutMs`，上限见 capabilities 的 `limits`。[成本与期限](reference/cost.md)

### 7. 查本机能力并组合

`inspect` 给工作簿事实，`capabilities` 给当前程序识别的名称，`help <精确名称>` 给请求契约。按对象用 `help query.<种类>` 查可读取的状态，然后选修改操作，再确定如何验证。例如迁移模型区块前用 `help compare` 定位内容差异，查 `help range.transfer` 了解源绑定和目标覆盖授权，必要时用 `preflight`。涉及某个函数时单查 `help fn.<函数名>`。[跨簿绑定](reference/external-links.md) · [函数边界](reference/function-limits.md)

### 8. 验证选中的对象

批量选择先核对地址、名称或系列身份。`expectedMatches` 只能断言数量；按公式当前值选择时指定 `source:"calculation"`。用版本绑定后的结果复核实际目标，再决定提交。[条件选择](reference/selection.md)

## 按任务深入

| 需要判断的事 | 按需阅读 |
|---|---|
| 指标在分组、总计和刷新之后如何计算 | [透视与切片器](reference/pivots.md) |
| 交付前检查换行文字、图表系列或原生规则 | [版面与对象](reference/presentation.md) |
| 能否编辑受保护的区域，怎样恢复权限 | [保护](reference/protection.md) |
| 写自己的会话客户端 | [协议与调用方式](reference/api.md) · [逐条应答驱动](reference/session-loop.md) |
| 文件含宏、外部来源或不受支持的对象 | [保留与计算边界](reference/unsupported.md) |

程序包的 `recipes/README.md` 索引可运行的组合示例，`VALIDATION.md` 说明该版本的实测范围。
