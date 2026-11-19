# 错误码

先看 `error.message` 与 `error.details`，用 `help find <词>` 查相关帮助。

## `details.suggestions` 的读法

超时与读取超限的错误在 `details.suggestions` 里给出候选的改法，`details.hints` 写明字段该放在哪里。

- 按 `condition` 判断适用条件，按 `path` 合并局部修改 `set`，保留输入、编辑、版本与计算环境。`request` 可作为独立只读请求。
- 简写的步骤先把原操作包成 `operations:[原操作]`（原生编辑是 `nativeEdits`）；旧式查询先删 `calculate` 再设 `calculation`。
- 单项限 1024 字符，合计 4096；截断见 `details.suggestionsTruncated`。脚本错误路径见末节。详情过大时有 `detailsOmitted:true`。

## 请求与读取

| 码 | 是什么 | 怎么办 |
|---|---|---|
| `INVALID_REQUEST` | 请求结构不对，或字段放错了位置 | `details.path` 指出出错的对象；有 `details.hints` 就照它写明的位置改。单个操作 `{op:…}` 旁边写了 `calculation`、`read` 时，把操作放进 `operations` |
| `UNKNOWN_HELP_TOPIC` | `help` 的名字不对 | 从 `details.suggestions` 里挑一个，或 `help find <词>` |
| `READ_TOO_LARGE` | 一个范围这一页超过上限 | `paging.md` |
| `VERSION_REQUIRED` | 续页没带版本，或会话、脚本里没有登记过的续页 | 同上 |
| `SOURCE_CHANGED` | 续页之前有过成功的步骤；或输入文件在读入之后被改了 | 续页：从偏移 0 重读。输入被改：重新打开重做 |
| `CALCULATION_INCOMPLETE` | 默认是计算不完整的警告，仍发布；`requireCompleteCalculation:true` 时是拒绝发布的错误。`strict` 脚本里算不完整的读取仍以它抛出 | 看 `details` 的诊断，`audit.md` |

## 计算

| 码 | 是什么 | 怎么办 |
|---|---|---|
| `UNSUPPORTED_FUNCTION` | 函数没有计算入口，或是 VBA 自定义函数、Excel 4 宏表函数 | `help fn.<函数名>`；`unsupported.md` |
| `EXTERNAL_REFERENCE` | 依赖外部工作簿 | `external-links.md` |
| `CIRCULAR_REFERENCE` | 真实循环；不支持迭代计算 | 报告，或按业务改公式，多个循环在同一步里改 |
| `UNSUPPORTED_UNION_REFERENCE` | 联合多区域引用用在了不支持的位置 | 该公式不可得，报告 |
| `ARRAY_MODE_REQUIRED` | 公式的数组模式不受支持；Excel 去掉显式交集后保存的旧式公式也会触发 | `function-limits.md`；存值不能代替本次结果 |
| `PIVOT_REFRESH_REQUIRED` | 透视的来源变了，却没有在同一步刷新 | 改来源的同一步里 `pivot.refresh`，`pivots.md` |
| `PIVOT_RESULT_UNAVAILABLE` | 公式引用的透视缓存设了打开时刷新，或同一步里刷新过、之后来源又变了；或另一张透视以未刷新的透视结果为来源 | 在同一步刷新该透视，用该步自带的 `read` 读；打开时刷新的透视请用户在 Excel 里刷新，`pivots.md` |
| `PIVOT_NOT_REFRESHED` | 只保留、不支持编辑的透视，来源变了 | 不当作本次结果；请用户在 Excel 里刷新 |
| `CALCULATION_TIMEOUT`、`STEP_TIMEOUT`、`BATCH_TIMEOUT` | 计算、整步、整批超过期限 | `cost.md` |
| `…_CALCULATION_INCOMPLETE`（`RANGE`、`DATA`、`FILTER`、`PIVOT`、`AUTOFIT`）、`…_FORMULA_VALUES_REQUIRED`（`FILTER`、`DATA`） | 操作自己要用的当前值算不出来；或要用公式的值而没写 `formulaValues:"calculate"` | `calculation.md` |
| `RULE_CALCULATION_REQUIRED`、`SELECTION_CALCULATION_REQUIRED` | 要当前值的校验、规则结果或选择，配了 `"mode":"none"` | 去掉 `none` |
| `FORMULAS_WITHOUT_VALUES`（警告） | 发布的文件里有公式没有值；与计算不可得数量不同 | `deliverable.md` |
| `MANUAL_RECALCULATION_REQUIRED`（警告） | 手动模式缺值公式，或手动/autoNoTable 的失效模拟运算表须人工重算 | `deliverable.md` |

## 编辑与保存

| 码 | 是什么 | 怎么办 |
|---|---|---|
| `OUTPUT_EXISTS`、`TARGET_CHANGED` | 目标已存在，或在读入之后被改了 | 换输出名；确认覆盖就把 `details.currentVersion` 填进 `overwrite.expectedVersion` |
| `INVALID_OUTPUT`、`OUTPUT_UNAVAILABLE` | 输出的扩展名不对；输出目录不存在或路径不可用 | 前者改扩展名；后者先建目录，或按 `details` 改路径 |
| `WORKBOOK_FORMAT_MISMATCH` | 扩展名与工作簿类型不符 | `.xlsm` 仍存为 `.xlsm` |
| `FORMULA_TOO_LONG` | 写入的公式换成原生存储形式后超过 8192 字符 | 缩短实际写入的公式；`editing-and-saving.md` |
| `FILE_BUSY`、`LOCK_UNAVAILABLE` | 另一个写入者持有锁，或遗留的锁记录无法确认；取不到锁（如目录不可写） | 看 `details`；不要为了解锁删 guard 文件 |
| `FILE_IDENTITY_MISMATCH` | 含 `CELL` 公式的会话另存到别的文件名 | `sessions.md` |
| `SELECTION_COUNT_MISMATCH`、`SELECTION_INCOMPLETE` | 命中个数与 `expectedMatches` 不符；条件判不出 | `selection.md` |
| `EXTERNAL_LINK_IN_USE`、`EXTERNAL_LINK_RENUMBERING` | 外部链接还在用，或删除会改变别的链接的序号 | `external-links.md` |
| `NO_CHANGES` | 不是错误：`session.save` 成功响应里的 `status`，没有写文件 | `sessions.md` |
| `SESSION_EXISTS` | 会话 id 已被占用 | 换一个 `sessionId` |
| `CANCELLED` | 请求被取消 | 这一步不改状态；脚本整体回滚 |

## 脚本与会话

| 码 | 是什么 | 怎么办 |
|---|---|---|
| `SCRIPT_FAILED` | 脚本抛错、有语法错误或线程退出，整体回滚 | 原错误在 `error.details.cause`（`strict`）或 `error.details.lastCallFailure`，不是每次都有；`scripts.md` |
| `SCRIPT_TIMEOUT`、`SCRIPT_STEP_LIMIT`、`SCRIPT_RESULT_NOT_SERIALIZABLE` | 整段超时；`wb.run` 次数超限；返回值不能序列化。都整体回滚 | 同上 |

会话应答顶层的 `error`（`-32700` 等）只表示协议问题，不在 `result.response` 里；各码的含义见 `api.md`。
