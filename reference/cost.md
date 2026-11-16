# 成本、期限与上限

什么贵、怎样少花、四种期限各管什么、超时之后怎样改、有哪些硬上限。计算方式怎么选见 [calculation.md](calculation.md)。

## 什么贵

- 每个写步骤都把整本写出再读回核对；每次读取、查询都重新解析整本；每条一次性命令另付进程启动与读写文件。
- 整本重算是最大的一项。算出的值不带到下一次调用。
- 唯一少解析一遍的情形：会话里连续的 `session.run`，每一步都只含写常量（字符串、数字、布尔、`null`；公式不算）的 `write` 或没有操作，不带 `validate`、原生编辑和结构变化，工作簿不含动态数组。读取、查询、`inspect`、脚本、保存与另存、失败或取消之后要重新开始，所以每步之后都保存的流程没有这项收益。不用配置，响应不变。

## 少调用、少输出、少计算

- 已知的改动放进一份 `run` 或同一步的 `operations`；核对用的读取写进这一步的 `read`。
- 要分支、循环或读很多区域，用一个脚本，不要多次 `run` 或多次命令行 `read`；同一个文件的多块区域放进一次查询的 `read`。
- 只读需要的范围，用紧凑形式；响应重定向到文件再取字段，不要整份打印；`help all` 同理。脚本只返回需要的数据，`"report":"full"` 的体积随步数线性增长。
- 只查公式、结构和引用用 `none`。纯格式或视图的改动可以用 `none`，但要看 `FORMULAS_WITHOUT_VALUES`。
- 只关心少数结果用 `targets`。有动态数组的工作簿上 `targets` 可能比 `all` 更慢、更占内存，这时用 `all` 或 `none`。
- 会话或多步脚本里，不需要中间结果的步骤用 `none`，交付那一步再按要求选 `targets` 或 `all`。
- 大模型上直接带足 `timeoutMs` 去 `run`，不要先 `preflight`。

## 四种期限

期限写在各自请求或步骤的**外层**，不写在 `calculation` 里，也不写在读取范围里。

| 写在哪里 | 管什么 | 超时 |
|---|---|---|
| `run`、`preflight`、`query`、`audit`、`session.read`、`wb.read` 的 `timeoutMs`；命令行 `read` 的 `--timeout-ms` | 计算阶段；`query` 里每个 `select`、每个 `find` 另按同一个数各自计时。不是整份请求的总期限 | `CALCULATION_TIMEOUT`，`run` 不写文件；选择是 `SELECTION_TIMEOUT`，文本查找是 `TEXT_MATCH_TIMEOUT` |
| 批次的 `steps[i]`、`session.run`、`wb.run` 上的 `timeoutMs` | 整步：解析、编辑、计算、写出与核对 | `STEP_TIMEOUT`，这一步不生效 |
| 批次请求顶层的 `timeoutMs` | 整批 | `BATCH_TIMEOUT`，不发布 |
| 脚本的 `scriptTimeoutMs` | 整段脚本 | `SCRIPT_TIMEOUT`，整体回滚 |

- `session.open`、脚本请求顶层的 `timeoutMs` 是其中各步的默认值。`session.read`、`wb.read` 也沿用它；`session.query`、`wb.query` 不沿用，没写就是 30000。
- 会话没有总期限，也没有空闲超时。每步数十秒的大模型上跑长循环，脚本的整段期限可能不够：改用会话，或分成几段脚本、每段以上一段的输出为输入。

## 超时之后

错误的 `details` 写明是哪一种期限（`deadline` 为 `calculation`、`step` 或 `batch`）、这次的 `timeoutMs`、可调的上限 `maxTimeoutMs` 和该改哪个字段（`hints`），并在 `suggestions` 里给出候选。程序不会自动重试，不会替你改计算方式或期限。`suggestions` 怎么用见 [errors.md](errors.md)。

- 批次先改顶层的整批期限，再按需要改各步。
- "只算目标"的候选始终以"如果只需要这些格"为条件：算多少由交付要求定，不由超时定。读取范围少时候选直接给出目标；没有读取范围、超过 100 个或放不下时，要你把示例换成确实需要的目标，至多 100 项。
- 上限 300000 是允许的最大值，不是预计用时。已知是大模型，第一次就把 `timeoutMs` 设大。
- 超过上限的 `timeoutMs` 报 `INVALID_REQUEST`（命令行参数是 `INVALID_ARGUMENT`），`details.path` 指向被拒的字段。
- 有些超时路径不带这些提示，已经到上限时也没有调整的建议；以实际的 `details` 为准。

## 上限与大模型

数值在 `capabilities` 的 `limits`，越界即拒：

| `limits` 的键 | 管什么 | 越界时 |
|---|---|---|
| `editedCells` | 一次请求或一步编辑的格数 | `BATCH_TOO_LARGE` |
| `readCells` | 一个读取范围一页的格数 | `READ_TOO_LARGE`，见 [paging.md](paging.md) |
| `validateCells` | `validate` 每个区域的格数 | `VALIDATION_CHECK_TOO_LARGE` |
| `calculationTargets`、`batchSteps`、`timeoutMs`、`scriptTimeoutMs` | 计算目标个数、批次步数、两种期限的最大值 | `INVALID_REQUEST` |
| `scriptSteps` | 一段脚本里 `wb.run` 的次数 | `SCRIPT_STEP_LIMIT` |
| `scriptBytes`、`scriptArgumentsBytes` | 脚本文本、`arguments` 的字节数 | `SCRIPT_TOO_LARGE`、`SCRIPT_ARGUMENTS_TOO_LARGE` |
| `scriptLogEntries`、`scriptLogBytes` | 脚本日志的条数与总量 | 不报错，`logsTruncated:true` |

`defaultTimeoutMs`、`defaultScriptTimeoutMs` 是两种期限没写时的值。

约百万公式级的整本计算可能超过期限的上限；性能不保证任意工作簿都能在固定预算内算完。
