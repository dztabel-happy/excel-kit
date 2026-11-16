---
name: excel-kit
description: 需要在本地读取、编辑、重算并保存 XLSX／XLSM 工作簿，且已安装 excel-kit 命令或解压程序包时使用。
---

# excel-kit

npm 安装后运行 `excel-kit <命令>`。以下命令以此为例；ZIP 用户在包目录用 `./excel-kit`（Windows 是 `.\excel-kit.exe`）替代命令名。`run`、`query` 这类命令读一份 JSON 请求文件；每个响应在标准输出上占一行 JSON，大的输出先存文件再解析取字段，按行用 `head` 截不短它。先 `excel-kit inspect input.xlsx` 看有哪些工作表与对象，再把要做的事写成一份请求：

```sh
cat > edit.json <<'EOF'
{"input":"input.xlsx","output":"output.xlsx",
 "operations":[{"op":"write","sheet":"Sheet1","range":"B2","values":[[1200]]}],
 "read":[{"sheet":"Sheet1","range":"B10","form":"compact"}]}
EOF
excel-kit run edit.json > edit.out.json
```

只读不改用 `query`：请求里同样写 `read`，没有 `output`。输出旁的 `.excel-kit-*.lock.guard.sqlite` 是写入锁留下的，不是交付物，也不要删。操作名与参数查 `excel-kit help <名字>`。动手前要知道的是下面八条；其余按文末的表读一篇。

## 契约

### 1. `run` 是事务

读入 `input`、执行 `operations`、计算、写出并读回核对、原子替换 `output`，任何一步失败都不动 `output`。`output` 已存在就拒绝，除非在 `overwrite.expectedVersion` 给出它当前的版本；核对的是写出的文件与内存一致，不是与 Excel 一致。

### 2. 四种调用方式

默认用一次 `run`，改动、计算、读取写进同一份请求。其余三种：

- `batch`：多个写好的步骤各自成败。默认一步失败就都不发布；`commit:"successful"` 允许只发布成功的步骤，这时返回 `ok:false`、`status:"partial"`。
- `script`：下一步取决于上一步的结果。正常返回后尝试发布成功的步骤；失败的 `wb.run` 只是没生效，不影响外层的 `ok`，要看 `script.stepFailures`；抛错、超时等才整体回滚。`strict:true` 让失败的 `wb.*` 抛错：不捕获就整体回滚并带出原错误，捕获了就只是那次没生效。
- `session`：跨调用保留工作簿，可另存多份。`session.run` 只改内存，失败的一步不留痕迹；`session.save` 才写文件，没保存就结束则修改丢弃。

批次和脚本有没有写文件，看 `output.saved`。

### 3. 值的来源

紧凑读取（`"form":"compact"`，数据在 `rows`，`cells` 是空数组）里公式格是 `{"f","v","s"}`。`s` 为 `result` 才是本次请求算出的当前值；`cache` 是文件里存着的上次结果；`invalidated` 是被本次修改作废的存值，`v` 仍是旧数；`unavailable` 是没有值。`result` 只出现在算它的那次请求里：刚算好并保存的值，下一次不计算的读取里就是 `cache`。

### 4. 计算方式

- `all`：尝试整本计算；不完整也发布可得结果，`ok:true`、`calculation.status:"incomplete"`，并给 `CALCULATION_INCOMPLETE` 警告。要旧的拒绝行为，在请求顶层写 `requireCompleteCalculation:true`。
- `targets`：算列出的格及其全部上游，只保证这些格。其余受本次修改影响的普通公式在文件里没有值（警告 `FORMULAS_WITHOUT_VALUES`），不重算的读取方看到的就是没有值。
- `none`：不做最后的重算，受影响的普通公式同上；操作自己要当前值时仍会就地算一小块。

普通公式缺值保存时，`auto/autoNoTable` 自动设 `fullCalcOnLoad:true`（覆盖原 false/缺省）；手动模式保留并警告。模拟运算表失效缓存仍保留，`autoNoTable` 仍需人工重算。完整不可得清单在发布时给出的 JSONL 文件，读法见 [deliverable.md](reference/deliverable.md)。

省略 `calculation`：`run`、`preflight`、`audit` 是 `all`；批次步骤、`session.run`、`wb.run` 沿用批次根、`session.open`、脚本请求上的那个；读取和查询不沿用，读取默认 `none`。

### 5. 取值与判失败

先看 `ok`；失败的响应不保证有成功时的字段。

- `run`、`preflight`、`read`、`query`：错误在顶层 `error`。读取在顶层 `ranges`，其他查询在同名键。读取算不完整时 `ok:false` 而没有 `error`，看 `calculation`。
- `batch`：每步在 `batch.steps[i]`，读取在其 `result.ranges`，错误在其 `error`；整批中止或执行前被拒才有顶层 `error`，后者没有 `batch`。
- `script`：返回值在 `script.result`，错误在顶层 `error`。
- 会话：每行应答的 `result.response` 才是业务响应，错误在其 `error`；`session.run` 的读取在其 `result.ranges`。

协议层成功不等于做成了：JSON-RPC 应答有 `result`、退出码 0 只说明请求被处理，它顶层的 `error` 只报协议错误。`ok:true` 也要看：保存步骤和 `query` 的 `calculation.status` 都可以是 `incomplete`；算完整了、`s` 是 `result` 的值也可以是 `{"e":"#DIV/0!"}` 等错误值。

### 6. 什么贵

贵的是步数与整本重算：每个写步骤都把整本写出再读回核对，每次读取、查询都重新解析整本；只要少数结果用 `targets`，中间步骤用 `none`。默认期限 30 秒（`timeoutMs`，上限 300000）。硬上限在 `capabilities` 的 `limits`。

### 7. 怎么查

- `help` 是索引；`help <名字>` 查一项，主题、操作、原生编辑用本名，其余带类别前缀，如 `query.<种类>`、`session.<方法>`、`wb.<成员>`、`fn.<函数名>`；`help all <主题>` 给全文；`help find <词…>` 按名字和描述找。
- `capabilities`：程序认得的全部名字，不保证参数都能算。
- `inspect <文件>`：工作簿概况与文件版本，不计算；其余对象用 `query` 查。
- `preflight`：把同一份 `run` 请求在内存里完整做一遍而不写文件。它查输出目录在不在、`output` 是否已存在；不占写入锁，不查能否写入，不保证之后没人改文件。

### 8. 选择与数量断言

`query` 的 `select` 与 `selection.edit` 用同一套条件。`expectedMatches` 只核对个数：个数相符不证明选中的是你要的对象，改完看 `selections[i].targets`。条件涉及公式的值要写 `source:"calculation"`，否则公式的值算未知，整个条件因此判不出时，查询标 `status:"incomplete"`，`selection.edit` 整步拒绝。文本的值比较区分大小写，数字与文本不互转（`"001"` 不等于 `1`）；写错不报错，只是选错。

## 下一跳

参数与字段只在 `help` 里。`reference/` 下每篇只讲一个主题里推不出来的语义，按情形读一篇：

| 情形 | 读 |
|---|---|
| 读单元格、在代码里取值 | `reference/reading.md` |
| 一页读不完、要续页（`READ_TOO_LARGE`、`VERSION_REQUIRED`） | `reference/paging.md` |
| 写请求、覆盖文件（`OUTPUT_EXISTS`）、预检、批次、`FILE_BUSY` | `reference/editing-and-saving.md` |
| 按条件选出或批量修改 | `reference/selection.md` |
| 决定算多少、读懂计算范围 | `reference/calculation.md` |
| 核对交付件；公式没有值（`FORMULAS_WITHOUT_VALUES`）；要 Excel 打开时重算 | `reference/deliverable.md` |
| 慢、超时（`*_TIMEOUT`）、工作簿大 | `reference/cost.md` |
| 计算不完整警告或严格拒绝（`CALCULATION_INCOMPLETE`），要找源头与影响范围 | `reference/audit.md` |
| 宏、加密、模拟运算表等不支持项，怎样向用户说明 | `reference/unsupported.md` |
| 某个函数算不出来或与 Excel 不同 | `reference/function-limits.md` |
| 不认识的错误码、错误里的 `suggestions` | `reference/errors.md` |
| 下一步取决于上一步的结果 | `reference/scripts.md` |
| 边看边改、一次打开交付多份 | `reference/sessions.md`，可运行的驱动在 `reference/session-loop.md` |
| 工作表或工作簿受保护 | `reference/protection.md` |
| 有透视表或切片器 | `reference/pivots.md` |
| 公式引用别的工作簿 | `reference/external-links.md` |
| 自己实现会话的客户端 | `reference/api.md` |

现成的脚本在 `recipes/README.md`，验证范围在 `VALIDATION.md`。
