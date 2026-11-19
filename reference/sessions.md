# 会话

会话怎样驱动、保存与另存的规则、哪些默认值与一次性命令不同。会话保证什么见 `SKILL.md` 第 2、5 条；各方法的参数查 `help session.<方法>`；协议的细节在 [api.md](api.md)；逐条读应答的驱动见 [session-loop.md](session-loop.md)。

## 两种驱动方式

- **整段管入**：`session.open` 时自己定 `sessionId`，后面的请求直接用它，整段对话一次送进标准输入。后面的请求看不到前面的应答，不能据此分支。
- **逐条读应答**：保持管道打开，发一行读一行，根据应答决定下一步。
- 执行环境不向子进程传标准输入时，会话收不到任何请求，也没有输出：改用 `run` 或 `script`。

## 例：整段管入

```sh
./excel-kit session <<'EOF'
{"jsonrpc":"2.0","id":1,"method":"session.open","params":{"sessionId":"s","input":"input.xlsx","output":"output.xlsx","calculation":{"mode":"none"}}}
{"jsonrpc":"2.0","id":2,"method":"session.run","params":{"sessionId":"s","op":"write","sheet":"Sheet1","range":"B2","values":[[1200]]}}
{"jsonrpc":"2.0","id":3,"method":"session.run","params":{"sessionId":"s","operations":[{"op":"write","sheet":"Sheet1","range":"B3","values":[[300]]}],"calculation":{"mode":"all"},"read":[{"sheet":"Sheet1","range":"B10","form":"compact"}]}}
{"jsonrpc":"2.0","id":4,"method":"session.read","params":{"sessionId":"s","ranges":[{"sheet":"Sheet1","range":"B10","form":"compact"}]}}
{"jsonrpc":"2.0","id":5,"method":"session.save","params":{"sessionId":"s"}}
{"jsonrpc":"2.0","id":6,"method":"session.close","params":{"sessionId":"s"}}
EOF
```

- 第 2 行是简写的一步，沿用打开时的 `none`，不计算。第 3 行要计算和读取，所以用完整写法；读到的值在这条应答的 `result.response.result.ranges`，`s` 是 `result`。
- 第 4 行随后再读同一格，值相同，`s` 却是 `cache`：`session.read` 不沿用计算方式，默认不计算。
- 每行应答按 `id` 对应请求，业务响应在 `result.response`。

## 保存与另存

- `session.save` 不带 `output` 保存到当前目标；带 `output` 是另存，扩展名须与当前目标相同。另存成功后会话改为绑定新文件，早先的交付件不再被改写；所以一个会话可以依次交付多份。
- 原地编辑（`input` 与 `output` 是同一个文件）在 `session.open` 时就要给 `overwrite.expectedVersion`。
- 保存时目标已存在或被改过，报 `OUTPUT_EXISTS`、`TARGET_CHANGED`；确认覆盖的写法查 `help session.save`。来源与目标不是同一个文件、且还没保存过时，原始输入变了报 `SOURCE_CHANGED`，只能重新打开。保存被拒不改变会话状态，修正后可以再试。
- 没有可保存的东西时，保存返回 `{"ok":true,"status":"NO_CHANGES","output":{"saved":false}}`，不写文件：保存到当前目标，条件是自打开或上次保存以来没有成功的步骤；另存，条件是自打开以来没有成功的步骤。
- 想把一个没改动的工作簿另存为副本，先发一个空步骤 `session.run {}`。空步骤沿用打开会话时的计算方式；工作簿里有不支持的函数时，把它写成 `{"calculation":{"mode":"none"}}`。
- 工作簿里有 `CELL` 公式时，另存到别的文件名报 `FILE_IDENTITY_MISMATCH`，因为 `CELL` 的结果按当前文件路径计算。先保存到当前目标，再以它为输入、新文件为输出开一个新会话重算。

## 与一次性命令不同的地方

- `session.run` 省略 `calculation`、`timeoutMs` 时沿用 `session.open` 上的。`session.read`、`session.query` 不沿用计算方式，默认不计算；期限 `session.read` 沿用，`session.query` 不沿用（没写就是 30000）。
- 时钟与随机种子在 `session.open` 时固定，整个会话取同一组值；要别的时钟就在 `calculation.clock` 里写。
- `session.read`、`session.query`、`session.inspect` 读的是内存里的当前状态，不访问磁盘。
- 没有 `audit` 方法：`session.query` 加 `"calculation":{"mode":"all"}` 就是它的默认行为。
- 不接受 `structure` 来源。相对路径按 `excel-kit session` 进程启动时的工作目录解析。
- 依赖透视结果的公式按保存的透视结果计算；要反映刷新后的透视，在刷新的那一步里读，见 [pivots.md](pivots.md)。续页的版本在会话里自动绑定，见 [paging.md](paging.md)。

## 结束与取消

标准输入关闭后，会话做完已经收到的请求，丢弃没保存的状态，退出码 0。取消一条请求用通知 `{"jsonrpc":"2.0","method":"$/cancel","params":{"id":<请求 id>}}`。信号、并发与协议错误见 [api.md](api.md)。

## 试算与提交

先记下当前版本。用 `help session.script` 的 `trial:true` 在该版本及其未保存编辑上试算。脚本只返回所需的普通 JSON。应答中的试算结果属于该情景，之后仍可在同一基准上换参数。选定后以原版本为 `expectedSessionVersion` 正常执行，再保存。

需要增量读取时查 `help session.open`，打开时设 `reread:"precise"`。成功编辑后按 `reread.version` 重读列出的范围和图表对象。返回整簿时重新查询整簿。试算恢复后不发变更提示。
