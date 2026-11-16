# 会话协议

给自己实现 `excel-kit session` 客户端的人；只用命令行不需要读这一篇。字段查 `excel-kit help`；调用方式的契约在 `SKILL.md`，会话怎样用见 [sessions.md](sessions.md)。

## 1. 会话协议

`excel-kit session` 在标准输入输出上说换行分隔的 JSON-RPC 2.0：每行一个 JSON 对象，不支持批量数组，标准输出只有协议消息。请求按到达顺序逐个执行，一个进程里的多个会话也是串行的。

```json rpc
{"jsonrpc":"2.0","id":1,"method":"session.open","params":{"sessionId":"s","input":"input.xlsx","output":"output.xlsx"}}
{"jsonrpc":"2.0","id":2,"method":"session.run","params":{"sessionId":"s","op":"write","sheet":"Sheet1","range":"B2","values":[[1200]]}}
{"jsonrpc":"2.0","id":3,"method":"session.save","params":{"sessionId":"s"}}
```

应答有三种样子，按 `id` 对应请求：

```json response
{"jsonrpc":"2.0","id":2,"result":{"session":{"id":"s","version":"<SHA-256>"},"elapsedMs":41.2,"queuedMs":0.3,"response":{"ok":true,"status":"succeeded","inputVersion":"<SHA-256>","outputVersion":"<SHA-256>","result":{}}}}
```

```json response
{"jsonrpc":"2.0","id":2,"result":{"session":{"id":"s","version":"<SHA-256>"},"elapsedMs":12.5,"queuedMs":0.3,"response":{"ok":false,"status":"failed","inputVersion":"<SHA-256>","error":{"code":"SHEET_NOT_FOUND","message":"Worksheet does not exist.","details":{}}}}}
```

```json response
{"jsonrpc":"2.0","id":4,"error":{"code":-32601,"message":"Unknown method."}}
```

- `result.response` 是该方法的业务响应，业务失败也在这里；顶层的 `error` 只报协议错误，这一行没有被当作请求执行。
- `result.session.version` 是当前内存字节的 SHA-256，与读取、查询响应里的 `workbookVersion` 是同一个值。`elapsedMs` 是执行时长，排队时长在 `queuedMs`。
- 应答不保证与请求同序：一行的格式、`id`、方法名、`params` 不对时立即应答，可能排在更早请求的应答之前。

| 协议错误 | 含义 |
|---|---|
| `-32700` | 该行不是合法 JSON；应答的 `id` 是 `null` |
| `-32600` | 不是合法请求：不是对象、缺 `jsonrpc` 或 `method`、缺 `id`，或 `id` 属于一个还没结束的请求 |
| `-32601` | 未知方法 |
| `-32602` | `params` 不是对象；`sessionId` 缺失或不是字符串（`session.open` 自选的要 1–128 个字符） |
| `-32001` | 未知会话：没有打开过，或已经关闭。轮到这条请求执行时才判，所以打开与后续调用可以一起管入 |

`id` 只接受字符串或数字，`1` 与 `"1"` 是不同的 id。像 JSON-RPC，不像的地方是除 `$/cancel` 外每条消息都必须带 `id`：缺 `id` 的消息得到 `-32600`，不会被静默执行。

取消用通知 `{"jsonrpc":"2.0","method":"$/cancel","params":{"id":<请求 id>}}`：排队中的请求立即以 `CANCELLED` 应答，运行中的步骤以 `CANCELLED` 失败且不改状态，运行中的脚本整体回滚；未知或已结束的 id 被忽略。

结束：

- 标准输入关闭后，已收到的请求全部执行并应答，然后丢弃未保存的状态，退出码 0。
- SIGINT、SIGTERM：停止读入，取消运行中的调用，排队中的请求逐条以 `CANCELLED` 应答，退出码 130。Windows 上没有 SIGTERM，用 `$/cancel` 取消、用关闭标准输入结束。
- 标准输出管道断开：中止运行中的调用，退出码 1。

一个进程可以打开多个会话，每个持有一份工作簿字节，不设数量上限；要并行就启动多个进程。两个会话指向同一个目标文件时，保存时的版本核对保证后保存的一方不会悄悄覆盖先保存的。会话不是守护进程：没有端口、鉴权和后台常驻，生命周期就是这个子进程。

## 2. 退出码

响应 `ok:false` 是 1，其中 `error.code` 为 `CANCELLED` 是 130，其余是 0。所以 0 不等于做成了：一次性脚本有失败步骤、甚至一步都没成功时是 0；会话在输入结束后退出也是 0，不论其中有多少业务失败。

## 3. Node.js 与其他客户端

本发行包提供独立命令行程序，不导出内核库。使用子进程调用 CLI，或通过标准输入输出连接上述会话协议。Node.js 可用 `child_process.spawn`；参数用数组传入，不需要 shell。
