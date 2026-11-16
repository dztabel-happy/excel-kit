# 逐条读应答的会话驱动

下一步要看上一条应答才能定、又想让工作簿一直留在内存里时，自己驱动 `excel-kit session`：发一行，读一行。下面的例子顺带做了一件常见的事：同一个模板改一个输入，交付多份。会话的规则见 [sessions.md](sessions.md)。

## 例

```sh
python3 - <<'EOF'
import json, subprocess
from pathlib import Path

inputs = [100, 200, 300]                       # 每份交付改成的输入值
sheet, input_cell, target = "Sheet1", "B2", "B10"
out = Path("deliveries"); out.mkdir()
p = subprocess.Popen(["./excel-kit", "session"], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
count = 0

def call(method, **params):
    global count
    count += 1
    p.stdin.write(json.dumps({"jsonrpc": "2.0", "id": count, "method": method,
                              "params": {"sessionId": "s", **params}}) + "\n")
    p.stdin.flush()
    reply = json.loads(p.stdout.readline())
    if "error" in reply:                       # 协议层：这一行没有被当作请求执行
        raise RuntimeError(reply["error"])
    response = reply["result"]["response"]     # 业务响应
    if not response["ok"]:
        raise RuntimeError(response.get("error") or response)
    return response

try:
    call("session.open", input="input.xlsx", output=str(out / "first.xlsx"), calculation={"mode": "none"})
    call("session.run", operations=[{"op": "workbook.settings",
         "calculation": {"mode": "auto", "fullCalcOnLoad": True}}])
    for value in inputs:
        step = call("session.run",
                    operations=[{"op": "write", "sheet": sheet, "range": input_cell, "values": [[value]]}],
                    calculation={"mode": "targets", "targets": [{"sheet": sheet, "range": target}]},
                    read=[{"sheet": sheet, "range": target, "form": "compact"}])
        cell = step["result"]["ranges"][0]["rows"][0][0]
        if cell["s"] != "result":
            raise RuntimeError(cell)
        saved = call("session.save", output=str(out / f"{value}.xlsx"))
        if not saved["output"]["saved"]:
            raise RuntimeError(saved)
        print(value, cell["v"], saved.get("warnings"))
    call("session.close")
finally:
    p.stdin.close()
    p.wait()
EOF
```

## 改成自己的任务时

- `call` 的两层判断都要留：先看应答顶层的协议 `error`，再看业务响应的 `ok`。一次只发一条、读一条，应答就与请求一一对应。失败就停；已经保存的文件仍在。
- 这是连续修改同一个会话：另存之后会话绑定刚保存的文件，下一轮从它继续，早先的交付件不被改写。不要每一份重新打开模板，保存之前也不需要再发一次 `all`。
- 打开时用 `none`，所以设置那一步不计算；每一份把写输入、只算目标、读取放在同一步。读取范围不等于计算范围：目标之外的格没有本次结果。
- 这样交付的文件里，目标之外受影响的公式没有当前值（保存时有 `FORMULAS_WITHOUT_VALUES`，步骤的 `calculation.scope.omittedFormulas` 是没算的个数），自动模式下缺值保存会设上打开重算，手动模式则警告，见 [deliverable.md](deliverable.md)。要尝试整本当前值可用 `all`，仍须检查不完整警告与清单。
- 模板里有 `CELL` 公式时不能这样换文件名另存，有模拟运算表时目标要避开依赖它的格，分别见 [sessions.md](sessions.md)、[deliverable.md](deliverable.md)。
- 大模型上给 `session.open` 带足 `timeoutMs`，各步沿用它。
