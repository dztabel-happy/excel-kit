# 分页与续页

一个读取范围一页读不完时怎样续、续页要绑定哪个版本、`READ_TOO_LARGE` 之后怎样改。读取的入口与形状见 [reading.md](reading.md)，字段查 `help read`。

## 一页多少、怎样续

- 一个范围一页至多 10,000 格（行数乘列数，`limits.readCells`），超过报 `READ_TOO_LARGE`。`offsetRows`、`limitRows` 写在各个范围里。
- 每个范围的结果里有 `page.nextOffsetRows`：为 `null` 才是这一项读完了。一份请求读多个范围时各管各的：第 0 项读完不代表第 1 项读完。续页保留整个范围数组与顺序，只改那一项的 `offsetRows`。

## 续页的版本

续页要绑定第一页读到的版本，保证各页来自同一份工作簿。

- **一次性命令**：第二页起必须在请求**外层**带 `expectedInputVersion`，值是第一页响应的 `workbookVersion`，否则报 `VERSION_REQUIRED`。命令行 `read` 用 `--expected-version`。
- **会话与脚本**：全部偏移为 0 的成功读取登记"页身份"，即去掉偏移与显式版本之后的请求，含 `limitRows` 与计算参数。同一身份的续页可以不带版本，自动沿用第一页的。
  - 登记之后只要有一个步骤成功（不含编辑的、只做计算的也算），自动续页报 `SOURCE_CHANGED`，`details` 里有 `boundVersion` 与 `currentVersion`；从偏移 0 重读就重新登记。
  - 显式的 `expectedInputVersion` 优先，但仍须与当前字节一致。没有登记过的续页报 `VERSION_REQUIRED`。
  - 脚本回滚时登记一起回滚。

## `READ_TOO_LARGE` 之后

错误的 `message` 给出实际列宽与每页能放的行数，`details.suggestions` 给出改法（读法见 [errors.md](errors.md)）：

- 独立的单范围只读请求（`read`、`query`、`session.read`、`wb.read`）：候选里的小 `request` 可以照发，它就是第一页。
- `run`、`preflight` 与步骤：候选只给 `$.read[i]`、`$.steps[i].read[j]` 的 `set`，用它替换那一个读取项，其余编辑与读取项原样保留。原来的非零偏移也从 0 重读。
- 多个范围同时超限：每次只报第一项。逐项改到第一页成功，再各自按 `nextOffsetRows` 读到 `null`。
- 一行就超过上限：第一页缩成 10,000 列，其余的列块配 `limitRows:1`。纯读取也可以改用 `query` 的 `cells`：结果在 `cells[].items`，按 `nextOffset` 续，同样绑定第一页的版本。
- 带编辑的请求只重试一次。之后要续读，用 `query` 读已发布的输出，或用会话的 `read` 读已提交的状态，都从偏移 0 起。

## 批次与预检里的页

- 批次步骤里读到的页不能与发布后的页拼接：后面的步骤可能改了同一个范围。发布成功后对最终输出从偏移 0 重读，之后各页绑定这一页的 `workbookVersion`。
- `"preflight":true` 的批次不发布，没有可续读的输出。单独的 `preflight` 不写文件，续页绑定它响应里的 `inputVersion`。

## 例

跟着 `nextOffsetRows` 读完一个范围。偏移写在范围里，版本写在外层。

```sh
python3 - <<'EOF'
import json, subprocess
request = {"input": "input.xlsx", "read": [
    {"sheet": "Sheet1", "range": "A1:D200", "form": "compact", "offsetRows": 0, "limitRows": 80}]}
rows = []
while True:
    json.dump(request, open("page.json", "w"))
    out = subprocess.run(["./excel-kit", "query", "page.json"], capture_output=True, text=True).stdout
    page = json.loads(out)
    if not page["ok"]:
        raise SystemExit(f"failed: {json.dumps(page.get('error') or page.get('calculation'))}")
    rows += page["ranges"][0]["rows"]
    offset = page["ranges"][0]["page"]["nextOffsetRows"]
    if offset is None:
        break
    request["read"][0]["offsetRows"] = offset
    request["expectedInputVersion"] = page["workbookVersion"]
print(len(rows))
EOF
```

## 找两个版本之间的内容变化

用 `help compare`。比较单元格输入与展开后的公式，先缩到可疑的同名表与矩形。首批响应给出两份文件版本，续页同时传回两者。若只改名定义、格式或对象，应改查相应对象；若同值公式变常量，比较会直接列出两侧类型。
