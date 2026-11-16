# 工作表与工作簿保护

怎样查清保护状态、怎样在不破坏原有权限的前提下编辑受保护的工作簿、保护带口令时怎么办。参数查 `help sheet.protect`、`help workbook.protect`、`help query.protection`。

## 先查清保护状态

`inspect` 给出保护的概况；要看完整权限，用 `query` 的 `protection`。下例中 `{}` 表示工作簿结构保护，`{"sheet":"Sheet1"}` 表示这张工作表的保护：

```sh
cat > protection.json <<'EOF'
{ "input": "input.xlsx", "protection": [ {}, { "sheet": "Sheet1" }, { "sheet": "Data" } ] }
EOF
./excel-kit query protection.json > protection.out.json
python3 - <<'EOF'
import json
d = json.load(open('protection.out.json'))
if not d.get('ok'):
    raise SystemExit(f"query failed: {d.get('error')}")
for p in d['protection']:
    print(p.get('scope'), p.get('sheet'), p.get('protected'), p.get('password'), p.get('unsupportedAttributes'))
EOF
```

- 结果在响应的 `protection` 数组里，与请求一一对应。每项有 `scope`、`protected`、`password.present` 和 `unsupportedAttributes`。
- 工作表保护的结果另有 `allow`（权限）；工作簿结构保护的结果没有 `allow`，另有 `lockWindows`、`lockRevision` 等。

## 怎样编辑

- **只往未锁定的格里填写**：可以先直接写。失败的步骤不改变任何状态。
- **需要解除无口令保护才能编辑**：用配方 `recipes/edit-protected-workbook`。它在同一个事务里保存原有权限、只解除本次写入涉及的工作表保护（以及新增工作表所需的结构保护）、执行编辑、恢复并核对权限，最后只发布一次；中途任何一步失败都不发布。配方只接受 `write` 和 `sheet.add` 两种操作；`calculation` 省略时用 `none` 做保存维护。在程序包目录里运行（`input` 换成工作簿的实际路径）：

```sh
cat > protected.json <<'EOF'
{
  "input": "input.xlsx",
  "output": "output.xlsx",
  "strict": true,
  "scriptPath": "recipes/edit-protected-workbook/recipe.js",
  "arguments": {
    "operations": [ { "op": "write", "sheet": "Sheet1", "range": "B2", "values": [[12]] } ],
    "calculation": { "mode": "all" }
  }
}
EOF
./excel-kit script protected.json > protected.out.json
python3 - <<'EOF'
import json
d = json.load(open('protected.out.json'))
if not d.get('ok'):
    err = d.get('error') or {}
    raise SystemExit(f"{err.get('code')}: {err.get('message')}")
print(d['output'].get('saved'), d['script']['result']['restored'])
EOF
```

- **自己恢复保护时**，要传回查询得到的**完整 `allow`**：`{"op":"sheet.protect","sheet":"Sheet1","allow":<查询得到的 allow>}`。只写一个不带参数的 `sheet.protect`，会换成新保护的默认权限。解除、编辑、恢复要放在同一步里，这样中途失败也不会留下一个已解除保护的版本。
- **用户要求加保护的**，放在所有填写之后做，免得保护挡住前面的填写（配方 `recipes/fill-template` 就是这样做的）。

## 保护带口令时

不猜测口令，也不绕过校验。请用户在 Excel 里解除相应保护、另存一份副本，再用副本重试。配方遇到需要解除的保护带口令时，会在编辑前失败并给出同样的提示。有扩展的、程序不支持的保护属性（`unsupportedAttributes` 不为空）时，配方同样拒绝。

保护只是存在、本次操作并没有触及它时，不影响本次计算与发布。
