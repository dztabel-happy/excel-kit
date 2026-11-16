# 找出算不出来的公式与影响范围

整本计算不完整时，怎样列出问题、追到源头、查清一处改动影响到哪里。不支持的原因见 [unsupported.md](unsupported.md)，字段查 `help query.findings`、`help query.dependencies`。

## 问题清单：`findings`

- `query` 加 `"calculation":{"mode":"all"}` 和 `findings`，就是 `audit` 的默认行为。计算不完整时查询仍是 `ok:true`，要自己看 `calculation.status`。
- `findings.items` 不全是算不出来的格：`kind` 为 `diagnostic` 的是诊断（`code` 是原因，如 `UNSUPPORTED_FUNCTION`），`cache_difference` 是本次结果与文件里存值不同的格。按 `kind` 与 `classification` 筛。
- 默认每页 100 条，`limit` 让一页多取；`findings.nextOffset` 不为 `null` 还有下一页。`assessment` 是汇总：已确认的错误、不可得、存值差异各多少。

## 引用关系：`dependencies`

- 查一个格的上游（`"direction":"precedents"`）或下游（`"dependents"`），加 `"transitive":true` 一直追到头。名称会展开，比自己解析公式可靠。
- 结果与请求一一对应。条目用 `sheet` 加 `range` 定位，`range` 可能是一块区域；`nextOffset` 不为 `null` 还有下一页。
- 它只给**静态**引用，也就是公式里写明的。`INDIRECT`、`OFFSET` 这类运行时才定的区域不在其中：结果的 `limitations.count` 大于 0，`items` 里标出是哪些格（如 `RUNTIME_REFERENCES_PENDING`）。这时即使翻完所有页，也不是完整的影响范围；用它组计算目标或说明影响之前先看 `limitations`。
- 要判断一个上游格是输入还是公式，读回它看 `kind`，紧凑形式看有没有 `f`。

## 整本算不完整时

默认 all 计算不完整仍保存可靠结果，先读 calculation.unavailable.report 指向的 JSONL 清单。显式 requireCompleteCalculation:true 的 all 被拒绝、没有写文件时：

1. 读错误 `details` 里的诊断，按格找出不可得的源头：不支持的函数、外部引用、循环、模拟运算表、打开时刷新的透视。要列全就用上面的 `findings`。
2. 用 `dependents` 加 `transitive` 查这些源头影响到的格，先看 `limitations`。
3. 任务要的结果里，不受影响的那些作为 `targets` 计算。这是按交付要求缩小范围，不是绕过失败：受影响的结果仍然没有。
4. 报告里写清哪些是本次结果，哪些没有值或没有更新，各自是什么原因；不要笼统地说"源文件有错"。

修复时，互相牵连的问题（多个循环、未知函数）要在同一步里改，否则中间状态仍算不出来；改完再审计一次。包里的配方 `recipes/audit-and-repair` 是完整的做法。

## 例

一份查询同时要问题清单和两处引用关系：

```sh
cat > audit.json <<'EOF'
{"input":"input.xlsx","calculation":{"mode":"all"},
 "findings":{"limit":1000},
 "dependencies":[{"sheet":"Sheet1","cell":"B10","direction":"precedents","transitive":true},
                 {"sheet":"Data","cell":"C2","direction":"dependents","transitive":true}]}
EOF
./excel-kit query audit.json > audit.out.json || true
python3 - <<'EOF'
import json
d = json.load(open('audit.out.json'))
if not d['ok']:
    raise SystemExit(f"failed: {json.dumps(d['error'])}")
print(d['calculation']['status'], d['assessment'], len(d['findings']['items']), d['findings']['nextOffset'])
for dep in d['dependencies']:
    print(dep['limitations']['count'], dep['nextOffset'], [(i['sheet'], i['range']) for i in dep['items']][:3])
EOF
```
