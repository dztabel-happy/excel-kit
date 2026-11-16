# 算多少

按交付要求选计算方式、读懂响应里的计算范围、各调用省略 `calculation` 时算什么、`none` 下哪些操作仍会算。三种方式各保证什么见 `SKILL.md` 第 4 条；算完之后文件里存的是什么见 [deliverable.md](deliverable.md)；期限见 [cost.md](cost.md)；`calculation` 对象的字段查 `help calculation.policy`。

## 三种"没算出来"

- Excel 的错误值（如 `#NUM!`）是计算所得，可以按 `IFERROR` 的规则处理。
- 计算不可得是程序没有本次结果，`IFERROR` 不能把它变成可信的零或别的数。
- 保存被拒是没有发布输出，未必是公式算不了（例如公式过长）。

先看计算状态与值的来源，再看保存结果。

## 按交付要求选方式

- 只要少数本次结果，接受其余公式没有当前值：`targets`。把写输入、算目标、读结果放在同一步。
- 交付件要全部公式都是当前值、下游读者不会重算却要用全部结果、要整本数值审计：`all`。算不完整就报告阻断原因。
- 只查公式与结构：`none`。
- 选 `targets` 的依据是交付范围，不是 `all` 失败之后任意缩小。目标范围内算不完整，这一步默认仍保存可得结果并列清单；顶层 `requireCompleteCalculation:true` 才按旧行为拒绝。该选项在批次、会话打开与一次性脚本顶层设默认，各分组步骤可用 true/false 覆盖；与 calculation 独立继承，不沿用前一步覆盖。
- `targets` 里 `s` 为 `result` 的值可以采信，但只证明所列的格及其上游，不能据此说整本已经重算。

## 范围怎么读

响应的 `calculation.scope` 有 `mode`、`targets`、`selectedFormulas`、`omittedFormulas`、`workbookComplete`。`status:"complete"` 配 `workbookComplete:false` 只表示所选范围算完了。

- 目标至多 100 个（`limits.calculationTargets`），每个是一格或一个矩形；相邻的格合并成矩形写。
- 带 `validate` 时校验的范围并入这次计算，动态数组也会把范围扩大；保证仍只对你列出的格成立。

## 省略时的默认与沿用

| 调用 | 省略 `calculation` 时 |
|---|---|
| `run`、`preflight`、`audit` | `all` |
| 批次的步骤、`session.run`、`wb.run` | 沿用批次根、`session.open`、脚本请求上的那个；那里也没写就是 `all` |
| `read` 命令、`session.read`、`wb.read` | `none`，不沿用上层 |
| `query`、`session.query`、`wb.query` | `none`，不沿用上层；带 `ruleResults` 或 `source:"calculation"` 的 `select` 时，程序自动只算它们要用的格 |

- 步骤自己写了 `calculation`，就整个对象替换上层的：上层写的 `locale`、`characterSet`、`textSortLocale` 不带过来，要用就在这一步重写。时钟与随机种子例外，在一个批次、会话、脚本的范围内固定，各步共用。
- `calculation` 写了就必须带 `mode`。
- 命令行 `read`：`--calculate` 是 `all`；`--calculation targets` 的目标就是所读的范围。

## `none` 下仍会算的地方

`none` 只是不做最后那次重算。下面的操作自己要用当前值，会就地算所需的区域及其上游：

| 操作 | 什么时候算 | 算不出来时 |
|---|---|---|
| `range.copy`，`mode:"values"` | 来源里有公式或数组 | `RANGE_CALCULATION_INCOMPLETE` |
| `data` 的工作表来源 | 写了 `formulaValues:"calculate"`；不写就拒绝含公式的来源 | `DATA_CALCULATION_INCOMPLETE` |
| `selection.edit` | 选择写了 `source:"calculation"` | `SELECTION_INCOMPLETE` |
| `filter.set`、`filter.reapply` | 写了 `formulaValues:"calculate"`；要用公式的值而没写，报 `FILTER_FORMULA_VALUES_REQUIRED` | `FILTER_CALCULATION_INCOMPLETE` |
| `range.sort`、`range.sort.reapply`、`range.deduplicate`、`range.split` | 同上；没写报 `DATA_FORMULA_VALUES_REQUIRED` | `DATA_CALCULATION_INCOMPLETE` |
| `pivot.create`、`pivot.set`、`pivot.refresh`、`pivot.copy`、`pivot.move` | 来源区域里有公式 | `PIVOT_CALCULATION_INCOMPLETE` |
| `layout.autofit` | 范围里有公式 | `AUTOFIT_CALCULATION_INCOMPLETE` |

这些值不算这一步的结果：响应顶层的 `calculation.status` 仍是 `not_requested`；同一步读来源格，得到的是 `cache`，这一步改过它的上游时是 `invalidated`（旧数，保存后没有值）。

## 时钟与随机数

`RAND`、`TODAY`、`NOW` 按请求的随机种子与时钟（`calculation.clock`）计算，不会与 Excel 的某次重算相同。

## 例

只算一个目标，读回目标之内与之外各一格：

```sh
cat > targets.json <<'EOF'
{"input":"input.xlsx","output":"targets.xlsx",
 "operations":[{"op":"write","sheet":"Data","range":"C2","values":[[75]]}],
 "calculation":{"mode":"targets","targets":[{"sheet":"Sheet1","range":"D2"}]},
 "read":[{"sheet":"Sheet1","range":"D2:D3","form":"compact"}]}
EOF
./excel-kit run targets.json > targets.out.json || true
python3 - <<'EOF'
import json
d = json.load(open('targets.out.json'))
if not d['ok']:
    raise SystemExit(f"failed: {json.dumps(d.get('error') or d.get('calculation'))}")
print(d['calculation']['status'], d['calculation']['scope'])   # complete，而 workbookComplete 是 false
print(d['warnings'])                                          # FORMULAS_WITHOUT_VALUES
print(d['ranges'][0]['rows'])                                 # D2 是 result，D3 不是
EOF
```
