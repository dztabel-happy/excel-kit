# 按条件选与改

按条件选出格、行或对象，再一次改掉它们：怎样先数再改，条件在哪些地方不按直觉走。条件的写法查 `help select.where`，字段名查 `help select.fields`，动作查 `help selection.edit`。

## 先数再改

- 用 `query` 的 `select` 数出匹配数量（`select[0].total`），再用 `selection.edit` 修改，带上 `expectedMatches`。个数不符报 `SELECTION_COUNT_MISMATCH`（`details` 里有 `expected` 与 `actual`），这一步什么都不改。
- 先数和修改放在同一个脚本里做，中间不会有别的写入。
- 个数相符不证明选中的是你要的对象。改完看这一步响应里的 `selections[i].targets`：它列出改动之前的位置，只列前 100 个，多了带 `truncated:true`。
- `selection.edit` 的选择不带 `offset`、`limit`：改的总是完整的命中集，不能靠查询的分页挑着改。删行自下而上进行。
- 原生编辑里也有一个 `expectedMatches`（缺省 1），那是另一回事。

## 条件的语义

- 列名不区分大小写；值的文本比较**区分**大小写，要不区分就写 `"caseSensitive":false`。
- 数字与文本不互转。数字列对文本值：`eq` 一个都不中，`ne` 全部命中，`gt` 这类比较不中、并在结果的 `diagnostics` 里记一条 `SELECTION_NON_COMPARABLE`（带个数与实际值的样本）。
- 查询完整而命中 0 个时，结果里有 `noMatch`（`SELECTION_NO_MATCH`），附每个条件实际遇到的值的样本。先看样本的类型与大小写，再判断是真的没有。
- 公式的值：不写 `"source":"calculation"` 时算未知。条件按三值逻辑合并：`any` 里有一个为真就是真，`all` 里有一个为假就是假；这样仍判不出的记在 `unavailable` 里，查询标 `status:"incomplete"`，`selection.edit` 报 `SELECTION_INCOMPLETE`、整步拒绝。
- 写了 `"source":"calculation"`，程序只算条件要用的格。`query` 里这时不要再显式写 `"mode":"none"`（`SELECTION_CALCULATION_REQUIRED`），省略 `calculation` 即可；`selection.edit` 在 `none` 的步骤里也会就地算。
- 行的三种目标：`{"kind":"rows","table":…}`、`{"kind":"rows","sheet":…,"range":…}`（第一行是表头）、只写 `{"kind":"rows"}` 让程序按列名找。最后一种优先找真正的 Table，否则推断表头区域，并在结果的 `binding` 里标 `stable:false`；先看它找对了没有。

## 例

把 `Amount` 不小于 1990 的行的 `Limit` 改成 0。`Amount` 一列里有公式，所以条件要写 `source:"calculation"`；不写的话 `found.status` 是 `incomplete`。

```sh
cat > select-edit.js <<'EOF'
const selection = { target: { kind: 'rows', sheet: 'Sheet1', range: 'A1:D200' },
  where: { column: 'Amount', operator: 'gte', value: 1990 }, source: 'calculation' };
const found = (await wb.query({ select: [{ ...selection, limit: 1 }] })).select[0];
if (found.status !== 'complete') throw new Error(`selection ${found.status}: ${JSON.stringify(found.unavailable)}`);
if (found.total === 0) return { matched: 0 };
const step = await wb.run({ op: 'selection.edit', selection, expectedMatches: found.total,
  action: { kind: 'set', values: [{ column: 'Limit', value: 0 }] } });
return { matched: found.total, targets: step.result.selections[0].targets };
EOF
cat > select-edit.json <<'EOF'
{"input":"input.xlsx","output":"edited.xlsx","strict":true,"scriptPath":"select-edit.js"}
EOF
./excel-kit script select-edit.json > select-edit.out.json
```

请求带 `"strict":true`，所以脚本里不用逐个判 `ok`：任何一次失败都会抛出（见 [scripts.md](scripts.md)）。
