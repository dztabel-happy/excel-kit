// Normalize a code format across the whole workbook with one regular
// expression, then prove with a search that nothing in the old format is left.
//
// args: { pattern, replacement, caseSensitive?, sheets? }
//   pattern      ECMAScript regular expression for the irregular form, e.g. "SKU-(\\d{4})-X"
//   replacement  replacement with $1-style tokens, e.g. "$1"
//   sheets       optional list of sheet names; default is every sheet, hidden ones included
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const caseSensitive = args.caseSensitive ?? false;
const search = sheet => ({ ...(sheet ? { sheet } : {}), field: 'value', value: { type: 'string', value: args.pattern }, match: 'contains', syntax: 'regex', caseSensitive, limit: 1000 });
const scopes = args.sheets ?? [undefined];

const before = must(await wb.query({ find: scopes.map(search) }), 'Find codes').find;
const found = before.reduce((sum, result) => sum + result.total, 0);
if (found === 0) return { found: 0, replacements: 0, remaining: 0, cells: [] };

let replacements = 0;
for (const sheet of scopes) {
  const step = await wb.run({ op: 'range.replace', ...(sheet ? { sheet } : {}), find: args.pattern, replacement: args.replacement, syntax: 'regex', match: 'partial', caseSensitive });
  if (!step.ok) throw new Error(`Replacement failed: ${step.error.code} ${step.error.message}`);
  for (const event of step.result.rangeOperations ?? []) replacements += event.replacements ?? 0;
}

// The same search, against the edited workbook. Anything left is a reason to stop, not to publish.
const after = must(await wb.query({ find: scopes.map(search) }), 'Find codes').find;
const remaining = after.reduce((sum, result) => sum + result.total, 0);
if (remaining > 0) throw new Error(`${remaining} cells still match ${args.pattern}: ${after.flatMap(result => result.items).slice(0, 5).map(item => `${item.sheet}!${item.cell}`).join(', ')}`);
return { found, replacements, remaining, cells: before.flatMap(result => result.items.map(item => `${item.sheet}!${item.cell}`)) };
