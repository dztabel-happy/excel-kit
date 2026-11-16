// Find the records that match a column condition, change them in one step,
// recalculate only the results you care about and read them back.
//
// args: { target, where, set, read, expectMatches? }
//   target  selection target, e.g. { kind: 'rows', table: '交易明细' }
//   where   selection condition, e.g. { column: '数量', operator: 'gte', value: 2 }
//   set     [{ column, value }] applied to every matching row
//   read    [{ sheet, range }] results to recalculate and return
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const select = { target: args.target, where: args.where };
const found = must(await wb.query({ select: [{ ...select, limit: 1 }] }), 'Select').select[0];
if (args.expectMatches !== undefined && found.total !== args.expectMatches) throw new Error(`Expected ${args.expectMatches} matching records, found ${found.total}.`);
if (found.total === 0) return { matched: 0, changed: false, results: [] };

// expectedMatches makes the edit refuse to run if the number of matches differs from the one just counted.
// Edit, calculate just the requested results and read them in the same step.
const edit = await wb.run({
  operations: [{ op: 'selection.edit', selection: select, action: { kind: 'set', values: args.set }, expectedMatches: found.total }],
  calculation: { mode: 'targets', targets: args.read }, read: args.read,
}).catch(error => error.response);
if (!edit.ok) {
  const first = edit.error?.details?.diagnostics?.[0];
  throw new Error(`Edit failed: ${edit.error.code} ${edit.error.message}${first ? ` (${first.code} at ${first.sheet}!${first.cell})` : ''}`);
}
return {
  matched: found.total, changed: true,
  ...(edit.result.calculation.status === 'incomplete' ? { unavailable: edit.result.calculation.unavailable } : {}),
  results: edit.result.ranges.map(range => ({ sheet: range.sheet, range: range.range, values: range.cells.map(row => row.map(cell => cell.result?.value ?? null)) })),
};
