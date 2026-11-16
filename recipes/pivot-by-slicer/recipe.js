// Append detail rows, refresh the pivot built on them, then walk through a
// slicer one item at a time and read the formulas that depend on the pivot
// for each selection. Finishes with the selection cleared and checks that
// the per-item values add up to the unfiltered total.
//
// Formulas downstream of a pivot are only calculated in the step that
// refreshed it: a stored pivot result from an earlier step is never passed
// off as current. So every selection reads its results inside the same step.
//
// args: { append?, pivot, slicer, items, read, total }
//   append  { table, columns, values } rows to add first (optional)
//   pivot   { sheet, name }; slicer { sheet, name }
//   items   slicer items to visit, in order
//   read    [{ sheet, range }] downstream cells to read for every selection
//   total   index into `read` of a single cell that is additive across items, used for the consistency check
const value = cell => cell.result?.value ?? null;
const cellsOf = step => step.result.ranges.map(range => range.cells.map(row => row.map(value)));
const must = (step, what) => { if (!step.ok) throw new Error(`${what}: ${step.error.code} ${step.error.message}`); return step; };

if (args.append) must(await wb.run({ operations: [{ op: 'table.append', ...args.append }, { op: 'pivot.refresh', ...args.pivot }] }), 'Append and refresh');

const selections = [];
for (const item of args.items) {
  const step = must(await wb.run({ id: String(item), operations: [{ op: 'slicer.select', ...args.slicer, items: [item] }], read: args.read }), `Select ${item}`);
  selections.push({ item, values: cellsOf(step) });
}
// Leave the workbook unfiltered, and read the same cells once more for the total.
const cleared = must(await wb.run({ id: 'all', operations: [{ op: 'slicer.select', ...args.slicer, items: null }], read: args.read }), 'Clear selection');
const all = cellsOf(cleared);

const pick = values => values[args.total][0][0];
const sum = selections.reduce((total, selection) => total + (pick(selection.values) ?? 0), 0);
if (Math.abs(sum - pick(all)) > 1e-9) throw new Error(`Per-item values add up to ${sum}, but the unfiltered total is ${pick(all)}. Is an item missing from args.items?`);
return { selections, all, sumOfItems: sum };
