// Change only the stated assumptions, recalculate only the stated results,
// and return each result with everything it depends on: the formulas in its
// chain, the input cells at the ends of the chain with their values, and the
// workbook versions before and after. Refuses an "assumption" that none of
// the results depends on, and refuses to overwrite a formula.
//
// args: { assumptions, results }
//   assumptions  [{ sheet, cell, value }] inputs to change
//   results      [{ sheet, cell }] results to recalculate and explain
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const at = item => `${item.sheet}!${item.cell}`;
const targets = args.results.map(item => ({ sheet: item.sheet, range: item.cell }));
const resultValues = async () => must(await wb.read({ ranges: targets, calculation: { mode: 'targets', targets } }), 'Read results').ranges.map(range => range.cells[0][0]);

// The static dependency chain of every result, all the way up.
const chains = must(await wb.query({ dependencies: args.results.map(item => ({ ...item, direction: 'precedents', transitive: true, limit: 1000 })) }), 'Read dependencies').dependencies;
for (const chain of chains) if (chain.nextOffset !== null) throw new Error(`${at(chain)} has more than 1,000 upstream ranges; narrow the request.`);
const upstreamOf = chains.map(chain => [...new Set(chain.items.map(item => `${item.sheet}!${item.range}`))]);

// Read every upstream range once to tell formulas from inputs.
const ranges = [...new Set(upstreamOf.flat())].map(text => { const split = text.lastIndexOf('!'); return { sheet: text.slice(0, split), range: text.slice(split + 1) }; });
const upstreamCells = new Map();
if (ranges.length) for (const range of must(await wb.read({ ranges }), 'Read upstream').ranges) for (const row of range.cells) for (const cell of row) upstreamCells.set(`${range.sheet}!${cell.address}`, cell);

for (const assumption of args.assumptions) {
  const cell = upstreamCells.get(at(assumption));
  if (!cell) throw new Error(`${at(assumption)} is not upstream of any requested result; changing it cannot affect them.`);
  if (cell.kind === 'formula') throw new Error(`${at(assumption)} holds a formula (${cell.formula}); an assumption must be an input cell.`);
}

const before = await resultValues(), inputVersion = wb.version;
const step = await wb.run({ operations: args.assumptions.map(item => ({ op: 'write', sheet: item.sheet, range: item.cell, values: [[item.value]] })), calculation: { mode: 'targets', targets }, read: targets });
must(step, 'Update failed');
if (step.result.changes.cells !== args.assumptions.length) throw new Error(`Expected to change ${args.assumptions.length} cells, changed ${step.result.changes.cells}.`);
const after = step.result.ranges.map(range => range.cells[0][0]);

const changed = new Map(args.assumptions.map(item => [at(item), item.value]));
return {
  inputVersion, outputVersion: wb.version, changedCells: step.result.changes.cells,
  results: args.results.map((item, index) => {
    const cells = [...upstreamCells.entries()].filter(([address]) => belongs(address, upstreamOf[index]));
    return {
      sheet: item.sheet, cell: item.cell, formula: after[index].formula, before: before[index].result?.value ?? null, after: after[index].result?.value ?? null,
      upstreamFormulas: cells.filter(([, cell]) => cell.kind === 'formula').map(([address, cell]) => ({ cell: address, formula: cell.formula })),
      upstreamInputs: cells.filter(([, cell]) => cell.kind !== 'formula' && cell.result !== null && cell.result.type !== 'blank').map(([address, cell]) => ({ cell: address, value: changed.has(address) ? changed.get(address) : cell.result.value, changed: changed.has(address) })),
    };
  }),
};

// Whether a cell address lies inside one of the upstream ranges of this result.
function belongs(address, list) {
  const split = address.lastIndexOf('!'), sheet = address.slice(0, split), cell = address.slice(split + 1);
  const point = text => { const match = /^\$?([A-Z]+)\$?(\d+)$/.exec(text); return { column: [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0), row: Number(match[2]) }; };
  const here = point(cell);
  return list.some(text => {
    const cut = text.lastIndexOf('!'); if (text.slice(0, cut) !== sheet) return false;
    const [from, to = from] = text.slice(cut + 1).split(':').map(point);
    return here.row >= from.row && here.row <= to.row && here.column >= from.column && here.column <= to.column;
  });
}
