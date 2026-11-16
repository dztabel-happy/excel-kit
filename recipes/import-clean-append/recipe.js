// Import CSV files, clean and map their columns, drop duplicates, append the
// rest to a Table that has an active filter, reapply the filter and check the
// visible subtotal against an independent sum.
//
// args: { table, sources, columns, map, key, required, check }
//   table     Table to append to
//   sources   [{ csv, delimiter? }] files to import, in order
//   columns   typed CSV columns for the import, e.g. { name: '日期文本', type: 'date', format: 'YYYY-MM-DD' }
//   map       { <Table column>: { from: <CSV column>, trim?, pattern?, replacement? } }
//   key       Table column that identifies a record
//   required  Table columns that must not be empty after cleaning
//   check     { sheet, cell, column } a SUBTOTAL cell and the Table column it sums over visible rows
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const plain = range => range.cells.map(row => row.map(cell => cell.result?.value ?? null));

// 1. Stage every CSV on its own temporary sheet with typed columns. Dates and numbers are parsed by the import, never guessed.
const staging = args.sources.map((_, index) => `_import_${index + 1}`);
const staged = await wb.run({ operations: args.sources.flatMap((source, index) => [
  { op: 'sheet.add', name: staging[index] },
  { op: 'data', id: `source${index + 1}`, sheet: staging[index], range: 'A1', sources: [{ id: 'csv', csv: source.csv, delimiter: source.delimiter ?? ',', header: true }], columns: args.columns, steps: [] },
]) });
must(staged, 'Import failed');

// 2. Read the staged rows back and clean them here, in code.
const read = must(await wb.read({ ranges: staged.result.data.map(data => ({ sheet: data.destination.sheet, range: data.destination.range })) }), 'Read imports');
const clean = (value, rule) => {
  if (typeof value !== 'string') return value;
  let text = rule.trim ? value.trim().replace(/\s+/g, ' ') : value;
  if (rule.pattern) { const pattern = new RegExp(rule.pattern); if (!pattern.test(text)) return null; text = text.replace(pattern, rule.replacement ?? '$&'); }
  return text;
};
const records = [];
for (const [index, range] of read.ranges.entries()) {
  const [header, ...rows] = plain(range);
  for (const [offset, row] of rows.entries()) {
    const source = Object.fromEntries(header.map((name, column) => [name, row[column]]));
    const record = Object.fromEntries(Object.entries(args.map).map(([column, rule]) => [column, clean(source[rule.from], rule)]));
    records.push({ record, origin: `${args.sources[index].csv} row ${offset + 2}` });
  }
}

// 3. Reject incomplete records, then drop duplicates: within the files (first wins) and against the Table.
const table = must(await wb.inspect(), 'Inspect table').workbook.tables.find(item => item.name === args.table);
if (!table) throw new Error(`Table ${args.table} not found.`);
const names = table.columns.map(column => column.name);
const body = plain(must(await wb.read({ ranges: [{ sheet: table.sheet, range: table.range }] }), 'Read table').ranges[0]).slice(1);
const seen = new Set(body.map(row => row[names.indexOf(args.key)]));
const accepted = [], rejected = [], duplicates = [];
for (const { record, origin } of records) {
  const missing = args.required.filter(column => record[column] === null || record[column] === undefined || record[column] === '');
  if (missing.length) { rejected.push({ origin, missing }); continue; }
  if (seen.has(record[args.key])) { duplicates.push({ origin, key: record[args.key] }); continue; }
  seen.add(record[args.key]); accepted.push(record);
}

// 4. One step: drop the staging sheets, append, and reapply the Table's filter. Appending under an active filter requires the reapply in the same step.
const columns = Object.keys(args.map);
const appended = await wb.run({ operations: [
  ...staging.map(sheet => ({ op: 'sheet.delete', sheet })),
  ...(accepted.length ? [{ op: 'table.append', table: args.table, columns, values: accepted.map(record => columns.map(column => record[column])) }, { op: 'filter.reapply', target: { table: args.table } }] : []),
] });
must(appended, 'Append failed');

// 5. Check the subtotal: recalculate, then add up the visible rows independently.
const grown = must(await wb.inspect(), 'Inspect table').workbook.tables.find(item => item.name === args.table);
const after = must(await wb.read({ ranges: [{ sheet: args.check.sheet, range: args.check.cell }, { sheet: grown.sheet, range: grown.range }], calculation: { mode: 'all' } }), 'Read subtotal');
const subtotal = after.ranges[0].cells[0][0].result?.value ?? null, amount = names.indexOf(args.check.column);
const visible = after.ranges[1].cells.slice(1).filter(row => row[0].rowVisibility === 'visible');
const expected = visible.reduce((sum, row) => sum + row[amount].result.value, 0);
if (Math.abs(subtotal - expected) > 1e-9) throw new Error(`Subtotal ${subtotal} differs from the visible rows' sum ${expected}.`);
return { imported: records.length, appended: accepted.map(record => record[args.key]), duplicates, rejected, rows: after.ranges[1].cells.length - 1, visibleRows: visible.length, subtotal };
