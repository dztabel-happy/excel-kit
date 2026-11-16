// Roll a period workbook forward: copy the latest period sheet, enter the new
// period's figures, change shared assumptions, add the period to the summary,
// then refresh every chart, sparkline group and pivot in the workbook so none
// of them shows stale numbers.
//
// args: { source, target, updates?, assumptions?, summary? }
//   source       period sheet to copy; target: name of the new sheet
//   updates      [{ cell, value }] written on the new sheet
//   assumptions  [{ sheet, cell, value }] shared inputs to change
//   summary      { sheet, range, values } a row to add to the summary (values may hold formulas)
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const names = must(await wb.inspect(), 'Inspect').workbook.sheets.map(sheet => sheet.name);
if (!names.includes(args.source)) throw new Error(`Source sheet ${args.source} not found.`);
if (names.includes(args.target)) throw new Error(`Sheet ${args.target} already exists; this period was rolled forward before.`);

// 1. Copy and edit in one step. The copy brings its chart, sparklines, print settings and sheet-local names, retargeted to the new sheet.
//    Pivots whose source changes must be refreshed in the step that changes it, so the existing pivots are listed first and refreshed here.
const listed = must(await wb.query({ pivots: names.map(sheet => ({ sheet })) }), 'List pivots');
const pivots = listed.pivots.flatMap(page => page.items.map(item => ({ sheet: page.sheet, name: item.name })));
must(await wb.run({ id: 'roll-forward', operations: [
  { op: 'sheet.copy', sheet: args.source, name: args.target },
  ...(args.updates ?? []).map(item => ({ op: 'write', sheet: args.target, range: item.cell, values: [[item.value]] })),
  ...(args.assumptions ?? []).map(item => ({ op: 'write', sheet: item.sheet, range: item.cell, values: [[item.value]] })),
  ...(args.summary ? [{ op: 'write', sheet: args.summary.sheet, range: args.summary.range, values: args.summary.values }] : []),
  ...pivots.map(item => ({ op: 'pivot.refresh', ...item })),
] }), 'Roll forward');

// 2. Charts and sparklines cache their points. List them on every sheet, the new one included, and refresh them all in one step.
const sheets = must(await wb.inspect(), 'Inspect').workbook.sheets.map(sheet => sheet.name);
const found = must(await wb.query({ charts: sheets.map(sheet => ({ sheet })), sparklines: sheets.map(sheet => ({ sheet })) }), 'List objects');
const charts = found.charts.flatMap(page => page.items.map(item => ({ sheet: page.sheet, name: item.name })));
const sparklines = found.sparklines.flatMap(page => page.items.map(item => ({ sheet: page.sheet, index: item.index })));
must(await wb.run({ id: 'refresh-objects', operations: [
  ...charts.map(item => ({ op: 'chart.refresh', ...item })),
  ...sparklines.map(item => ({ op: 'sparkline.refresh', ...item })),
] }), 'Refresh');

// 3. Report what the refreshed objects now hold.
const after = must(await wb.query({ charts: [...new Set(charts.map(item => item.sheet))].map(sheet => ({ sheet })), pivots: [...new Set(pivots.map(item => item.sheet))].map(sheet => ({ sheet })) }), 'Read refreshed objects');
const chartValues = after.charts.flatMap(page => page.items.map(item => ({ sheet: page.sheet, name: item.name,
  values: item.series.map(series => (series.sources.find(source => source.role === 'val').cache?.points ?? []).map(point => Number(point.value))) })));
return { copied: { from: args.source, to: args.target }, refreshed: { charts: charts.length, sparklines: sparklines.length, pivots: pivots.length }, charts: chartValues,
  pivots: after.pivots.flatMap(page => page.items.map(item => ({ sheet: page.sheet, name: item.name, state: item.state ?? null }))) };
