// Apply the same formatting, print settings and view to every worksheet whose
// name matches a pattern, one sheet per step, and report what was done to
// which sheet. Sheets that do not match are not touched.
//
// args: { pattern, format?, print?, view?, freeze?, expectSheets? }
//   pattern       ECMAScript regular expression tested against each sheet name
//   format        [{ range, style }] applied on every matching sheet
//   print         print.set properties
//   view          sheet.view settings
//   freeze        { rows, columns }
//   expectSheets  number of sheets that must match, as a guard against a wrong pattern
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const pattern = new RegExp(args.pattern, 'u');
const sheets = must(await wb.inspect(), 'Inspect sheets').workbook.sheets;
const matching = sheets.filter(sheet => pattern.test(sheet.name));
if (args.expectSheets !== undefined && matching.length !== args.expectSheets) throw new Error(`Pattern matched ${matching.length} sheets (${matching.map(sheet => sheet.name).join(', ') || 'none'}), expected ${args.expectSheets}.`);

const applied = [];
for (const { name: sheet } of matching) {
  const operations = [
    ...(args.format ?? []).map(item => ({ op: 'format', sheet, range: item.range, style: item.style })),
    ...(args.freeze ? [{ op: 'layout', sheet, freeze: args.freeze }] : []),
    ...(args.print ? [{ op: 'print.set', sheet, ...args.print }] : []),
    ...(args.view ? [{ op: 'sheet.view', sheet, settings: args.view }] : []),
  ];
  // Formatting does not change any value, so no calculation is needed for these steps.
  const step = await wb.run({ id: sheet, operations, calculation: { mode: 'none' } });
  // One sheet that cannot take the settings (for example a protected one) makes the whole job fail: a half-styled report is worse than none.
  if (!step.ok) throw new Error(`${sheet}: ${step.error.code} ${step.error.message}`);
  applied.push({ sheet, operations: operations.length });
}
return { matched: matching.map(sheet => sheet.name), untouched: sheets.filter(sheet => !pattern.test(sheet.name)).map(sheet => sheet.name), applied };
