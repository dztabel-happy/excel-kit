// Fill a structured template, let the workbook's own data-validation rules
// judge the result before anything is published, then add print settings and
// protection.
//
// args: { sheet, data, numbers?, validate, print?, protect? }
//   sheet     the form sheet
//   data      JSON object for the {{key}} placeholders (text fields)
//   numbers   [{ cell, value }] native numbers; template.fill only writes text
//   validate  [range, ...] cells whose validation rules must pass
//   print     print.set properties for the sheet
//   protect   { allow?, password? } sheet protection to add last
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const { sheet } = args;
// One step: fill, write the numbers, and have the rules checked. If a rule fails the step fails and nothing of it remains.
const filled = await wb.run({
  operations: [
    { op: 'template.fill', sheet, data: args.data, unresolved: 'error' },
    ...(args.numbers ?? []).map(item => ({ op: 'write', sheet, range: item.cell, values: [[item.value]] })),
  ],
  validate: args.validate.map(range => ({ sheet, range })),
}).catch(error => error.response);
if (!filled.ok) {
  const failures = (filled.error.details?.ranges ?? []).flatMap(range => (range.failures ?? []).map(cell => `${range.sheet}!${cell.cell}`));
  throw new Error(`${filled.error.code}: ${filled.error.message}${failures.length ? ' Cells: ' + failures.join(', ') : ''}`);
}

// Presentation and protection come last, so a protected sheet never blocks the filling above.
const finished = await wb.run({ operations: [
  ...(args.print ? [{ op: 'print.set', sheet, ...args.print }] : []),
  ...(args.protect ? [{ op: 'sheet.protect', sheet, ...args.protect }] : []),
] });
if (!finished.ok) throw new Error(`${finished.error.code}: ${finished.error.message}`);

const template = filled.result.templates[0];
const state = must(await wb.query({ protection: [{ sheet }], printing: [{ sheet }] }), 'Read presentation');
return {
  replacements: template.replacements, changedCells: template.changedCells, usedKeys: template.usedKeys, unresolved: template.unresolved.occurrences,
  validation: filled.result.validation.ranges.map(range => ({ range: range.range, ...range.counts })),
  protected: state.protection[0].protected, printAreas: state.printing[0].areas,
};
