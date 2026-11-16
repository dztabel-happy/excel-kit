// Update values/formulas or add a summary sheet, preserving the original
// no-password protection permissions. The script publishes only after the
// restored state has been checked. args: { operations, calculation? }.
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const operations = args.operations;
if (!Array.isArray(operations) || !operations.length || operations.some(op => !op || !['write', 'sheet.add'].includes(op.op))) {
  throw new Error('Provide write and/or sheet.add operations. This recipe preserves existing sheet identities and cell locking.');
}
const sheets = must(await wb.inspect(), 'Inspect sheets').workbook.sheets;
const selections = [{}, ...sheets.map(sheet => ({ sheet: sheet.name }))];
const before = must(await wb.query({ protection: selections }), 'Read protection').protection;
const named = name => typeof name === 'string' ? name.toLocaleLowerCase('en-US') : '';
const written = new Set(operations.filter(op => op.op === 'write').map(op => named(op.sheet)));
const changedScopes = before.filter(state => state.protected && (state.scope === 'workbook'
  ? operations.some(op => op.op === 'sheet.add') : written.has(named(state.sheet))));
// Check every required scope before issuing any edit; never guess a password.
for (const state of changedScopes) {
  const label = state.scope === 'workbook' ? 'Workbook structure' : `Worksheet ${state.sheet}`;
  if (state.password.present) throw new Error(`${label} requires a password. 请用户在 Excel 中自行解除相应保护并另存副本，再用该副本重试；本配方不猜测或移除口令校验。`);
  if (state.unsupportedAttributes.length) throw new Error(`${label} has unsupported protection attributes: ${state.unsupportedAttributes.join(', ')}. No changes were published.`);
}
const unlock = changedScopes.map(state => state.scope === 'workbook'
  ? { op: 'workbook.unprotect' } : { op: 'sheet.unprotect', sheet: state.sheet });
const restore = [...changedScopes].reverse().map(state => state.scope === 'workbook'
  ? { op: 'workbook.protect' } : { op: 'sheet.protect', sheet: state.sheet, allow: state.allow });
// Unprotect, edit and restore are a single step; even a failure halfway through
// cannot leave an unlocked version behind. Explicit none preserves formulas
// without claiming that this maintenance task has calculated the whole model.
const edited = await wb.run({ operations: [...unlock, ...operations, ...restore], calculation: args.calculation ?? { mode: 'none' } });
if (!edited.ok) throw new Error(`${edited.error.code}: ${edited.error.message}`);
const after = must(await wb.query({ protection: selections }), 'Read protection').protection;
for (let i = 0; i < before.length; i++) {
  if (JSON.stringify(before[i]) !== JSON.stringify(after[i])) throw new Error(`Protection restoration did not match the original state for ${before[i].sheet ?? 'workbook structure'}.`);
}
return { operations: operations.length, restored: changedScopes.map(state => state.scope === 'workbook' ? { scope: 'workbook' } : { scope: 'sheet', sheet: state.sheet }), protection: after };
