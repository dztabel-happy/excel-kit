// Audit a whole workbook for formulas that cannot be trusted (broken
// references, unknown functions, cycles, errors), report where each problem
// originates, and apply the corrections that were chosen for them. Without
// args.fixes it only inspects. With fixes it publishes only if the second
// audit is clean and every stated business rule holds.
//
// args: { fixes?, assertions? }
//   fixes       [{ sheet, cell, formula }] corrections chosen after reading the candidates
//   assertions  query assertions ({ id, basis, kind, ... }) that must pass after the repair
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const audit = async () => must(await wb.query({ calculation: { mode: 'all' }, findings: { limit: 1000 }, ...(args.assertions ? { assertions: args.assertions } : {}) }), 'Audit');
const key = item => `${item.sheet}!${item.cell}`;

const before = await audit();
const diagnostics = before.findings.items.filter(item => item.kind === 'diagnostic');
const byCell = new Map();
for (const item of diagnostics) { const entry = byCell.get(key(item)) ?? { sheet: item.sheet, cell: item.cell, codes: [], messages: [] }; entry.codes.push(item.code); entry.messages.push(item.message); byCell.set(key(item), entry); }

// A problem cell is an origin when none of the cells it reads is itself a problem; otherwise it only inherits one.
const cells = [...byCell.values()];
const precedents = cells.length ? must(await wb.query({ dependencies: cells.map(entry => ({ sheet: entry.sheet, cell: entry.cell, direction: 'precedents' })) }), 'Read precedents').dependencies : [];
const inCycle = entry => entry.codes.includes('CIRCULAR_REFERENCE');
const candidates = cells.map((entry, index) => {
  const reads = precedents[index].items.map(item => `${item.sheet}!${item.range}`);
  const inherited = reads.filter(cell => byCell.has(cell) && cell !== key(entry) && !(inCycle(entry) && inCycle(byCell.get(cell))));
  return { ...entry, formula: precedents[index].formula, origin: inherited.length === 0, inheritedFrom: inherited };
});
const summary = result => ({ calculation: result.calculation.status, confirmed: result.assessment.confirmed, unavailable: result.assessment.unavailable, cacheDifferences: result.assessment.cacheDifferences });
if (!args.fixes?.length) return { mode: 'inspect', before: summary(before), candidates };

// All corrections go in one step: while a cycle or an unknown function remains, the workbook cannot be calculated at all.
const unknown = args.fixes.filter(fix => !byCell.has(key(fix)) && !fix.businessRule);
if (unknown.length) throw new Error(`These fixes address cells the audit did not flag; mark them businessRule:true if intended: ${unknown.map(key).join(', ')}`);
const repaired = await wb.run({ operations: args.fixes.map(fix => ({ op: 'write', sheet: fix.sheet, range: fix.cell, values: [[{ formula: fix.formula }]] })) });
if (!repaired.ok) throw new Error(`Repair failed: ${repaired.error.code} ${repaired.error.message}`);

const after = await audit();
const remaining = after.findings.items.filter(item => item.kind === 'diagnostic');
if (remaining.length) throw new Error(`Still not clean after the repair: ${remaining.map(item => `${key(item)} ${item.code}`).join(', ')}`);
const failed = (after.assertions ?? []).filter(item => item.status !== 'passed');
if (failed.length) throw new Error(`Business rules not satisfied: ${failed.map(item => `${item.id} (${item.status})`).join(', ')}`);
return { mode: 'repair', before: summary(before), after: summary(after), candidates, fixed: args.fixes.map(key), assertions: (after.assertions ?? []).map(item => ({ id: item.id, status: item.status })) };
