// List every picture and embedded attachment in a workbook with its size and
// SHA-256, and return one ready-to-run extraction request per item. A script
// cannot write files other than the workbook, so extraction itself is done
// afterwards with the image-extract and attachment-extract commands; their
// output.version must equal the sha256 listed here.
//
// args: { workbook, folder? }
//   workbook  path of the workbook as the extract commands should see it
//   folder    where extracted files should go (default "extracted")
const must = (response, what) => {
  if (!response.ok) {
    const first = response.calculation?.diagnostics?.[0];
    throw new Error(response.error ? `${what}: ${response.error.code} ${response.error.message}` : `${what}: CALCULATION_INCOMPLETE calculation ${response.calculation?.status}${first ? ` (${first.code} at ${first.sheet}!${first.cell}: ${first.message})` : ''}`);
  }
  return response;
};
const folder = args.folder ?? 'extracted';
const sheets = must(await wb.inspect(), 'Inspect sheets').workbook.sheets.map(sheet => sheet.name);

// Page through a per-sheet listing; every page after the first is bound to the version of the first.
async function listAll(kind, sheet) {
  const items = [];
  for (let offset = 0; offset !== null;) {
    const page = must(await wb.query({ [kind]: [{ sheet, offset, limit: 100 }], ...(offset ? { expectedInputVersion: wb.version } : {}) }), `List ${kind}`)[kind][0];
    items.push(...page.items); offset = page.nextOffset;
  }
  return items;
}
// Stored filenames are data, never paths: keep only a short alphanumeric extension.
const extension = (name, fallback) => { const match = /\.([A-Za-z0-9]{1,8})$/.exec(name ?? ''); return match ? match[1].toLowerCase() : fallback; };

const images = [], attachments = [];
for (const [index, sheet] of sheets.entries()) {
  for (const image of await listAll('images', sheet)) {
    const { format, bytes, sha256, width, height } = image.metadata;
    images.push({ sheet, name: image.name, format, width, height, bytes, sha256,
      extract: { command: 'image-extract', request: { input: args.workbook, sheet, name: image.name, output: `${folder}/image-${index + 1}-${images.length + 1}.${extension('.' + format, 'bin')}`, expectedInputVersion: wb.version } } });
  }
  for (const item of await listAll('ole', sheet)) {
    attachments.push({ sheet, id: item.id, progId: item.progId, filename: item.filename ?? null, container: item.container, bytes: item.bytes, sha256: item.sha256,
      extract: { command: 'attachment-extract', request: { input: args.workbook, sheet, id: item.id, output: `${folder}/attachment-${index + 1}-${item.id}.${extension(item.filename, 'bin')}`, expectedInputVersion: wb.version } } });
  }
}
return { workbookVersion: wb.version, sheets: sheets.length, images, attachments, totalBytes: [...images, ...attachments].reduce((sum, item) => sum + item.bytes, 0) };
