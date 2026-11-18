(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./core.js'), require('../vendor/sheetjs/xlsx.full.min.js'));
  } else {
    root.ReconcileWorkbook = factory(root.ReconcileCore, root.XLSX);
  }
})(globalThis, function (core, XLSX) {
  'use strict';

  function readWorkbook(data) {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const workbook = XLSX.read(bytes, {
      type: 'array',
      cellDates: true,
      cellNF: false,
      cellText: false,
      WTF: false
    });
    if (!workbook.SheetNames.length) throw new Error('工作簿中没有工作表。');
    return workbook;
  }

  function sheetHasData(sheet) {
    if (!sheet || !sheet['!ref']) return false;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    // Bound the scan used for the sheet picker. Every sheet remains selectable.
    const lastRow = Math.min(range.e.r, range.s.r + 39);
    const lastColumn = Math.min(range.e.c, range.s.c + 79);
    for (let r = range.s.r; r <= lastRow; r++) {
      for (let c = range.s.c; c <= lastColumn; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (cell && !core.isBlank(cell.v)) return true;
      }
    }
    return false;
  }

  function inspectSheets(workbook) {
    const metadata = Object.create(null);
    for (const name of workbook.SheetNames) metadata[name] = sheetHasData(workbook.Sheets[name]);
    return metadata;
  }

  function pickDefaultSheet(names, metadata) {
    return names.find((name) => metadata[name]) || names[0] || '';
  }

  function sheetToRows(sheet) {
    if (!sheet || !sheet['!ref']) return [];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
      // Start at A1 so the model can retain original row and column positions.
      range: { s: { r: 0, c: 0 }, e: range.e }
    });
    while (rows.length && rows.at(-1).every(core.isBlank)) rows.pop();
    return rows;
  }

  return { readWorkbook, inspectSheets, pickDefaultSheet, sheetToRows };
});
