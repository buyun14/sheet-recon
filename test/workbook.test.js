const assert = require('node:assert/strict');
const { test } = require('node:test');
const XLSX = require('../vendor/sheetjs/xlsx.full.min.js');
const core = require('excel-reconcile');
const excel = require('excel-reconcile/workbook');
const report = require('excel-reconcile/report');
const { manufacturing } = require('excel-reconcile/profiles');

function asSheet(rows, origin) {
  return XLSX.utils.aoa_to_sheet(rows, { origin });
}

function makeInput(rows, bookType) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, asSheet([]), '空白');
  XLSX.utils.book_append_sheet(workbook, asSheet(rows, 'B3'), '测试数据');
  return XLSX.write(workbook, { type: 'buffer', bookType });
}

function readModel(input, profile) {
  const workbook = excel.readWorkbook(input);
  const sheets = excel.inspectSheets(workbook);
  const name = excel.pickDefaultSheet(workbook.SheetNames, sheets);
  assert.equal(name, '测试数据');
  const rows = excel.sheetToRows(workbook.Sheets[name]);
  const { row } = core.detectHeaderRow(rows, 8, profile);
  return core.buildSheetModel(rows, row);
}

for (const bookType of ['xlsx', 'xlsm', 'xls']) {
  test(`通用工作簿往返 (${bookType})，保留行号、共同列和文本编码`, () => {
    const customer = readModel(makeInput([
      ['Ref', 'Value'], ['001', 2], ['002', 5], ['003', 8]
    ], bookType));
    const ours = readModel(makeInput([
      ['Ref', 'Value'], ['001', 3], ['002', 5], ['004', 8]
    ], bookType));

    assert.equal(customer.headerRow, 3);
    assert.equal(customer.rows[0].rowNo, 4);
    const key = core.guessMatchKey(customer, ours);
    assert.deepEqual(key.names, ['Ref']);
    const result = core.compareModels(customer, ours, key.names);
    const { pairedRows, missingRows, extraRows, diffCells } = result.metrics;
    assert.deepEqual([pairedRows, missingRows, extraRows, diffCells], [2, 1, 1, 1]);

    const reportWorkbook = report.buildReportWorkbook(result, { generatedAt: 'test' });
    const output = XLSX.write(reportWorkbook, { type: 'array', bookType: 'xlsx' });
    const reopened = excel.readWorkbook(output);
    assert.deepEqual(reopened.SheetNames, ['概览', '差异单元格明细', '缺失与多余行']);
    assert.deepEqual(excel.sheetToRows(reopened.Sheets.差异单元格明细)[1], [4, 4, '001', 'Value', '2', '3']);
    assert.equal(excel.sheetToRows(reopened.Sheets.缺失与多余行).length, 3);
  });
}

test('生产物料规则跳过说明行，并保留原来的复合键优先级', () => {
  const input = makeInput([
    ['这是说明'], ['设计批号', '任务代码', '料品'], ['B1', 'T1', 'P1']
  ], 'xlsx');
  const model = readModel(input, manufacturing);
  assert.equal(model.headerRow, 4);
  assert.deepEqual(core.guessMatchKey(model, model, manufacturing).names, ['设计批号', '任务代码', '料品']);
  assert.equal(readModel(input).headerRow, 3);
});

test('内置解析器读取 Excel 日期单元格', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = { A1: { t: 'n', v: 46318, z: 'yyyy-mm-dd' }, '!ref': 'A1' };
  XLSX.utils.book_append_sheet(workbook, sheet, '日期');
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const date = excel.sheetToRows(excel.readWorkbook(bytes).Sheets.日期)[0][0];
  assert.equal(core.text(date), '2026-10-23');
  assert.equal(core.cellEquals(date, '2026-10-23'), true);
});

test('报告在明细达到上限后保留完整统计和截断提示', () => {
  const customer = core.buildSheetModel([['ID', 'Value'], ['A', 1], ['B', 2]], 1);
  const ours = core.buildSheetModel([['ID', 'Value'], ['A', 3], ['C', 4]], 1);
  const result = core.compareModels(customer, ours, ['ID']);
  const workbook = report.buildReportWorkbook(result, { warnings: ['自定义提示'] }, 1);
  const overview = excel.sheetToRows(workbook.Sheets.概览);
  assert.deepEqual(overview.find(([label]) => label === '多余行（仅我方有）').slice(0, 2), ['多余行（仅我方有）', 1]);
  assert.ok(overview.some(([label]) => label === '自定义提示'));
  assert.ok(overview.some(([label]) => label.includes('导出已截断（共 2 行）')));
  assert.equal(excel.sheetToRows(workbook.Sheets.缺失与多余行).length, 2);
});
