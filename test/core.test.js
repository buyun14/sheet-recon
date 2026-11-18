const assert = require('node:assert/strict');
const { test } = require('node:test');
const RC = require('excel-reconcile');
const report = require('excel-reconcile/report');
const { manufacturing } = require('excel-reconcile/profiles');

const eq = assert.deepStrictEqual;
const throws = assert.throws;

function compareAoA(customer, ours, keyNames) {
  return RC.compareModels(RC.buildSheetModel(customer, 1), RC.buildSheetModel(ours, 1), keyNames);
}

test('表头识别：跳过前两行说明行', () => {
  const rows = [
    ['说明', null, null],
    ['表头', null, null],
    ['序号', '料品', '任务代码', '设计批号'],
    [1, 'A', 'T1', 'B1']
  ];
  const det = RC.detectHeaderRow(rows, 8, manufacturing);
  eq(det.row, 3);
  const model = RC.buildSheetModel(rows, det.row);
  eq(model.headers.map((h) => h.name), ['序号', '料品', '任务代码', '设计批号']);
  eq(model.rows.length, 1);
});

test('表头识别：表头行超出范围应报错', () => {
  throws(() => RC.buildSheetModel([['a'], ['b']], 3), /超出工作表范围/);
});

test('单元格比较：数字、空值、trim、大小写、日期', () => {
  eq(RC.cellEquals(100, '100.00'), true);
  eq(RC.cellEquals(1, '001'), true);
  eq(RC.cellEquals('001', '1'), false);
  eq(RC.cellEquals(null, '  '), true);
  eq(RC.cellEquals(' A ', 'A'), true);
  eq(RC.cellEquals('ABC', 'abc'), false);
  eq(RC.cellEquals(new Date(2026, 9, 23), '2026-10-23'), true);
  eq(RC.cellEquals(new Date(2026, 9, 22, 23, 59, 17), '2026-10-23'), true);
  eq(RC.text(new Date(2026, 9, 22, 23, 59, 17)), '2026-10-23');
  eq(RC.cellEquals('', null), true);
});

test('行级：缺失、多余、配对，单元格差异 1 个', () => {
  const result = compareAoA(
    [['序号', '料品', '数量'], [1, 'A', 100], [2, 'B', 200], [3, 'C', 300]],
    [['序号', '料品', '数量'], [1, 'A', '100'], [3, 'C', '301'], [4, 'D', 400]],
    ['料品']
  );
  eq(result.metrics, {
    customerRows: 3,
    ourRows: 3,
    pairedRows: 2,
    missingRows: 1,
    extraRows: 1,
    diffCells: 1,
    customerBlankKeys: 0,
    ourBlankKeys: 0,
    duplicateKeyGroups: 0,
    duplicateCustomerRows: 0,
    duplicateOurRows: 0,
    headerOnlyCustomer: 0,
    headerOnlyOurs: 0
  });
  eq(result.rows.map((r) => r.status), ['paired', 'missing', 'paired', 'extra']);
  eq(result.rows[2].diffs[0].name, '数量');
  eq(result.rows[2].diffs[0].customerValue, 300);
  eq(result.rows[2].diffs[0].ourValue, '301');
});

test('重复配对键：按出现顺序配对，并给出重复警告', () => {
  const result = compareAoA(
    [['k', 'v'], ['A', '1'], ['A', '2'], ['A', '3']],
    [['k', 'v'], ['A', 'x'], ['A', '2']],
    ['k']
  );
  eq([result.metrics.pairedRows, result.metrics.missingRows, result.metrics.extraRows], [2, 1, 0]);
  eq(result.metrics.diffCells, 1);
  eq(result.metrics.duplicateKeyGroups, 1);
  eq(result.metrics.duplicateCustomerRows, 2);
});

test('空配对键：按缺失/多余处理', () => {
  const result = compareAoA(
    [['k', 'v'], [null, '1'], [null, '2']],
    [['k', 'v'], [null, '1']],
    ['k']
  );
  eq([result.metrics.pairedRows, result.metrics.missingRows, result.metrics.extraRows], [0, 2, 1]);
  eq([result.metrics.customerBlankKeys, result.metrics.ourBlankKeys], [2, 1]);
});

test('表头单侧列：不参与逐单元格比较，但会提示', () => {
  const result = compareAoA(
    [['k', 'A'], ['x', '1']],
    [['k', 'B'], ['x', '2']],
    ['k']
  );
  eq(result.metrics.diffCells, 0);
  eq(result.metrics.headerOnlyCustomer, 1);
  eq(result.metrics.headerOnlyOurs, 1);
  eq(result.columns.map((c) => [c.display, c.oneSide]), [['k', null], ['A', 'customer'], ['B', 'ours']]);
  eq(result.headerDiff.onlyCustomer, ['A']);
  eq(result.headerDiff.onlyOurs, ['B']);
});

test('列顺序不同：按列名对齐，不产生差异', () => {
  const result = compareAoA(
    [['k', 'A', 'B'], ['x', '1', '2']],
    [['k', 'B', 'A'], ['x', '2', '1']],
    ['k']
  );
  eq(result.metrics.diffCells, 0);
  eq(result.columns.map((c) => c.display), ['k', 'A', 'B']);
  eq(result.columns.map((c) => [c.customerIndex, c.ourIndex]), [[0, 0], [1, 2], [2, 1]]);
});

test('重复表头名：自动加后缀，两侧仍能对齐', () => {
  const c = RC.buildSheetModel([['k', '备注', '备注'], ['x', 'a', 'b']], 1);
  const o = RC.buildSheetModel([['k', '备注', '备注'], ['x', 'a', 'b']], 1);
  eq(c.headers.map((h) => h.name), ['k', '备注', '备注 (2)']);
  const result = RC.compareModels(c, o, ['k']);
  eq(result.metrics.diffCells, 0);
  eq(result.columns.length, 3);
});

test('自动配对键优先级：设计批号 + 任务代码 + 料品', () => {
  const result = compareAoA(
    [['设计批号', '任务代码', '料品', 'BOM号', '物料代码'], ['B1', 'T1', 'A', 'BOM1', 'M1']],
    [['设计批号', '任务代码', '料品', 'BOM号', '物料代码'], ['B1', 'T1', 'A', 'BOM1', 'M1']],
    ['设计批号', '任务代码', '料品']
  );
  const c = RC.buildSheetModel([['设计批号', '任务代码', '料品', 'BOM号', '物料代码'], ['B1', 'T1', 'A', 'BOM1', 'M1']], 1);
  const o = RC.buildSheetModel([['设计批号', '任务代码', '料品', 'BOM号', '物料代码'], ['B1', 'T1', 'A', 'BOM1', 'M1']], 1);
  eq(RC.guessMatchKey(c, o, manufacturing).names, ['设计批号', '任务代码', '料品']);
  eq(result.metrics.diffCells, 0);
});

test('自动配对键优先级：任务代码 + BOM号 + 物料代码', () => {
  const c = RC.buildSheetModel([['任务代码', 'BOM号', '物料代码'], ['T1', 'BOM1', 'M1']], 1);
  const o = RC.buildSheetModel([['任务代码', 'BOM号', '物料代码'], ['T1', 'BOM1', 'M1']], 1);
  eq(RC.guessMatchKey(c, o, manufacturing).names, ['任务代码', 'BOM号', '物料代码']);
});

test('自动配对键优先级：受益项目 + 料品', () => {
  const c = RC.buildSheetModel([['受益项目号', '料品'], ['P1', 'A']], 1);
  const o = RC.buildSheetModel([['受益项目号', '料品'], ['P1', 'A']], 1);
  eq(RC.guessMatchKey(c, o, manufacturing).names, ['受益项目号', '料品']);
});

test('自动配对键优先级：只有料品时回退到料品', () => {
  const c = RC.buildSheetModel([['料品', '数量'], ['A', '1']], 1);
  const o = RC.buildSheetModel([['料品', '数量'], ['A', '1']], 1);
  eq(RC.guessMatchKey(c, o, manufacturing).names, ['料品']);
});

test('配对键不存在时明确报错', () => {
  const c = RC.buildSheetModel([['k', 'v'], ['a', '1']], 1);
  const o = RC.buildSheetModel([['k', 'v'], ['a', '1']], 1);
  throws(() => RC.compareModels(c, o, ['不存在']), /没有同时出现在两侧/);
});

test('数字键：100 与 "100.00" 可以配对上', () => {
  const c = RC.buildSheetModel([['序号', 'v'], [100, 'a']], 1);
  const o = RC.buildSheetModel([['序号', 'v'], ['100.00', 'a']], 1);
  const result = RC.compareModels(c, o, ['序号']);
  eq([result.metrics.pairedRows, result.metrics.missingRows, result.metrics.extraRows], [1, 0, 0]);
});

test('导出：差异明细与缺失/多余行的行数、列数正确', () => {
  const result = compareAoA(
    [['k', 'A', 'B'], ['x', '1', '2'], ['y', '3', '4']],
    [['k', 'A', 'B'], ['x', '1', '9'], ['z', '5', '6']],
    ['k']
  );
  const diff = report.buildDiffExportRows(result, 100);
  eq(diff.total, 1);
  eq(diff.exported, 1);
  eq(diff.rows.length, 2);
  eq(diff.rows[1].slice(0, 4), [2, 2, 'x', 'B']);
  const miss = report.buildMissingExtraExportRows(result, 100);
  eq(miss.total, 2);
  eq(miss.rows.length, 3);
  eq(miss.rows[0].slice(0, 3), ['类型', '原行号', '配对键']);
  eq(miss.rows[1][0], '仅客户有（缺失行）');
  eq(miss.rows[2][0], '仅我方有（多余行）');
  eq(miss.rows[0].length, 3 + result.columns.length);
});

test('导出：超限时正确标记截断', () => {
  const result = compareAoA(
    [['k', 'A'], ['x', '1'], ['y', '2']],
    [['k', 'A'], ['x', '9'], ['y', '8']],
    ['k']
  );
  const diff = report.buildDiffExportRows(result, 1);
  eq([diff.exported, diff.total, diff.truncated], [1, 2, true]);
  const miss = report.buildMissingExtraExportRows(result, 1);
  eq([miss.exported, miss.total, miss.truncated], [0, 0, false]);
});

test('复制单行：包含差异列和两侧值', () => {
  const result = compareAoA(
    [['k', 'A'], ['x', '1']],
    [['k', 'A'], ['x', '2']],
    ['k']
  );
  const text = report.buildRowCopyText(result.rows[0], result.columns);
  eq(text.includes('客户「1」 / 我方「2」'), true, text);
  eq(text.includes('差异 1 处'), true, text);
});

test('摘要：一致 / 有差异两种语气', () => {
  const clean = compareAoA([['k', 'A'], ['x', '1']], [['k', 'A'], ['x', '1']], ['k']);
  eq(report.buildSummaryText(clean).tone, 'success');
  const dirty = compareAoA([['k', 'A'], ['x', '1']], [['k', 'A'], ['x', '2']], ['k']);
  eq(report.buildSummaryText(dirty).tone, 'warning');
});

test('尾部空行：不进入数据行计数', () => {
  const model = RC.buildSheetModel([
    ['k', 'v'],
    ['a', '1'],
    [null, null],
    [null, null]
  ], 1);
  eq(model.rows.length, 1);
});
