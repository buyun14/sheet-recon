(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./core.js'), require('../vendor/sheetjs/xlsx.full.min.js'));
  } else {
    root.ReconcileReport = factory(root.ReconcileCore, root.XLSX);
  }
})(globalThis, function (core, XLSX) {
  'use strict';

  const { text } = core;
  const maxExportRows = 200000;
  const numberFormat = new Intl.NumberFormat('zh-CN');

  function safeCell(v) {
    var s = text(v);
    if (s.length > 32700) s = s.slice(0, 32700) + '…[已截断]';
    return s;
  }
  function buildSummaryText(result) {
    var m = result.metrics;
    var parts = [];
    parts.push('客户 ' + m.customerRows + ' 行 / 我方 ' + m.ourRows + ' 行');
    parts.push('配对 ' + m.pairedRows + ' 行');
    if (m.missingRows) parts.push('缺失 ' + m.missingRows + ' 行');
    if (m.extraRows) parts.push('多余 ' + m.extraRows + ' 行');
    if (m.diffCells) parts.push('不一致 ' + m.diffCells + ' 个单元格');
    if (m.duplicateKeyGroups) parts.push('配对键重复 ' + m.duplicateKeyGroups + ' 组');
    if (m.customerBlankKeys || m.ourBlankKeys) {
      parts.push('空配对键 ' + (m.customerBlankKeys + m.ourBlankKeys) + ' 行');
    }
    if (m.headerOnlyCustomer || m.headerOnlyOurs) {
      parts.push('单侧列 ' + (m.headerOnlyCustomer + m.headerOnlyOurs) + ' 列');
    }
    var risky = m.duplicateKeyGroups > 0 || m.customerBlankKeys > 0 ||
      m.ourBlankKeys > 0 || m.headerOnlyCustomer > 0 || m.headerOnlyOurs > 0;
    var hasDiff = m.diffCells > 0 || m.missingRows > 0 || m.extraRows > 0;
    var clean = !hasDiff && !risky;
    return {
      tone: clean ? 'success' : 'warning',
      text: (clean ? '对账一致：' : (hasDiff ? '发现差异：' : '已配对，但需复核：')) + parts.join('，') + '。'
    };
  }
  function buildRowCopyText(row, columns) {
    var lines = [];
    var statusText = row.status === 'paired' ? '配对行' : (row.status === 'missing' ? '仅客户有（缺失行）' : '仅我方有（多余行）');
    lines.push('【' + statusText + '】客户行 ' + (row.customerRowNo === null ? '—' : row.customerRowNo) + ' / 我方行 ' + (row.ourRowNo === null ? '—' : row.ourRowNo));
    lines.push('配对键：' + (row.keyDisplay || '（空）'));
    if (row.status === 'paired') {
      if (!row.diffCount) {
        lines.push('该行无差异。');
      } else {
        lines.push('差异 ' + row.diffCount + ' 处：');
        row.diffs.forEach(function (d) {
          lines.push('- ' + d.display + '：客户「' + text(d.customerValue) + '」 / 我方「' + text(d.ourValue) + '」');
        });
      }
    } else {
      var vals = row.status === 'missing' ? row.customerValues : row.ourValues;
      lines.push('该行在另一侧不存在，字段如下：');
      columns.forEach(function (c) {
        var idx = row.status === 'missing' ? c.customerIndex : c.ourIndex;
        if (idx < 0 || !vals) return;
        var v = vals[idx];
        if (!isBlank(v)) lines.push('- ' + c.display + '：' + text(v));
      });
    }
    return lines.join('\n');
  }
  function buildOverviewRows(result, meta) {
    meta = meta || {};
    var m = result.metrics;
    var rows = [
      ['对账概览', ''],
      ['生成时间', meta.generatedAt || ''],
      ['客户文件', meta.customerFile || ''],
      ['客户工作表', meta.customerSheet || ''],
      ['客户表头行', meta.customerHeaderRow || ''],
      ['我方文件', meta.ourFile || ''],
      ['我方工作表', meta.ourSheet || ''],
      ['我方表头行', meta.ourHeaderRow || ''],
      ['配对键', result.keyNames.map(function (n) {
        var col = result.columns.filter(function (c) { return c.name === n; })[0];
        return col ? col.display : n;
      }).join(' + ')],
      ['', ''],
      ['指标', '数值'],
      ['客户行数', m.customerRows],
      ['我方行数', m.ourRows],
      ['配对行', m.pairedRows],
      ['缺失行（仅客户有）', m.missingRows],
      ['多余行（仅我方有）', m.extraRows],
      ['不一致单元格', m.diffCells],
      ['', ''],
      ['列级统计（仅配对行）', '', '', '', ''],
      ['列名', '一致数', '不一致数', '合计', '说明']
    ];
    result.columnStats.forEach(function (c) {
      rows.push([
        c.display,
        c.oneSide ? '—' : c.equal,
        c.oneSide ? '—' : c.diff,
        c.oneSide ? '—' : c.total,
        c.oneSide === 'customer' ? '仅客户有该列' : (c.oneSide === 'ours' ? '仅我方有该列' : '')
      ]);
    });
    var warnings = meta.warnings || [];
    rows.push(['', '']);
    if (warnings.length) {
      rows.push(['提示', '']);
      warnings.forEach(function (w) { rows.push([w, '']); });
    } else {
      rows.push(['提示', '无']);
    }
    return rows;
  }
  function buildDiffExportRows(result, limit) {
    var head = ['客户行号', '我方行号', '配对键', '列名', '客户值', '我方值'];
    var out = [];
    var total = result.metrics.diffCells;
    var cap = limit === undefined || limit === null ? total : limit;
    var truncated = false;
    outer:
    for (var i = 0; i < result.rows.length; i++) {
      var row = result.rows[i];
      if (row.status !== 'paired' || !row.diffs.length) continue;
      for (var j = 0; j < row.diffs.length; j++) {
        if (out.length >= cap) { truncated = true; break outer; }
        var d = row.diffs[j];
        out.push([row.customerRowNo, row.ourRowNo, row.keyDisplay, d.display, safeCell(d.customerValue), safeCell(d.ourValue)]);
      }
    }
    return { rows: [head].concat(out), total: total, exported: out.length, truncated: truncated };
  }
  function buildMissingExtraExportRows(result, limit) {
    var head = ['类型', '原行号', '配对键'].concat(result.columns.map(function (c) { return c.display; }));
    var target = result.rows.filter(function (r) { return r.status !== 'paired'; });
    var cap = limit === undefined || limit === null ? target.length : limit;
    var truncated = target.length > cap;
    var out = [];
    for (var i = 0; i < Math.min(target.length, cap); i++) {
      var row = target[i];
      var vals = [];
      result.columns.forEach(function (c) {
        var idx = row.status === 'missing' ? c.customerIndex : c.ourIndex;
        var arr = row.status === 'missing' ? row.customerValues : row.ourValues;
        vals.push(idx >= 0 && arr ? safeCell(arr[idx]) : '');
      });
      out.push([
        row.status === 'missing' ? '仅客户有（缺失行）' : '仅我方有（多余行）',
        row.status === 'missing' ? row.customerRowNo : row.ourRowNo,
        row.keyDisplay
      ].concat(vals));
    }
    return { rows: [head].concat(out), total: target.length, exported: out.length, truncated: truncated };
  }

  function buildWarnings(result) {
    var m = result.metrics;
    var w = [];
    if (m.duplicateKeyGroups) {
      w.push('配对键重复：' + m.duplicateKeyGroups + ' 个键（客户多出 ' + m.duplicateCustomerRows +
        ' 行、我方多出 ' + m.duplicateOurRows + ' 行），已按出现顺序配对，请检查是否错配。');
    }
    if (m.customerBlankKeys || m.ourBlankKeys) {
      w.push('空配对键：客户 ' + m.customerBlankKeys + ' 行、我方 ' + m.ourBlankKeys +
        ' 行，已按缺失/多余行处理。');
    }
    if (m.headerOnlyCustomer || m.headerOnlyOurs) {
      var h = result.headerDiff;
      w.push('表头差异：仅客户有 ' + m.headerOnlyCustomer + ' 列' +
        (h.onlyCustomer.length ? '（' + h.onlyCustomer.slice(0, 8).join('、') + (h.onlyCustomer.length > 8 ? ' 等' : '') + '）' : '') +
        '；仅我方有 ' + m.headerOnlyOurs + ' 列' +
        (h.onlyOurs.length ? '（' + h.onlyOurs.slice(0, 8).join('、') + (h.onlyOurs.length > 8 ? ' 等' : '') + '）' : '') +
        '。');
    }
    if (!m.customerRows || !m.ourRows) {
      w.push('有一侧没有数据行，请检查表头行和工作表。');
    }
    return w;
  }

  function buildReportWorkbook(result, meta = {}, limit = maxExportRows) {
    const warnings = buildWarnings(result).concat(meta.warnings || []);
    const diff = buildDiffExportRows(result, limit);
    const missing = buildMissingExtraExportRows(result, limit);
    const cap = numberFormat.format(limit);
    if (diff.truncated) {
      warnings.push('差异单元格明细超过 ' + cap + ' 行，导出已截断（共 ' + numberFormat.format(diff.total) + ' 行）。');
    }
    if (missing.truncated) {
      warnings.push('缺失/多余行超过 ' + cap + ' 行，导出已截断（共 ' + numberFormat.format(missing.total) + ' 行）。');
    }
    const overview = buildOverviewRows(result, {
      generatedAt: new Date().toLocaleString('zh-CN', { hour12: false }),
      ...meta,
      warnings
    });
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet(overview);
    ws1['!cols'] = [{ wch: 26 }, { wch: 72 }, { wch: 20 }, { wch: 20 }, { wch: 26 }];
    const ws2 = XLSX.utils.aoa_to_sheet(diff.rows);
    ws2['!cols'] = [{ wch: 10 }, { wch: 10 }, { wch: 34 }, { wch: 24 }, { wch: 44 }, { wch: 44 }];
    const ws3 = XLSX.utils.aoa_to_sheet(missing.rows);
    ws3['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 34 }].concat(result.columns.map(() => ({ wch: 18 })));
    XLSX.utils.book_append_sheet(wb, ws1, '概览');
    XLSX.utils.book_append_sheet(wb, ws2, '差异单元格明细');
    XLSX.utils.book_append_sheet(wb, ws3, '缺失与多余行');
    return wb;
  }

  return {
    buildSummaryText, buildRowCopyText, buildOverviewRows, buildDiffExportRows,
    buildMissingExtraExportRows, buildReportWorkbook
  };
});
