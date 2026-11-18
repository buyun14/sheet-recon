(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./profiles.js'));
  } else {
    root.ReconcileCore = factory(root.ReconcileProfiles);
  }
})(globalThis, function (profiles) {
  'use strict';

  var US = '\u001f';

  function isDate(v) {
    return Object.prototype.toString.call(v) === '[object Date]';
  }
  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function pad3(n) { return n < 10 ? '00' + n : (n < 100 ? '0' + n : '' + n); }
  function normalizeDate(d) {
    if (!isDate(d) || isNaN(d.getTime())) return d;
    var h = d.getHours(), m = d.getMinutes();
    // SheetJS 的 cellDates 在部分时区会把纯日期解析到前一天 23:5x；
    // 对“接近午夜”的时间做取整，避免纯日期被当成前一天/带时间。
    if (h === 23 && m >= 50) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
    }
    if (h === 0 && m <= 10) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    return d;
  }
  function formatDate(d) {
    d = normalizeDate(d);
    if (!isDate(d) || isNaN(d.getTime())) return '';
    var s = d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    var h = d.getHours(), m = d.getMinutes(), sec = d.getSeconds(), ms = d.getMilliseconds();
    if (h || m || sec || ms) {
      s += ' ' + pad2(h) + ':' + pad2(m) + ':' + pad2(sec);
      if (ms) s += '.' + pad3(ms);
    }
    return s;
  }
  function text(v) {
    if (v === null || v === undefined) return '';
    if (isDate(v)) return formatDate(v);
    if (typeof v === 'number') {
      if (isNaN(v)) return 'NaN';
      if (!isFinite(v)) return String(v);
      if (Object.is(v, -0)) return '0';
      return String(v);
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }
  function colName(n) {
    var s = '';
    n = Number(n);
    if (!isFinite(n) || n < 0) return '?';
    n = Math.floor(n);
    do {
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return s;
  }
  function normalizeHeader(v) {
    return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
  }
  function headerNames(headerRow, width) {
    var counts = new Map();
    var list = [];
    for (var c = 0; c < width; c++) {
      var raw = normalizeHeader(headerRow ? headerRow[c] : '');
      var display = raw || ('未命名列' + colName(c));
      var n = counts.get(display) || 0;
      var name = n === 0 ? display : display + ' (' + (n + 1) + ')';
      while (list.some(function (h) { return h.name === name; })) {
        n++;
        name = display + ' (' + (n + 1) + ')';
      }
      counts.set(display, n + 1);
      list.push({ name: name, display: name, index: c, empty: !raw });
    }
    return list;
  }
  function rowHasValue(row) {
    if (!row) return false;
    for (var i = 0; i < row.length; i++) {
      if (!isBlank(row[i])) return true;
    }
    return false;
  }
  function scoreHeaderRow(row, profile) {
    var keywords = profile.headerKeywords || [];
    var strongKeywords = profile.strongKeywords || [];
    var hits = 0, strong = 0, nonEmpty = 0;
    if (!row) return { score: 0, hits: 0, strong: 0, nonEmpty: 0 };
    for (var i = 0; i < row.length; i++) {
      var s = normalizeHeader(row[i]);
      if (!s) continue;
      nonEmpty++;
      var lower = s.toLowerCase();
      var matched = '';
      for (var k = 0; k < keywords.length; k++) {
        if (lower.indexOf(keywords[k].toLowerCase()) !== -1) {
          matched = keywords[k];
          break;
        }
      }
      if (matched) {
        hits++;
        if (strongKeywords.indexOf(matched) !== -1) strong++;
      }
    }
    var score = hits * 100 + strong * 10 + Math.min(nonEmpty, 40) * 0.2;
    return { score: Math.round(score * 100) / 100, hits: hits, strong: strong, nonEmpty: nonEmpty };
  }
  function detectHeaderRow(rows, maxScan, profile) {
    profile = profile || profiles.general;
    if (profile.headerStrategy === 'first-nonempty') {
      var index = (rows || []).findIndex(rowHasValue);
      return { row: index < 0 ? 1 : index + 1, score: 0, hits: 0, strong: 0,
        nonEmpty: index < 0 ? 0 : rows[index].filter(function (v) { return !isBlank(v); }).length,
        scanned: index < 0 ? (rows || []).length : index + 1 };
    }
    var limit = Math.min(maxScan || 8, rows ? rows.length : 0);
    var best = null;
    for (var i = 0; i < limit; i++) {
      var r = scoreHeaderRow(rows[i], profile);
      if (!best || r.score > best.score) {
        best = { row: i + 1, score: r.score, hits: r.hits, strong: r.strong, nonEmpty: r.nonEmpty };
      }
    }
    if (!best) best = { row: 1, score: 0, hits: 0, strong: 0, nonEmpty: 0 };
    best.scanned = limit;
    return best;
  }
  function buildSheetModel(rows, headerRow) {
    if (!rows || !rows.length) throw new Error('工作表没有数据');
    var hr = Number(headerRow);
    if (!isFinite(hr) || hr < 1 || Math.floor(hr) !== hr) throw new Error('表头行号必须是正整数');
    if (hr > rows.length) throw new Error('表头行号超出工作表范围（共 ' + rows.length + ' 行）');
    var header = rows[hr - 1] || [];
    var last = rows.length;
    while (last > hr && !rowHasValue(rows[last - 1])) last--;
    var dataRows = [];
    for (var i = hr; i < last; i++) {
      var r = rows[i] || [];
      if (rowHasValue(r)) dataRows.push({ rowNo: i + 1, cells: r });
    }
    var width = header.length;
    for (var d = 0; d < dataRows.length; d++) {
      if (dataRows[d].cells.length > width) width = dataRows[d].cells.length;
    }
    if (!width) throw new Error('表头行没有内容，请指定正确的表头行');
    var mask = new Array(width);
    for (var c = 0; c < width; c++) mask[c] = false;
    for (c = 0; c < header.length; c++) {
      if (!isBlank(header[c])) mask[c] = true;
    }
    for (d = 0; d < dataRows.length; d++) {
      for (c = 0; c < dataRows[d].cells.length; c++) {
        if (!isBlank(dataRows[d].cells[c])) mask[c] = true;
      }
    }
    var keep = [];
    for (c = 0; c < width; c++) if (mask[c]) keep.push(c);
    if (!keep.length) throw new Error('没有找到可比较的列，请检查表头行');
    var allHeaders = headerNames(header, width);
    var headers = keep.map(function (c) { return allHeaders[c]; });
    var modelRows = dataRows.map(function (row) {
      var values = new Array(keep.length);
      for (var j = 0; j < keep.length; j++) {
        var v = row.cells[keep[j]];
        values[j] = v === undefined ? null : v;
      }
      return { rowNo: row.rowNo, values: values };
    });
    return {
      headerRow: hr,
      headers: headers,
      rows: modelRows,
      width: width,
      keep: keep,
      lastRow: last,
      firstDataRow: hr + 1
    };
  }
  function headerIndexMap(headers) {
    var map = new Map();
    for (var i = 0; i < headers.length; i++) map.set(headers[i].name, i);
    return map;
  }
  function buildUnionColumns(customer, ours) {
    var cMap = headerIndexMap(customer.headers);
    var oMap = headerIndexMap(ours.headers);
    var seen = new Set();
    var union = [];
    customer.headers.forEach(function (h) {
      if (seen.has(h.name)) return;
      seen.add(h.name);
      union.push({
        name: h.name, display: h.display,
        customerIndex: cMap.has(h.name) ? cMap.get(h.name) : -1,
        ourIndex: oMap.has(h.name) ? oMap.get(h.name) : -1
      });
    });
    ours.headers.forEach(function (h) {
      if (seen.has(h.name)) return;
      seen.add(h.name);
      union.push({
        name: h.name, display: h.display,
        customerIndex: cMap.has(h.name) ? cMap.get(h.name) : -1,
        ourIndex: oMap.has(h.name) ? oMap.get(h.name) : -1
      });
    });
    return union;
  }
  function commonHeaderNames(customer, ours) {
    var oset = new Set(ours.headers.map(function (h) { return h.name; }));
    return customer.headers.filter(function (h) { return oset.has(h.name); }).map(function (h) { return h.name; });
  }
  function findHeaderName(names, aliases) {
    var i, j, a;
    for (i = 0; i < aliases.length; i++) {
      a = aliases[i].toLowerCase();
      for (j = 0; j < names.length; j++) {
        if (names[j].toLowerCase() === a) return names[j];
      }
    }
    for (i = 0; i < aliases.length; i++) {
      a = aliases[i].toLowerCase();
      for (j = 0; j < names.length; j++) {
        if (names[j].toLowerCase().indexOf(a) !== -1) return names[j];
      }
    }
    return null;
  }
  function listKeyPresets(customer, ours, profile) {
    profile = profile || profiles.general;
    var common = commonHeaderNames(customer, ours);
    var out = [];
    (profile.keyPresets || []).forEach(function (preset) {
      var names = [];
      var ok = true;
      preset.groups.forEach(function (group) {
        var found = findHeaderName(common, group);
        if (!found) { ok = false; return; }
        if (names.indexOf(found) === -1) names.push(found);
      });
      if (ok && names.length) out.push({ names: names, label: preset.label, source: 'preset' });
    });
    return out;
  }
  function keyText(v) {
    if (isBlank(v)) return '';
    if (isDate(v)) return formatDate(v);
    if (typeof v === 'number') return text(v);
    if (typeof v === 'boolean') return text(v);
    var s = String(v).trim();
    if (s.length <= 15 && /^-?(0|[1-9]\d*)(\.\d+)?$/.test(s)) {
      var n = Number(s);
      if (isFinite(n)) return String(n);
    }
    return s;
  }
  function numericStringEq(num, str) {
    var s = String(str).trim();
    if (!s) return false;
    var n = Number(s);
    return isFinite(n) && n === num;
  }
  function cellEquals(a, b) {
    var ab = isBlank(a), bb = isBlank(b);
    if (ab || bb) return ab && bb;
    var ad = isDate(a), bd = isDate(b);
    if (ad || bd) {
      if (ad && bd) return normalizeDate(a).getTime() === normalizeDate(b).getTime();
      return text(a) === text(b);
    }
    var an = typeof a === 'number', bn = typeof b === 'number';
    if (an && bn) return a === b || (isNaN(a) && isNaN(b));
    if (an && typeof b === 'string') return numericStringEq(a, b);
    if (bn && typeof a === 'string') return numericStringEq(b, a);
    if (typeof a === 'boolean' || typeof b === 'boolean') return text(a) === text(b);
    return String(a).trim() === String(b).trim();
  }
  function guessMatchKey(customer, ours, profile) {
    var presets = listKeyPresets(customer, ours, profile);
    if (presets.length) return presets[0];
    var common = commonHeaderNames(customer, ours);
    if (!common.length) return { names: [], label: '两侧没有同名可用的列', source: 'none' };
    var cMap = headerIndexMap(customer.headers);
    var sample = customer.rows.slice(0, 300);
    var best = null;
    common.forEach(function (name) {
      var idx = cMap.get(name);
      var values = [];
      var nonblank = 0;
      sample.forEach(function (row) {
        var v = row.values[idx];
        if (!isBlank(v)) {
          nonblank++;
          values.push(keyText(v));
        }
      });
      if (!sample.length) return;
      var ratio = nonblank / sample.length;
      var uniqRatio = values.length ? (new Set(values).size / values.length) : 0;
      var score = ratio * 0.4 + uniqRatio * 0.6;
      if (ratio >= 0.5 && uniqRatio >= 0.6 && (!best || score > best.score)) {
        best = { name: name, score: score };
      }
    });
    if (best) return { names: [best.name], label: '唯一值较高的列：' + best.name, source: 'unique' };
    return { names: [common[0]], label: '回退到共同列：' + common[0], source: 'fallback' };
  }

  function compareModels(customer, ours, keyNames) {
    keyNames = (keyNames || []).filter(Boolean);
    if (!keyNames.length) throw new Error('请至少选择一个配对键列');
    var cMap = headerIndexMap(customer.headers);
    var oMap = headerIndexMap(ours.headers);
    var missingKeys = keyNames.filter(function (n) { return !cMap.has(n) || !oMap.has(n); });
    if (missingKeys.length) {
      throw new Error('配对键列「' + missingKeys.join('、') + '」没有同时出现在两侧，请重新选择');
    }
    var union = buildUnionColumns(customer, ours);
    var commonCols = union.filter(function (c) { return c.customerIndex >= 0 && c.ourIndex >= 0; });
    var statMap = new Map();
    commonCols.forEach(function (c) {
      statMap.set(c.name, { name: c.name, display: c.display, total: 0, equal: 0, diff: 0 });
    });

    function meta(row, map, idx) {
      var vals = keyNames.map(function (n) { return row.values[map.get(n)]; });
      var blank = vals.every(isBlank);
      return {
        row: row,
        idx: idx,
        vals: vals,
        blank: blank,
        key: vals.map(keyText).join(US),
        display: vals.map(text).join(' | ')
      };
    }

    var customerMeta = customer.rows.map(function (row, idx) { return meta(row, cMap, idx); });
    var ourMeta = ours.rows.map(function (row, idx) { return meta(row, oMap, idx); });

    var cGroups = new Map();
    var oGroups = new Map();
    customerMeta.forEach(function (m) {
      if (m.blank) return;
      var g = cGroups.get(m.key);
      if (!g) { g = { list: [] }; cGroups.set(m.key, g); }
      g.list.push(m);
    });
    ourMeta.forEach(function (m) {
      if (m.blank) return;
      var g = oGroups.get(m.key);
      if (!g) { g = { list: [], pos: 0 }; oGroups.set(m.key, g); }
      g.list.push(m);
    });

    var rows = [];
    var matchedO = new Uint8Array(ours.rows.length);
    var pairedRows = 0, missingRows = 0, extraRows = 0, diffCells = 0;
    var customerBlankKeys = customerMeta.filter(function (m) { return m.blank; }).length;
    var ourBlankKeys = ourMeta.filter(function (m) { return m.blank; }).length;

    customerMeta.forEach(function (m) {
      if (m.blank) {
        missingRows++;
        rows.push({
          id: rows.length, status: 'missing', keyDisplay: m.display, keyValues: m.vals,
          customerRowNo: m.row.rowNo, ourRowNo: null,
          customerValues: m.row.values, ourValues: null, diffs: [], diffCount: 0
        });
        return;
      }
      var g = oGroups.get(m.key);
      var om = null;
      if (g && g.pos < g.list.length) om = g.list[g.pos++];
      if (!om) {
        missingRows++;
        rows.push({
          id: rows.length, status: 'missing', keyDisplay: m.display, keyValues: m.vals,
          customerRowNo: m.row.rowNo, ourRowNo: null,
          customerValues: m.row.values, ourValues: null, diffs: [], diffCount: 0
        });
        return;
      }
      matchedO[om.idx] = 1;
      var diffs = [];
      for (var i = 0; i < commonCols.length; i++) {
        var col = commonCols[i];
        var cv = m.row.values[col.customerIndex];
        var ov = om.row.values[col.ourIndex];
        var st = statMap.get(col.name);
        st.total++;
        if (cellEquals(cv, ov)) {
          st.equal++;
        } else {
          st.diff++;
          diffCells++;
          diffs.push({ name: col.name, display: col.display, customerValue: cv, ourValue: ov });
        }
      }
      pairedRows++;
      rows.push({
        id: rows.length, status: 'paired', keyDisplay: m.display, keyValues: m.vals,
        customerRowNo: m.row.rowNo, ourRowNo: om.row.rowNo,
        customerValues: m.row.values, ourValues: om.row.values,
        diffs: diffs, diffCount: diffs.length
      });
    });

    ourMeta.forEach(function (m) {
      if (!matchedO[m.idx]) {
        extraRows++;
        rows.push({
          id: rows.length, status: 'extra', keyDisplay: m.display, keyValues: m.vals,
          customerRowNo: null, ourRowNo: m.row.rowNo,
          customerValues: null, ourValues: m.row.values, diffs: [], diffCount: 0
        });
      }
    });

    var duplicateKeySet = new Set();
    cGroups.forEach(function (g, k) { if (g.list.length > 1) duplicateKeySet.add(k); });
    oGroups.forEach(function (g, k) { if (g.list.length > 1) duplicateKeySet.add(k); });
    var duplicateCustomerRows = 0, duplicateOurRows = 0;
    cGroups.forEach(function (g) { if (g.list.length > 1) duplicateCustomerRows += g.list.length - 1; });
    oGroups.forEach(function (g) { if (g.list.length > 1) duplicateOurRows += g.list.length - 1; });

    var columns = union.map(function (c) {
      var st = statMap.get(c.name);
      var oneSide = c.customerIndex < 0 ? 'ours' : (c.ourIndex < 0 ? 'customer' : null);
      return {
        name: c.name, display: c.display,
        customerIndex: c.customerIndex, ourIndex: c.ourIndex,
        total: st ? st.total : 0,
        equal: st ? st.equal : 0,
        diff: st ? st.diff : 0,
        oneSide: oneSide
      };
    });
    var columnStats = columns.slice().sort(function (a, b) {
      return b.diff - a.diff || b.total - a.total || a.display.localeCompare(b.display, 'zh-CN');
    });
    var onlyCustomer = union.filter(function (c) { return c.customerIndex >= 0 && c.ourIndex < 0; }).map(function (c) { return c.display; });
    var onlyOurs = union.filter(function (c) { return c.customerIndex < 0 && c.ourIndex >= 0; }).map(function (c) { return c.display; });

    return {
      rows: rows,
      columns: columns,
      columnStats: columnStats,
      keyNames: keyNames.slice(),
      metrics: {
        customerRows: customer.rows.length,
        ourRows: ours.rows.length,
        pairedRows: pairedRows,
        missingRows: missingRows,
        extraRows: extraRows,
        diffCells: diffCells,
        customerBlankKeys: customerBlankKeys,
        ourBlankKeys: ourBlankKeys,
        duplicateKeyGroups: duplicateKeySet.size,
        duplicateCustomerRows: duplicateCustomerRows,
        duplicateOurRows: duplicateOurRows,
        headerOnlyCustomer: onlyCustomer.length,
        headerOnlyOurs: onlyOurs.length
      },
      headerDiff: { onlyCustomer: onlyCustomer, onlyOurs: onlyOurs }
    };
  }

  return {
    isBlank, text, keyText, formatDate, normalizeDate, normalizeHeader,
    detectHeaderRow, buildSheetModel, buildUnionColumns, commonHeaderNames,
    listKeyPresets, guessMatchKey, cellEquals, compareModels
  };
});
