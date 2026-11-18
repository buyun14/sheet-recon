(function (RC, profiles, workbook, report, XLSX) {
  'use strict';

  var numberFormat = new Intl.NumberFormat('zh-CN');
  function fmt(n) { return numberFormat.format(n); }
  var toastTimer = null;
  var sideEls = {};
  var el = {};
  var requestedProfile = new URLSearchParams(window.location.search).get('profile');

  var state = {
    customer: newSide(),
    ours: newSide(),
    profile: Object.hasOwn(profiles, requestedProfile) ? profiles[requestedProfile] : profiles.general,
    keyNames: [],
    keyMode: 'auto',
    keyInfo: null,
    keyError: '',
    result: null,
    filtered: [],
    columnMap: new Map(),
    compareMs: 0,
    compareToken: 0,
    engine: { state: 'loading', label: '加载中…' },
    view: {
      filter: 'all',
      column: null,
      search: '',
      sort: 'default',
      dir: 'asc',
      page: 1,
      pageSize: 100,
      diffColsOnly: false
    },
    pageRows: [],
    displayColumns: [],
    noticeSig: '',
    keyOptionsSig: '',
    searchTimer: null,
    recomputeTimer: null
  };

  function newSide() {
    return {
      file: null,
      workbook: null,
      sheetNames: [],
      sheetName: '',
      sheetMeta: Object.create(null),
      rawRows: null,
      model: null,
      headerRow: null,
      headerHits: 0,
      headerTouched: false,
      headerWarning: '',
      modelError: '',
      modelWarning: '',
      error: '',
      loading: false,
      _sheetSig: ''
    };
  }

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function sideLabel(key) { return key === 'customer' ? '客户文件' : '我方文件'; }
  function headerWarningText(key, row) {
    return sideLabel(key) + '第 ' + row + ' 行没有找到表头关键词（已扫描前 8 行），请手动指定表头行号。';
  }
  function fmtTime(ms) { return ms < 1 ? '<1 ms' : Math.round(ms) + ' ms'; }
  function formatBytes(n) {
    if (!isFinite(n) || n < 0) return '—';
    if (n < 1024) return n + ' B';
    var units = ['KB', 'MB', 'GB'];
    var i = -1;
    do { n = n / 1024; i++; } while (n >= 1024 && i < units.length - 1);
    return (n >= 10 ? n.toFixed(0) : n.toFixed(1)) + ' ' + units[i];
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function timestamp() {
    var d = new Date();
    return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' +
      pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
  }
  function nextPaint(ms) {
    return new Promise(function (resolve) {
      requestAnimationFrame(function () { setTimeout(resolve, ms || 0); });
    });
  }
  function arraysEqual(a, b) {
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  /* ---------------- DOM cache / boot ---------------- */
  function cacheEls() {
    el.engineStatus = $('#engine-status');
    el.notices = $('#notices');
    el.configPanel = $('#config-panel');
    el.profile = $('#profile-select');
    el.results = $('#results');
    el.emptyState = $('#empty-state');
    el.recalcStatus = $('#recalc-status');
    el.summaryLine = $('#summary-line');
    el.mCustomer = $('#m-customer');
    el.mOurs = $('#m-ours');
    el.mPaired = $('#m-paired');
    el.mMissing = $('#m-missing');
    el.mExtra = $('#m-extra');
    el.mDiff = $('#m-diff');
    el.colstatsBody = $('#colstats-body');
    el.colstatNote = $('#colstat-note');
    el.rowFilters = $('#row-filters');
    el.searchInput = $('#search-input');
    el.diffColsOnly = $('#diff-cols-only');
    el.sortSelect = $('#sort-select');
    el.sortDir = $('#sort-dir');
    el.exportBtn = $('#export-btn');
    el.activeFilters = $('#active-filters');
    el.detailHead = $('#detail-head');
    el.detailBody = $('#detail-body');
    el.detailScroll = $('#detail-scroll');
    el.detailCount = $('#detail-count');
    el.pageSize = $('#page-size');
    el.pageInput = $('#page-input');
    el.pageTotal = $('#page-total');
    el.pageFirst = $('#page-first');
    el.pagePrev = $('#page-prev');
    el.pageNext = $('#page-next');
    el.pageLast = $('#page-last');
    el.keyPresets = $('#key-presets');
    el.keyChips = $('#key-chips');
    el.keyOptions = $('#key-options');
    el.keyCount = $('#key-count');
    el.keyModeBadge = $('#key-mode-badge');
    el.overlay = $('#overlay');
    el.overlayTitle = $('#overlay-title');
    el.overlayDetail = $('#overlay-detail');
    el.overlayProgress = $('#overlay-progress');
    el.toast = $('#toast');
    el.copyModal = $('#copy-modal');
    el.copyModalText = $('#copy-modal-text');
    el.copyModalClose = $('#copy-modal-close');

    sideEls.customer = {
      drop: $('#drop-customer'),
      input: $('#file-customer'),
      empty: $('#dz-empty-customer'),
      file: $('#dz-file-customer'),
      name: $('#file-name-customer'),
      sub: $('#file-sub-customer'),
      status: $('#file-status-customer'),
      reselect: $('#reselect-customer'),
      sheet: $('#sheet-customer'),
      header: $('#header-customer'),
      autoHeader: $('#auto-header-customer'),
      hint: $('#header-hint-customer')
    };
    sideEls.ours = {
      drop: $('#drop-ours'),
      input: $('#file-ours'),
      empty: $('#dz-empty-ours'),
      file: $('#dz-file-ours'),
      name: $('#file-name-ours'),
      sub: $('#file-sub-ours'),
      status: $('#file-status-ours'),
      reselect: $('#reselect-ours'),
      sheet: $('#sheet-ours'),
      header: $('#header-ours'),
      autoHeader: $('#auto-header-ours'),
      hint: $('#header-hint-ours')
    };
  }

  function renderEngineStatus() {
    var ready = !!(XLSX && XLSX.version);
    state.engine = { state: ready ? 'ready' : 'error' };
    el.engineStatus.dataset.state = state.engine.state;
    el.engineStatus.querySelector('span').textContent = ready
      ? '解析引擎：SheetJS ' + XLSX.version + ' · 本地'
      : '解析引擎：加载失败';
  }

  /* ---------------- toast / overlay / notices ---------------- */
  function toast(msg, type) {
    el.toast.textContent = msg;
    el.toast.className = 'toast show ' + (type || 'info');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      el.toast.className = 'toast';
    }, type === 'error' ? 5600 : 3400);
  }
  function showOverlay(title, detail, progress) {
    el.overlay.hidden = false;
    document.body.classList.add('busy');
    el.overlayTitle.textContent = title || '正在处理…';
    el.overlayDetail.textContent = detail || '';
    setOverlayProgress(progress);
  }
  function setOverlayProgress(p) {
    if (p === null || p === undefined || isNaN(p)) {
      el.overlayProgress.classList.add('indeterminate');
      el.overlayProgress.firstElementChild.style.width = '';
    } else {
      el.overlayProgress.classList.remove('indeterminate');
      el.overlayProgress.firstElementChild.style.width = Math.max(0, Math.min(100, p)) + '%';
    }
  }
  function hideOverlay() {
    el.overlay.hidden = true;
    document.body.classList.remove('busy');
  }
  function isBusy() { return !el.overlay.hidden; }

  function collectNotices() {
    var out = [];
    if (state.engine.state === 'error') {
      out.push({
        type: 'error',
        title: '解析引擎加载失败',
        detail: '请保留完整的项目目录并刷新页面，确认 vendor/sheetjs/xlsx.full.min.js 文件存在。'
      });
    }
    ['customer', 'ours'].forEach(function (key) {
      var side = state[key];
      var label = sideLabel(key);
      if (side.error) out.push({ type: 'error', title: label + '载入失败', detail: side.error });
      if (side.modelError) out.push({ type: 'error', title: label + '无法建立比对表', detail: side.modelError });
      if (side.headerWarning) out.push({ type: 'error', title: label + '表头未识别', detail: side.headerWarning });
      if (side.modelWarning) out.push({ type: 'warning', title: label + '没有数据行', detail: side.modelWarning });
    });
    if (state.keyError) out.push({ type: 'error', title: '配对键不可用', detail: state.keyError });
    if (state.result) {
      var m = state.result.metrics;
      if (m.duplicateKeyGroups) {
        out.push({
          type: 'warning',
          title: '配对键存在重复',
          detail: '有 ' + m.duplicateKeyGroups + ' 个配对键重复（客户多出 ' + m.duplicateCustomerRows +
            ' 行、我方多出 ' + m.duplicateOurRows + ' 行）。已按出现顺序配对，建议检查是否错配。'
        });
      }
      if (m.customerBlankKeys || m.ourBlankKeys) {
        out.push({
          type: 'warning',
          title: '存在空配对键',
          detail: '客户 ' + m.customerBlankKeys + ' 行、我方 ' + m.ourBlankKeys +
            ' 行的配对键为空，已按缺失/多余行处理。请考虑更换配对键列。'
        });
      }
      if (m.headerOnlyCustomer || m.headerOnlyOurs) {
        var h = state.result.headerDiff;
        function cut(arr) { return arr.slice(0, 8).join('、') + (arr.length > 8 ? ' 等' : ''); }
        out.push({
          type: 'warning',
          title: '两侧表头不完全一致',
          detail: '仅客户有 ' + m.headerOnlyCustomer + ' 列' +
            (h.onlyCustomer.length ? '（' + cut(h.onlyCustomer) + '）' : '') +
            '；仅我方有 ' + m.headerOnlyOurs + ' 列' +
            (h.onlyOurs.length ? '（' + cut(h.onlyOurs) + '）' : '') +
            '。这些列不参与逐单元格比较，但仍会显示在明细中。'
        });
      }
    }
    return out;
  }
  function renderNotices() {
    var notices = collectNotices();
    var sig = JSON.stringify(notices);
    if (sig === state.noticeSig) return;
    state.noticeSig = sig;
    el.notices.replaceChildren();
    notices.forEach(function (n) {
      var box = document.createElement('div');
      box.className = 'notice notice-' + n.type;
      var icon = document.createElement('div');
      icon.className = 'notice-icon';
      icon.textContent = n.type === 'error' ? '!' : (n.type === 'warning' ? '!' : 'i');
      var body = document.createElement('div');
      var title = document.createElement('div');
      title.className = 'notice-title';
      title.textContent = n.title;
      body.appendChild(title);
      if (n.detail) {
        var detail = document.createElement('div');
        detail.className = 'notice-detail';
        detail.textContent = n.detail;
        body.appendChild(detail);
      }
      box.appendChild(icon);
      box.appendChild(body);
      el.notices.appendChild(box);
    });
  }

  /* ---------------- sheet helpers ---------------- */
  function loadSheetFromWorkbook(key, name) {
    var side = state[key];
    side.sheetName = name;
    side.rawRows = null;
    side.model = null;
    side.modelError = '';
    side.modelWarning = '';
    var ws = side.workbook && side.workbook.Sheets[name];
    if (!ws || !ws['!ref']) throw new Error('工作表「' + name + '」为空。');
    var rows = workbook.sheetToRows(ws);
    if (!rows.length) throw new Error('工作表「' + name + '」没有数据。');
    side.rawRows = rows;
    detectSideHeader(key);
  }
  function detectSideHeader(key) {
    var side = state[key];
    var det = RC.detectHeaderRow(side.rawRows, 8, state.profile);
    side.headerRow = det.row;
    side.headerHits = det.hits;
    side.headerTouched = false;
    side.headerWarning = state.profile.headerStrategy === 'keywords' && det.hits === 0
      ? headerWarningText(key, det.row)
      : '';
    buildSideModel(key);
  }
  function buildSideModel(key) {
    var side = state[key];
    side.model = null;
    side.modelError = '';
    side.modelWarning = '';
    if (!side.rawRows || !side.rawRows.length) {
      side.modelError = '工作表没有数据。';
      return false;
    }
    try {
      side.model = RC.buildSheetModel(side.rawRows, side.headerRow);
      if (!side.model.rows.length) side.modelWarning = '表头行之后没有数据行，请确认表头行是否正确。';
      return true;
    } catch (e) {
      side.modelError = e.message || String(e);
      return false;
    }
  }

  /* ---------------- file loading ---------------- */
  function readFileBuffer(file, onProgress) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onprogress = function (e) {
        if (e.lengthComputable) onProgress(e.loaded, e.total);
      };
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(new Error('浏览器读取文件失败。')); };
      reader.readAsArrayBuffer(file);
    });
  }
  function friendlyParseError(e) {
    var msg = String((e && e.message) || e);
    if (/工作表|表头/.test(msg)) return msg;
    if (/password|encrypt|加密/i.test(msg)) return '文件可能已加密或受密码保护，暂不支持。';
    if (/zip|corrupt|end of data|Bad|Invalid|Unsupported/i.test(msg)) {
      return '无法解析文件，可能已损坏或不是有效的 Excel 文件。原始错误：' + msg;
    }
    return '解析失败：' + msg;
  }
  async function handleFile(key, file) {
    if (!file) return;
    if (isBusy()) { toast('正在处理另一个文件，请稍候。', 'error'); return; }
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      toast('仅支持 .xlsx / .xlsm / .xls 文件。', 'error');
      return;
    }
    if (!file.size) { toast('文件为空，请重新选择。', 'error'); return; }

    var side = state[key];
    var label = sideLabel(key);
    if (state.engine.state !== 'ready') {
      toast('解析引擎未加载，无法读取文件。', 'error');
      return;
    }

    side.file = file;
    side.workbook = null;
    side.sheetNames = [];
    side.sheetName = '';
    side.sheetMeta = Object.create(null);
    side.rawRows = null;
    side.model = null;
    side.headerRow = null;
    side.headerHits = 0;
    side.headerTouched = false;
    side.headerWarning = '';
    side.modelError = '';
    side.modelWarning = '';
    side.error = '';
    side.loading = true;
    side._sheetSig = '';
    state.result = null;
    state.filtered = [];
    state.keyMode = 'auto';
    state.keyNames = [];
    state.keyInfo = null;
    state.keyError = '';
    state.compareToken++;
    clearTimeout(state.recomputeTimer);
    state.recomputeTimer = null;
    resetView();
    renderAll();

    var start = performance.now();
    try {
      showOverlay('正在读取' + label + '…', file.name + ' · ' + formatBytes(file.size), 0);
      var buf = await readFileBuffer(file, function (loaded, total) {
        var pct = total ? loaded / total : 0;
        setOverlayProgress(Math.round(pct * 65));
        el.overlayDetail.textContent = file.name + ' · 读取 ' + Math.round(pct * 100) + '%';
      });
      setOverlayProgress(null);
      el.overlayDetail.textContent = '正在解析 Excel（大文件可能需要几十秒，请勿关闭页面）…';
      await nextPaint(50);
      var wb = workbook.readWorkbook(buf);
      side.workbook = wb;
      side.sheetNames = wb.SheetNames.slice();
      side.sheetMeta = workbook.inspectSheets(wb);
      side.sheetName = workbook.pickDefaultSheet(side.sheetNames, side.sheetMeta);
      el.overlayDetail.textContent = '正在读取工作表「' + side.sheetName + '」…';
      setOverlayProgress(85);
      await nextPaint(20);
      loadSheetFromWorkbook(key, side.sheetName);
      side.loading = false;
      var elapsed = performance.now() - start;
      hideOverlay();
      toast(label + '已载入：' + side.sheetNames.length + ' 个工作表 · ' +
        fmt(side.model ? side.model.rows.length : 0) + ' 行数据（' + fmtTime(elapsed) + '）', 'success');
      renderAll();
      scheduleRecompute(0);
    } catch (e) {
      side.loading = false;
      side.model = null;
      side.rawRows = null;
      side.modelError = '';
      side.modelWarning = '';
      side.error = friendlyParseError(e);
      hideOverlay();
      state.result = null;
      state.filtered = [];
      renderAll();
      renderNotices();
      toast(label + '载入失败：' + side.error, 'error');
    }
  }

  async function switchSheet(key, name) {
    if (isBusy()) { renderSheetSelects(); return; }
    var side = state[key];
    if (!side.workbook || !name || name === side.sheetName) return;
    clearTimeout(state.recomputeTimer);
    state.recomputeTimer = null;
    state.compareToken++;
    showOverlay('正在读取工作表…', name, null);
    await nextPaint(40);
    try {
      loadSheetFromWorkbook(key, name);
      side.error = '';
      state.keyError = '';
      state.result = null;
      state.filtered = [];
      resetView();
      hideOverlay();
      renderAll();
      scheduleRecompute(0);
      toast('已切换工作表：' + name, 'success');
    } catch (e) {
      side.model = null;
      side.rawRows = null;
      side.modelError = '';
      side.error = e.message || String(e);
      hideOverlay();
      state.result = null;
      state.filtered = [];
      renderAll();
      renderNotices();
      toast('切换工作表失败：' + side.error, 'error');
    }
  }

  /* ---------------- header / key panel ---------------- */
  function applyHeaderRow(key, value) {
    var side = state[key];
    if (!side.rawRows) return;
    var n = Math.floor(Number(value));
    if (!isFinite(n) || n < 1 || n > side.rawRows.length) {
      toast('表头行号需在 1 - ' + side.rawRows.length + ' 之间。', 'error');
      sideEls[key].header.value = side.headerRow || 1;
      return;
    }
    side.headerRow = n;
    side.headerTouched = true;
    side.headerWarning = '';
    state.keyError = '';
    buildSideModel(key);
    state.result = null;
    state.filtered = [];
    resetView();
    renderAll();
    scheduleRecompute(0);
  }
  function autoHeader(key) {
    var side = state[key];
    if (!side.rawRows) return;
    detectSideHeader(key);
    state.keyError = '';
    state.result = null;
    state.filtered = [];
    resetView();
    renderAll();
    scheduleRecompute(0);
  }

  function getUnionInfo() {
    if (!state.customer.model || !state.ours.model) return null;
    var union = RC.buildUnionColumns(state.customer.model, state.ours.model);
    var common = union.filter(function (c) { return c.customerIndex >= 0 && c.ourIndex >= 0; });
    return { union: union, common: common };
  }
  function rebuildKeyOptions(info) {
    el.keyOptions.replaceChildren();
    if (!info.common.length) {
      var p = document.createElement('p');
      p.className = 'hint';
      p.textContent = '两侧没有同名可用的列，请检查表头行或工作表。';
      el.keyOptions.appendChild(p);
      return;
    }
    info.union.forEach(function (c) {
      var oneSide = c.customerIndex < 0 || c.ourIndex < 0;
      var label = document.createElement('label');
      label.className = 'key-option' + (oneSide ? ' disabled' : '');
      label.title = oneSide
        ? (c.customerIndex < 0 ? '该列仅我方文件有，不能作为配对键' : '该列仅客户文件有，不能作为配对键')
        : c.display;
      var input = document.createElement('input');
      input.type = 'checkbox';
      input.value = c.name;
      input.dataset.keyName = c.name;
      input.checked = state.keyNames.indexOf(c.name) !== -1;
      input.disabled = oneSide;
      var span = document.createElement('span');
      span.textContent = c.display;
      label.appendChild(input);
      label.appendChild(span);
      if (oneSide) {
        var tag = document.createElement('span');
        tag.className = 'side-tag' + (c.customerIndex < 0 ? ' ours' : '');
        tag.textContent = c.customerIndex < 0 ? '仅我方' : '仅客户';
        label.appendChild(tag);
      }
      el.keyOptions.appendChild(label);
    });
  }
  function rebuildKeyPresets(info) {
    el.keyPresets.replaceChildren();
    var autoBtn = document.createElement('button');
    autoBtn.type = 'button';
    autoBtn.className = 'preset-btn';
    autoBtn.dataset.auto = '1';
    autoBtn.textContent = '自动猜测';
    autoBtn.title = '根据表头名重新自动推荐';
    el.keyPresets.appendChild(autoBtn);
    RC.listKeyPresets(state.customer.model, state.ours.model, state.profile).forEach(function (preset) {
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'preset-btn';
      btn.textContent = preset.label;
      btn.title = '选择：' + preset.names.join(' + ');
      btn.__names = preset.names.slice();
      el.keyPresets.appendChild(btn);
    });
  }
  function syncKeySelection(info) {
    if (!info) info = getUnionInfo();
    if (!info) return;
    var unionMap = new Map();
    info.union.forEach(function (c) { unionMap.set(c.name, c); });

    $all('input[data-key-name]', el.keyOptions).forEach(function (input) {
      input.checked = state.keyNames.indexOf(input.dataset.keyName) !== -1;
    });

    el.keyChips.replaceChildren();
    if (!state.keyNames.length) {
      var hint = document.createElement('span');
      hint.className = 'hint';
      hint.textContent = '请至少选择一个配对键列';
      el.keyChips.appendChild(hint);
    } else {
      state.keyNames.forEach(function (name) {
        var col = unionMap.get(name);
        var chip = document.createElement('span');
        chip.className = 'key-chip';
        var span = document.createElement('span');
        span.textContent = col ? col.display : name;
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = '×';
        btn.title = '移除该配对键';
        btn.dataset.key = name;
        btn.setAttribute('aria-label', '移除配对键 ' + (col ? col.display : name));
        chip.appendChild(span);
        chip.appendChild(btn);
        el.keyChips.appendChild(chip);
      });
    }

    $all('button', el.keyPresets).forEach(function (btn) {
      var active;
      if (btn.dataset.auto) {
        active = state.keyMode === 'auto';
      } else {
        active = state.keyMode === 'manual' && btn.__names && arraysEqual(btn.__names, state.keyNames);
      }
      btn.classList.toggle('active', !!active);
    });

    el.keyModeBadge.textContent = state.keyMode === 'auto' ? '自动' : '手动';
    el.keyModeBadge.className = 'mini-badge' + (state.keyMode === 'auto' ? '' : ' manual');
    el.keyCount.textContent = '（' + info.union.length + ' 列，可选 ' + info.common.length + ' 列）';
  }
  function renderKeyPanel() {
    var info = getUnionInfo();
    if (!info) {
      state.keyOptionsSig = '';
      el.keyPresets.replaceChildren();
      el.keyChips.replaceChildren();
      el.keyOptions.replaceChildren();
      var p = document.createElement('p');
      p.className = 'hint';
      p.textContent = '两侧文件都载入后显示可勾选的列。';
      el.keyOptions.appendChild(p);
      el.keyCount.textContent = '';
      el.keyModeBadge.textContent = state.keyMode === 'auto' ? '自动' : '手动';
      el.keyModeBadge.className = 'mini-badge' + (state.keyMode === 'auto' ? '' : ' manual');
      return;
    }
    var sig = state.profile.id + ':' + info.union.map(function (c) {
      return c.name + ':' + c.customerIndex + ':' + c.ourIndex;
    }).join('|');
    if (sig !== state.keyOptionsSig) {
      state.keyOptionsSig = sig;
      rebuildKeyOptions(info);
      rebuildKeyPresets(info);
    }
    syncKeySelection(info);
  }

  /* ---------------- compare ---------------- */
  function resetView() {
    state.view.filter = 'all';
    state.view.column = null;
    state.view.search = '';
    state.view.sort = 'default';
    state.view.dir = 'asc';
    state.view.page = 1;
    state.view.pageSize = 100;
    state.view.diffColsOnly = false;
    if (el.searchInput) el.searchInput.value = '';
    if (el.diffColsOnly) el.diffColsOnly.checked = false;
    if (el.pageSize) el.pageSize.value = '100';
  }
  function setRecalcStatus(text, busy) {
    el.recalcStatus.textContent = text;
    el.recalcStatus.classList.toggle('busy', !!busy);
  }
  function scheduleRecompute(delay) {
    clearTimeout(state.recomputeTimer);
    if (el.exportBtn) el.exportBtn.disabled = true;
    setRecalcStatus('计算中…', true);
    state.recomputeTimer = setTimeout(function () { runCompare(); }, delay === undefined ? 120 : delay);
  }
  function failCompare(message) {
    state.keyError = message || '';
    state.result = null;
    state.filtered = [];
    state.columnMap = new Map();
    renderResults();
    renderNotices();
    setRecalcStatus('需要调整');
  }
  async function runCompare() {
    clearTimeout(state.recomputeTimer);
    state.recomputeTimer = null;
    var token = ++state.compareToken;
    state.keyError = '';

    if (!state.customer.model || !state.ours.model) {
      state.result = null;
      state.filtered = [];
      state.columnMap = new Map();
      renderResults();
      renderNotices();
      setRecalcStatus('等待两侧文件');
      return;
    }
    if (state.customer.headerWarning || state.ours.headerWarning) {
      state.result = null;
      state.filtered = [];
      state.columnMap = new Map();
      renderResults();
      renderNotices();
      setRecalcStatus('需要指定表头');
      return;
    }
    var info = getUnionInfo();
    if (!info || !info.common.length) {
      failCompare('两侧表头没有同名可用的列（客户第 ' + (state.customer.headerRow || '?') +
        ' 行 / 我方第 ' + (state.ours.headerRow || '?') + ' 行），无法配对。请检查表头行是否选对。');
      return;
    }

    if (state.keyMode === 'auto') {
      var guessed = RC.guessMatchKey(state.customer.model, state.ours.model, state.profile);
      state.keyNames = guessed.names.slice();
      state.keyInfo = guessed;
    } else if (state.keyMode === 'manual' && state.keyNames.length) {
      var commonSet = new Set(info.common.map(function (c) { return c.name; }));
      var invalid = state.keyNames.filter(function (n) { return !commonSet.has(n); });
      if (invalid.length) {
        state.keyMode = 'auto';
        state.keyInfo = null;
        var reguess = RC.guessMatchKey(state.customer.model, state.ours.model, state.profile);
        state.keyNames = reguess.names.slice();
        state.keyInfo = reguess;
        toast('原配对键列在新工作表中不存在，已重新自动推荐。', 'info');
      }
    }

    if (!state.keyNames.length) {
      syncKeySelection(info);
      failCompare('请至少勾选一个配对键列。');
      return;
    }
    var commonNames = new Set(info.common.map(function (c) { return c.name; }));
    var bad = state.keyNames.filter(function (n) { return !commonNames.has(n); });
    if (bad.length) {
      syncKeySelection(info);
      failCompare('配对键列「' + bad.join('、') + '」没有同时出现在两侧，请重新选择。');
      return;
    }

    syncKeySelection(info);
    setRecalcStatus('计算中…', true);
    await nextPaint(0);
    if (token !== state.compareToken) return;

    try {
      var t0 = performance.now();
      var result = RC.compareModels(state.customer.model, state.ours.model, state.keyNames);
      result.rows.forEach(function (row, idx) { row.__id = idx; });
      state.compareMs = Math.max(0, performance.now() - t0);
      state.result = result;
      state.columnMap = new Map(result.columns.map(function (c) { return [c.name, c]; }));
      if (state.view.column && !state.columnMap.has(state.view.column)) state.view.column = null;
      if (state.view.sort.indexOf('column:') === 0) {
        var sortName = state.view.sort.slice(7);
        if (!state.columnMap.has(sortName)) state.view.sort = 'default';
      }
      state.view.page = 1;
      state.keyError = '';
      renderKeyPanel();
      renderResults();
      renderNotices();
      setRecalcStatus('已重算 · ' + fmt(result.metrics.pairedRows) + ' 对');
    } catch (e) {
      if (token !== state.compareToken) return;
      failCompare(e.message || String(e));
    }
  }

  /* ---------------- rendering: files / selectors ---------------- */
  function renderFileCards() {
    ['customer', 'ours'].forEach(function (key) {
      var side = state[key];
      var ui = sideEls[key];
      var hasFile = !!side.file;
      ui.empty.hidden = hasFile;
      ui.file.hidden = !hasFile;
      ui.drop.classList.toggle('has-file', hasFile);
      if (!hasFile) {
        ui.drop.setAttribute('aria-label', sideLabel(key) + '：拖拽到此处或点击选择');
        return;
      }
      ui.name.textContent = side.file.name;
      ui.name.title = side.file.name;
      var sub = formatBytes(side.file.size);
      if (side.workbook) sub += ' · ' + side.sheetNames.length + ' 个工作表';
      if (side.sheetName) sub += ' · ' + side.sheetName;
      ui.sub.textContent = sub;
      ui.status.className = 'file-status';
      if (side.loading) {
        ui.status.classList.add('loading');
        ui.status.textContent = '读取中…';
      } else if (side.error) {
        ui.status.classList.add('error');
        ui.status.textContent = '载入失败';
        ui.status.title = side.error;
      } else if (side.modelError) {
        ui.status.classList.add('error');
        ui.status.textContent = '表头/数据有误';
        ui.status.title = side.modelError;
      } else if (side.workbook) {
        ui.status.textContent = '已载入 · ' + fmt(side.model ? side.model.rows.length : 0) + ' 行数据';
        ui.status.title = '';
      } else {
        ui.status.textContent = '已选择文件';
      }
      ui.drop.setAttribute('aria-label', sideLabel(key) + '：' + side.file.name + '，点击可重新选择');
    });
  }
  function renderSheetSelects() {
    ['customer', 'ours'].forEach(function (key) {
      var side = state[key];
      var sel = sideEls[key].sheet;
      if (!side.workbook) {
        sel.disabled = true;
        sel.replaceChildren();
        var opt = document.createElement('option');
        opt.textContent = '请先载入' + sideLabel(key);
        sel.appendChild(opt);
        return;
      }
      var sig = side.sheetNames.join('\u001f') + '|' + side.sheetNames.map(function (n) {
        return side.sheetMeta[n] ? '1' : '0';
      }).join('');
      if (sig !== side._sheetSig) {
        side._sheetSig = sig;
        sel.replaceChildren();
        side.sheetNames.forEach(function (name) {
          var option = document.createElement('option');
          option.value = name;
          option.textContent = name + (side.sheetMeta[name] ? '' : '（空）');
          sel.appendChild(option);
        });
      }
      sel.disabled = false;
      if (side.sheetName) sel.value = side.sheetName;
    });
  }
  function renderHeaderControls() {
    ['customer', 'ours'].forEach(function (key) {
      var side = state[key];
      var ui = sideEls[key];
      var has = !!side.rawRows;
      ui.header.disabled = !has;
      ui.autoHeader.disabled = !has;
      ui.header.max = has ? side.rawRows.length : 1;
      var value = side.headerRow || 1;
      if (document.activeElement !== ui.header && ui.header.value !== String(value)) {
        ui.header.value = value;
      }
      ui.hint.className = 'header-hint';
      if (!has) {
        ui.hint.textContent = sideLabel(key) + '：载入后自动识别';
      } else if (side.headerWarning) {
        ui.hint.classList.add('warn');
        ui.hint.textContent = sideLabel(key) + '：第 ' + value + ' 行未命中表头关键词，请手动指定';
      } else if (side.headerTouched) {
        ui.hint.textContent = sideLabel(key) + '：手动指定第 ' + value + ' 行 · ' + fmt(side.model ? side.model.rows.length : 0) + ' 行数据';
      } else {
        ui.hint.classList.add('ok');
        var detail = state.profile.headerStrategy === 'keywords' ? '（命中 ' + side.headerHits + ' 个关键词）' : '';
        ui.hint.textContent = sideLabel(key) + '：自动识别第 ' + value + ' 行' + detail +
          ' · ' + fmt(side.model ? side.model.rows.length : 0) + ' 行数据';
      }
      if (side.modelError) {
        ui.hint.classList.add('warn');
        ui.hint.textContent = sideLabel(key) + '：' + side.modelError;
      }
    });
  }

  /* ---------------- rendering: results ---------------- */
  function renderAll() {
    el.configPanel.hidden = !(state.customer.file || state.ours.file);
    renderFileCards();
    renderSheetSelects();
    renderHeaderControls();
    renderKeyPanel();
    renderResults();
    renderNotices();
  }
  function renderEmptyState() {
    var c = state.customer;
    var o = state.ours;
    var title = '先放入两个 Excel 文件';
    var sub = '左侧客户文件，右侧我方文件。所有解析都在本机浏览器完成，不会上传。';
    if (c.file || o.file) {
      if (!c.file || !o.file) {
        var loaded = c.file ? '客户文件' : '我方文件';
        var waiting = c.file ? '我方文件' : '客户文件';
        title = '已载入' + loaded + '，等待' + waiting;
        sub = '把另一个文件拖入对应的区域即可开始比对。';
      } else if (c.loading || o.loading) {
        title = '正在解析文件…';
        sub = '大文件可能需要几十秒，请勿关闭页面。';
      } else if (state.engine.state === 'error') {
        title = '解析引擎未加载';
        sub = '请检查网络，或把 xlsx.full.min.js 放到本 HTML 同目录后刷新。';
      } else if (c.headerWarning || o.headerWarning) {
        title = '请指定表头行';
        sub = (c.headerWarning || o.headerWarning);
      } else if (c.modelError || o.modelError) {
        title = '表头或工作表有误';
        sub = c.modelError || o.modelError;
      } else if (state.keyError) {
        title = '配对键需要调整';
        sub = state.keyError;
      } else if (c.model && o.model) {
        title = '正在准备比对结果…';
        sub = '正在按配对键匹配行并逐列比较。';
      } else {
        title = '请根据上方提示调整设置';
        sub = '检查工作表、表头行和配对键后会自动重算。';
      }
    }
    var h2 = el.emptyState.querySelector('h2');
    var p = el.emptyState.querySelector('p');
    if (h2) h2.textContent = title;
    if (p) p.textContent = sub;
  }
  function renderResults() {
    var result = state.result;
    el.configPanel.hidden = !(state.customer.file || state.ours.file);
    el.results.hidden = !result;
    el.emptyState.hidden = !!result;
    if (!result) {
      el.exportBtn.disabled = true;
      renderEmptyState();
      return;
    }
    el.exportBtn.disabled = false;
    renderSummary(result);
    renderMetrics(result);
    renderSortOptions(result);
    renderColumnStats(result);
    renderDetail();
  }
  function keyDisplayNames() {
    var info = getUnionInfo();
    var map = info ? new Map(info.union.map(function (c) { return [c.name, c.display]; })) : new Map();
    return state.keyNames.map(function (n) { return map.get(n) || n; });
  }
  function renderSummary(result) {
    var s = report.buildSummaryText(result);
    el.summaryLine.className = 'summary-line ' + s.tone;
    el.summaryLine.replaceChildren();
    var icon = document.createElement('span');
    icon.className = 'summary-icon';
    icon.textContent = s.tone === 'success' ? '✓' : '!';
    var body = document.createElement('div');
    var main = document.createElement('span');
    main.textContent = s.text;
    var sub = document.createElement('span');
    sub.className = 'summary-sub';
    sub.textContent = '配对键：' + keyDisplayNames().join(' + ') +
      ' · 用时 ' + fmtTime(state.compareMs) +
      ' · 客户「' + (state.customer.sheetName || '—') + '」 / 我方「' + (state.ours.sheetName || '—') + '」';
    body.appendChild(main);
    body.appendChild(sub);
    el.summaryLine.appendChild(icon);
    el.summaryLine.appendChild(body);
  }
  function renderMetrics(result) {
    var m = result.metrics;
    el.mCustomer.textContent = fmt(m.customerRows);
    el.mOurs.textContent = fmt(m.ourRows);
    el.mPaired.textContent = fmt(m.pairedRows);
    el.mMissing.textContent = fmt(m.missingRows);
    el.mExtra.textContent = fmt(m.extraRows);
    el.mDiff.textContent = fmt(m.diffCells);
    el.mDiff.classList.toggle('danger', m.diffCells > 0);
    el.mDiff.classList.toggle('ok', m.diffCells === 0);
    el.mPaired.classList.toggle('ok', m.pairedRows > 0 && m.missingRows === 0 && m.extraRows === 0 && m.diffCells === 0);
  }
  function renderSortOptions(result) {
    var sel = el.sortSelect;
    sel.replaceChildren();
    function add(value, text) {
      var opt = document.createElement('option');
      opt.value = value;
      opt.textContent = text;
      sel.appendChild(opt);
    }
    add('default', '默认顺序');
    add('diff', '差异优先');
    add('customerRow', '客户行号');
    add('ourRow', '我方行号');
    add('key', '配对键');
    var group = document.createElement('optgroup');
    group.label = '按列排序';
    result.columns.forEach(function (c) {
      var opt = document.createElement('option');
      opt.value = 'column:' + c.name;
      opt.textContent = '按列：' + c.display;
      group.appendChild(opt);
    });
    sel.appendChild(group);
    var found = false;
    for (var i = 0; i < sel.options.length; i++) {
      if (sel.options[i].value === state.view.sort) { found = true; break; }
    }
    if (!found) state.view.sort = 'default';
    sel.value = state.view.sort;
    syncSortControls();
  }
  function syncSortControls() {
    if (el.sortSelect.value !== state.view.sort) el.sortSelect.value = state.view.sort;
    el.sortDir.textContent = state.view.dir === 'asc' ? '↑' : '↓';
    el.sortDir.title = state.view.dir === 'asc' ? '当前升序，点击降序' : '当前降序，点击升序';
    el.sortDir.setAttribute('aria-label', el.sortDir.title);
  }
  function renderColumnStats(result) {
    el.colstatsBody.replaceChildren();
    var diffCols = 0;
    result.columnStats.forEach(function (c) {
      if (c.diff > 0) diffCols++;
      var tr = document.createElement('tr');
      tr.className = 'colstat' + (state.view.column === c.name ? ' active' : '');
      tr.dataset.col = c.name;
      if (!c.oneSide) {
        tr.tabIndex = 0;
        tr.setAttribute('role', 'button');
        tr.setAttribute('aria-label', '按列 ' + c.display + ' 筛选差异行');
      }
      tr.title = c.oneSide
        ? (c.oneSide === 'customer' ? '该列仅客户文件有' : '该列仅我方文件有')
        : ('一致 ' + c.equal + ' · 不一致 ' + c.diff + ' · 合计 ' + c.total);

      var nameTd = document.createElement('td');
      var nameWrap = document.createElement('div');
      nameWrap.className = 'colstat-name';
      var nameSpan = document.createElement('span');
      nameSpan.className = 'name';
      nameSpan.textContent = c.display;
      nameWrap.appendChild(nameSpan);
      if (c.oneSide) {
        var badge = document.createElement('span');
        badge.className = 'badge' + (c.oneSide === 'ours' ? ' ours' : '');
        badge.textContent = c.oneSide === 'customer' ? '仅客户' : '仅我方';
        nameWrap.appendChild(badge);
      }
      nameTd.appendChild(nameWrap);
      tr.appendChild(nameTd);

      function countTd(text, cls) {
        var td = document.createElement('td');
        if (cls) td.className = cls;
        td.textContent = text;
        tr.appendChild(td);
      }
      if (c.oneSide) {
        countTd('—');
        countTd('—');
        countTd('—');
        var barTd = document.createElement('td');
        barTd.className = 'bar-cell';
        var barWrap = document.createElement('div');
        barWrap.className = 'bar-wrap';
        var bar = document.createElement('div');
        bar.className = 'bar';
        var barSpan = document.createElement('span');
        barSpan.className = 'zero';
        var rate = document.createElement('span');
        rate.className = 'rate';
        rate.textContent = '—';
        bar.appendChild(barSpan);
        barWrap.appendChild(bar);
        barWrap.appendChild(rate);
        barTd.appendChild(barWrap);
        tr.appendChild(barTd);
      } else {
        countTd(fmt(c.equal), 'equal-n');
        countTd(fmt(c.diff), c.diff ? 'diff-n' : '');
        countTd(fmt(c.total));
        var barTd2 = document.createElement('td');
        barTd2.className = 'bar-cell';
        var barWrap2 = document.createElement('div');
        barWrap2.className = 'bar-wrap';
        var bar2 = document.createElement('div');
        bar2.className = 'bar';
        var barSpan2 = document.createElement('span');
        var pct = c.total ? (c.diff / c.total) * 100 : 0;
        barSpan2.style.width = (pct ? Math.max(pct, 1.5) : 0) + '%';
        if (!pct) barSpan2.className = 'zero';
        var rate2 = document.createElement('span');
        rate2.className = 'rate';
        rate2.textContent = pct.toFixed(1) + '%';
        bar2.appendChild(barSpan2);
        barWrap2.appendChild(bar2);
        barWrap2.appendChild(rate2);
        barTd2.appendChild(barWrap2);
        tr.appendChild(barTd2);
      }
      el.colstatsBody.appendChild(tr);
    });
    el.colstatNote.textContent = '共 ' + result.columns.length + ' 列 · 有差异 ' + diffCols + ' 列';
  }

  /* ---------------- rendering: detail ---------------- */
  function getDisplayColumns() {
    var result = state.result;
    if (!result) return [];
    var cols = result.columns;
    if (state.view.diffColsOnly) {
      var keys = new Set(result.keyNames);
      cols = cols.filter(function (c) { return keys.has(c.name) || c.diff > 0 || c.oneSide; });
    }
    if (state.view.column) {
      var visible = new Set(cols.map(function (c) { return c.name; }));
      visible.add(state.view.column);
      cols = result.columns.filter(function (c) { return visible.has(c.name); });
    }
    if (!cols.length) {
      var keySet = new Set(result.keyNames);
      cols = result.columns.filter(function (c) { return keySet.has(c.name); });
      if (!cols.length) cols = result.columns.slice(0, 1);
    }
    return cols;
  }
  function renderDetail() {
    var result = state.result;
    if (!result) {
      state.filtered = [];
      state.pageRows = [];
      el.detailHead.replaceChildren();
      el.detailBody.replaceChildren();
      return;
    }
    var cols = getDisplayColumns();
    state.displayColumns = cols;
    renderDetailHead(cols);
    refreshFiltered();
    renderPage();
    updateFilterButtons();
    renderActiveFilters();
    el.detailScroll.scrollTop = 0;
    if (state.view.column) {
      scrollDetailToColumn(state.view.column);
    } else if (state.view.filter === 'diff') {
      scrollDetailToFirstDiff();
    }
  }
  function scrollDetailToColumn(name) {
    var head = el.detailHead.querySelector('tr');
    if (!head) return;
    var stickyWidth = 0;
    ['sticky-1', 'sticky-2', 'sticky-3'].forEach(function (cls) {
      var cell = head.querySelector('th.' + cls);
      if (cell && getComputedStyle(cell).position === 'sticky') stickyWidth += cell.offsetWidth;
    });
    var ths = head.querySelectorAll('th.col-head');
    for (var i = 0; i < ths.length; i++) {
      var btn = ths[i].querySelector('button[data-col]');
      if (btn && btn.dataset.col === name) {
        el.detailScroll.scrollLeft = Math.max(0, ths[i].offsetLeft - stickyWidth - 16);
        return;
      }
    }
  }
  function scrollDetailToFirstDiff() {
    var target = null;
    for (var i = 0; i < state.pageRows.length && !target; i++) {
      var row = state.pageRows[i];
      if (row.status === 'paired' && row.diffs.length) target = row.diffs[0];
    }
    if (!target) return;
    scrollDetailToColumn(target.name);
  }
  function renderDetailHead(cols) {
    el.detailHead.replaceChildren();
    var tr = document.createElement('tr');
    var th1 = document.createElement('th');
    th1.className = 'sticky-1';
    th1.scope = 'col';
    th1.innerHTML = '<div class="th-inner">行</div>';
    var th2 = document.createElement('th');
    th2.className = 'sticky-2';
    th2.scope = 'col';
    th2.innerHTML = '<div class="th-inner">状态</div>';
    var th3 = document.createElement('th');
    th3.className = 'sticky-3';
    th3.scope = 'col';
    th3.innerHTML = '<div class="th-inner">配对键</div>';
    tr.appendChild(th1);
    tr.appendChild(th2);
    tr.appendChild(th3);
    cols.forEach(function (c) {
      var th = document.createElement('th');
      th.className = 'col-head' + (state.view.column === c.name ? ' active-col' : '');
      th.scope = 'col';
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'col-sort';
      btn.dataset.col = c.name;
      btn.title = '按「' + c.display + '」排序';
      var name = document.createElement('span');
      name.className = 'name';
      name.textContent = c.display;
      btn.appendChild(name);
      if (state.view.sort === 'column:' + c.name) {
        var arrow = document.createElement('span');
        arrow.className = 'arrow';
        arrow.textContent = state.view.dir === 'asc' ? '▲' : '▼';
        btn.appendChild(arrow);
      }
      th.appendChild(btn);
      var sub = document.createElement('div');
      sub.className = 'col-sub';
      var left = document.createElement('span');
      left.className = 'c';
      left.textContent = '客户';
      var right = document.createElement('span');
      right.className = 'o';
      right.textContent = '我方';
      sub.appendChild(left);
      sub.appendChild(right);
      th.appendChild(sub);
      if (c.oneSide) {
        var tag = document.createElement('span');
        tag.className = 'col-side-tag' + (c.oneSide === 'ours' ? ' ours' : '');
        tag.textContent = c.oneSide === 'customer' ? '仅客户有' : '仅我方有';
        th.appendChild(tag);
      }
      tr.appendChild(th);
    });
    el.detailHead.appendChild(tr);
  }
  function rowValueForSort(row, name) {
    var col = state.columnMap.get(name);
    if (!col) return null;
    if (col.customerIndex >= 0 && row.customerValues) return row.customerValues[col.customerIndex];
    if (col.ourIndex >= 0 && row.ourValues) return row.ourValues[col.ourIndex];
    return null;
  }
  function compareCellValues(a, b) {
    var ab = RC.isBlank(a), bb = RC.isBlank(b);
    if (ab && bb) return 0;
    if (ab) return 1;
    if (bb) return -1;
    var an = typeof a === 'number', bn = typeof b === 'number';
    if (an && bn) return a - b;
    if (an && typeof b === 'string' && b.trim() !== '' && isFinite(Number(b))) return a - Number(b);
    if (bn && typeof a === 'string' && a.trim() !== '' && isFinite(Number(a))) return Number(a) - b;
    return RC.text(a).localeCompare(RC.text(b), 'zh-CN', { numeric: true, sensitivity: 'base' });
  }
  function severity(row) {
    if (row.status === 'missing') return 0;
    if (row.status === 'extra') return 1;
    if (row.diffCount > 0) return 2;
    return 3;
  }
  function rowNum(v) { return v == null ? 1000000000000000 : v; }
  function compareRows(a, b) {
    var dir = state.view.dir === 'desc' ? -1 : 1;
    var r = 0;
    if (state.view.sort === 'diff') {
      r = severity(a) - severity(b) || b.diffCount - a.diffCount;
    } else if (state.view.sort === 'customerRow') {
      r = rowNum(a.customerRowNo) - rowNum(b.customerRowNo);
    } else if (state.view.sort === 'ourRow') {
      r = rowNum(a.ourRowNo) - rowNum(b.ourRowNo);
    } else if (state.view.sort === 'key') {
      r = String(a.keyDisplay || '').localeCompare(String(b.keyDisplay || ''), 'zh-CN', { numeric: true, sensitivity: 'base' });
    } else if (state.view.sort.indexOf('column:') === 0) {
      var name = state.view.sort.slice(7);
      r = compareCellValues(rowValueForSort(a, name), rowValueForSort(b, name));
    } else {
      r = rowNum(a.customerRowNo) - rowNum(b.customerRowNo) ||
        rowNum(a.ourRowNo) - rowNum(b.ourRowNo);
    }
    if (r === 0) {
      r = rowNum(a.customerRowNo) - rowNum(b.customerRowNo) ||
        rowNum(a.ourRowNo) - rowNum(b.ourRowNo) ||
        a.id - b.id;
    }
    return r * dir;
  }
  function rowHaystack(row) {
    if (row.__hay !== undefined) return row.__hay;
    var parts = [row.keyDisplay || ''];
    if (row.customerValues) {
      for (var i = 0; i < row.customerValues.length; i++) parts.push(RC.text(row.customerValues[i]));
    }
    if (row.ourValues) {
      for (var j = 0; j < row.ourValues.length; j++) parts.push(RC.text(row.ourValues[j]));
    }
    row.__hay = parts.join('\n').toLowerCase();
    return row.__hay;
  }
  function refreshFiltered() {
    var result = state.result;
    if (!result) { state.filtered = []; return; }
    var q = state.view.search.trim().toLowerCase();
    var filter = state.view.filter;
    var column = state.view.column;
    var rows = result.rows.filter(function (row) {
      if (filter === 'diff' && !(row.status !== 'paired' || row.diffCount > 0)) return false;
      if (filter === 'missing' && row.status !== 'missing') return false;
      if (filter === 'extra' && row.status !== 'extra') return false;
      if (filter === 'paired' && row.status !== 'paired') return false;
      if (column) {
        if (row.status !== 'paired') return false;
        if (!row.diffs.some(function (d) { return d.name === column; })) return false;
      }
      if (q && rowHaystack(row).indexOf(q) === -1) return false;
      return true;
    });
    rows.sort(compareRows);
    state.filtered = rows;
  }
  function makeHalf(side, value, options) {
    options = options || {};
    var div = document.createElement('div');
    div.className = 'half half-' + side +
      (options.diff ? ' diff' : '') +
      (options.absent ? ' absent' : '');
    div.textContent = options.absent ? '—' : RC.text(value);
    if (options.title) div.title = options.title;
    return div;
  }
  function renderPage() {
    var result = state.result;
    if (!result) return;
    var cols = state.displayColumns;
    var total = state.filtered.length;
    var pageSize = state.view.pageSize;
    var pages = Math.max(1, Math.ceil(total / pageSize));
    if (state.view.page > pages) state.view.page = pages;
    if (state.view.page < 1) state.view.page = 1;
    var start = (state.view.page - 1) * pageSize;
    var pageRows = state.filtered.slice(start, start + pageSize);
    state.pageRows = pageRows;

    el.detailBody.replaceChildren();
    if (!total) {
      var noTr = document.createElement('tr');
      noTr.className = 'no-rows';
      var noTd = document.createElement('td');
      noTd.colSpan = cols.length + 3;
      var msg = '没有符合条件的行。';
      if (state.view.search.trim()) msg = '没有匹配「' + state.view.search.trim() + '」的行。';
      else if (state.view.column) {
        var col = state.columnMap.get(state.view.column);
        msg = '列「' + (col ? col.display : state.view.column) + '」在已配对行中没有差异。';
      } else if (state.view.filter === 'diff') msg = '没有差异行。';
      else if (state.view.filter === 'missing') msg = '没有仅客户有的行。';
      else if (state.view.filter === 'extra') msg = '没有仅我方有的行。';
      noTd.textContent = msg;
      noTr.appendChild(noTd);
      el.detailBody.appendChild(noTr);
    } else {
      var frag = document.createDocumentFragment();
      pageRows.forEach(function (row) {
        var tr = document.createElement('tr');
        tr.className = 'row-' + row.status;
        var diffMap = row.status === 'paired' && row.diffs.length
          ? new Map(row.diffs.map(function (d) { return [d.name, d]; }))
          : null;

        var td1 = document.createElement('td');
        td1.className = 'sticky-1 row-meta';
        var r1 = document.createElement('span');
        r1.className = 'rno';
        r1.appendChild(document.createTextNode('客户 '));
        var b1 = document.createElement('b');
        b1.textContent = row.customerRowNo == null ? '—' : row.customerRowNo;
        r1.appendChild(b1);
        var r2 = document.createElement('span');
        r2.className = 'rno';
        r2.appendChild(document.createTextNode('我方 '));
        var b2 = document.createElement('b');
        b2.textContent = row.ourRowNo == null ? '—' : row.ourRowNo;
        r2.appendChild(b2);
        var copy = document.createElement('button');
        copy.type = 'button';
        copy.className = 'row-copy';
        copy.dataset.rowId = row.__id;
        copy.textContent = '复制';
        copy.title = '复制该行的差异信息';
        td1.appendChild(r1);
        td1.appendChild(r2);
        td1.appendChild(copy);

        var td2 = document.createElement('td');
        td2.className = 'sticky-2 row-meta';
        var pill = document.createElement('span');
        pill.className = 'status-pill ' + row.status;
        pill.textContent = row.status === 'paired' ? '配对' : (row.status === 'missing' ? '缺失' : '多余');
        td2.appendChild(pill);
        if (row.status === 'paired' && row.diffCount > 0) {
          var badge = document.createElement('span');
          badge.className = 'row-diff-badge';
          badge.textContent = row.diffCount + ' 处';
          td2.appendChild(badge);
        }

        var td3 = document.createElement('td');
        td3.className = 'key-cell sticky-3';
        td3.textContent = row.keyDisplay || '（空键）';
        td3.title = row.keyDisplay || '（空键）';

        tr.appendChild(td1);
        tr.appendChild(td2);
        tr.appendChild(td3);

        cols.forEach(function (c) {
          var td = document.createElement('td');
          td.className = 'pair-cell';
          var cVal = null, oVal = null;
          var cAbsent = false, oAbsent = false;
          var diff = null;
          if (row.status === 'paired') {
            cAbsent = c.customerIndex < 0;
            oAbsent = c.ourIndex < 0;
            if (!cAbsent) cVal = row.customerValues[c.customerIndex];
            if (!oAbsent) oVal = row.ourValues[c.ourIndex];
            if (!cAbsent && !oAbsent && diffMap) diff = diffMap.get(c.name) || null;
          } else if (row.status === 'missing') {
            cAbsent = c.customerIndex < 0;
            oAbsent = true;
            if (!cAbsent) cVal = row.customerValues[c.customerIndex];
          } else {
            cAbsent = true;
            oAbsent = c.ourIndex < 0;
            if (!oAbsent) oVal = row.ourValues[c.ourIndex];
          }
          var diffTitle = diff
            ? ('客户：' + RC.text(diff.customerValue) + '\n我方：' + RC.text(diff.ourValue))
            : '';
          var oneSideTitle = '';
          if (row.status === 'paired' && (cAbsent || oAbsent)) {
            oneSideTitle = cAbsent ? '该列仅我方文件有' : '该列仅客户文件有';
          }
          var cTitle = diffTitle || (cAbsent && row.status === 'paired' ? oneSideTitle : '');
          var oTitle = diffTitle || (oAbsent && row.status === 'paired' ? oneSideTitle : '');
          td.appendChild(makeHalf('c', cVal, { diff: !!diff, absent: cAbsent, title: cTitle }));
          td.appendChild(makeHalf('o', oVal, { diff: !!diff, absent: oAbsent, title: oTitle }));
          tr.appendChild(td);
        });
        frag.appendChild(tr);
      });
      el.detailBody.appendChild(frag);
    }

    el.detailCount.textContent = total
      ? '共 ' + fmt(total) + ' 行 · 第 ' + state.view.page + ' / ' + pages + ' 页'
      : '共 0 行';
    el.pageTotal.textContent = pages;
    el.pageInput.max = pages;
    el.pageInput.value = state.view.page;
    el.pageFirst.disabled = state.view.page <= 1;
    el.pagePrev.disabled = state.view.page <= 1;
    el.pageNext.disabled = state.view.page >= pages;
    el.pageLast.disabled = state.view.page >= pages;
    el.pageSize.value = String(state.view.pageSize);
  }
  function updateFilterButtons() {
    $all('button[data-filter]', el.rowFilters).forEach(function (btn) {
      var active = btn.dataset.filter === state.view.filter;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
  }
  function renderActiveFilters() {
    el.activeFilters.replaceChildren();
    if (state.view.column) {
      var col = state.columnMap.get(state.view.column);
      var chip = document.createElement('span');
      chip.className = 'filter-chip';
      chip.appendChild(document.createTextNode('列：' + (col ? col.display : state.view.column)));
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.dataset.clear = 'column';
      btn.textContent = '×';
      btn.title = '清除列筛选';
      chip.appendChild(btn);
      el.activeFilters.appendChild(chip);
    }
    if (state.view.search.trim()) {
      var chip2 = document.createElement('span');
      chip2.className = 'filter-chip';
      var q = state.view.search.trim();
      chip2.appendChild(document.createTextNode('搜索：' + (q.length > 24 ? q.slice(0, 24) + '…' : q)));
      var btn2 = document.createElement('button');
      btn2.type = 'button';
      btn2.dataset.clear = 'search';
      btn2.textContent = '×';
      btn2.title = '清除搜索';
      chip2.appendChild(btn2);
      el.activeFilters.appendChild(chip2);
    }
  }
  function toggleColumnFocus(name) {
    var col = state.columnMap.get(name);
    if (!col) return;
    if (col.oneSide) {
      toast('该列仅在' + (col.oneSide === 'customer' ? '客户' : '我方') + '文件中存在，无法逐行比较。', 'error');
      return;
    }
    if (col.diff === 0) {
      toast('「' + col.display + '」在已配对行中没有差异。', 'info');
      return;
    }
    if (state.view.column === name) {
      state.view.column = null;
      state.view.filter = 'all';
    } else {
      state.view.column = name;
      state.view.filter = 'diff';
    }
    state.view.page = 1;
    renderColumnStats(state.result);
    renderDetail();
  }
  function setFilter(filter) {
    if (!state.result) return;
    state.view.filter = filter;
    state.view.column = null;
    state.view.page = 1;
    renderColumnStats(state.result);
    renderDetail();
  }
  function setSort(value, toggle) {
    if (toggle && state.view.sort === value) {
      state.view.dir = state.view.dir === 'asc' ? 'desc' : 'asc';
    } else {
      state.view.sort = value;
      state.view.dir = 'asc';
    }
    state.view.page = 1;
    syncSortControls();
    renderDetail();
  }

  /* ---------------- export / copy ---------------- */
  async function copyText(text) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (e) { /* fall through */ }
    var ta = null;
    try {
      ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      ta.style.top = '0';
      document.body.appendChild(ta);
      ta.select();
      return document.execCommand('copy');
    } catch (e) {
      return false;
    } finally {
      if (ta && ta.parentNode) ta.remove();
    }
  }
  var copyEscHandler = null;
  function closeCopyModal() {
    el.copyModal.hidden = true;
    el.copyModalText.value = '';
    document.body.classList.remove('busy');
    if (copyEscHandler) {
      document.removeEventListener('keydown', copyEscHandler);
      copyEscHandler = null;
    }
  }
  function showCopyFallback(text) {
    el.copyModalText.value = text;
    el.copyModal.hidden = false;
    document.body.classList.add('busy');
    if (copyEscHandler) document.removeEventListener('keydown', copyEscHandler);
    copyEscHandler = function (e) {
      if (e.key === 'Escape') {
        e.preventDefault();
        closeCopyModal();
      }
    };
    document.addEventListener('keydown', copyEscHandler);
    requestAnimationFrame(function () {
      el.copyModalText.focus();
      el.copyModalText.select();
    });
  }
  async function copyRow(id) {
    var result = state.result;
    if (!result) return;
    var row = result.rows[id];
    if (!row) return;
    var text = report.buildRowCopyText(row, result.columns);
    var ok = await copyText(text);
    if (ok) {
      toast('已复制该行差异信息。', 'success');
    } else {
      showCopyFallback(text);
      toast('自动复制被浏览器拦截，已打开手动复制窗口。', 'info');
    }
  }
  async function exportXlsx() {
    var result = state.result;
    if (!result) { toast('还没有可导出的对比结果。', 'error'); return; }
    if (isBusy()) { toast('正在处理其他任务，请稍候。', 'error'); return; }
    showOverlay('正在生成 Excel…', '正在整理三个工作表，请稍候。', null);
    await nextPaint(30);
    try {
      var wb = report.buildReportWorkbook(result, {
        customerFile: state.customer.file ? state.customer.file.name : '',
        customerSheet: state.customer.sheetName,
        customerHeaderRow: state.customer.headerRow,
        ourFile: state.ours.file ? state.ours.file.name : '',
        ourSheet: state.ours.sheetName,
        ourHeaderRow: state.ours.headerRow
      });
      var filename = '对账结果_' + timestamp() + '.xlsx';
      XLSX.writeFile(wb, filename, { compression: true });
      hideOverlay();
      toast('已导出：' + filename, 'success');
    } catch (e) {
      hideOverlay();
      toast('导出失败：' + (e.message || String(e)), 'error');
    }
  }

  /* ---------------- events / init ---------------- */
  function bindDropzone(key) {
    var ui = sideEls[key];
    var drop = ui.drop;
    var input = ui.input;
    drop.addEventListener('click', function (e) {
      if (e.target === input) return;
      if (e.target.closest('button')) return;
      input.click();
    });
    drop.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        input.click();
      }
    });
    input.addEventListener('change', function () {
      var file = input.files && input.files[0];
      input.value = '';
      if (file) handleFile(key, file);
    });
    ui.reselect.addEventListener('click', function (e) {
      e.stopPropagation();
      input.click();
    });
    ['dragenter', 'dragover'].forEach(function (type) {
      drop.addEventListener(type, function (e) {
        e.preventDefault();
        e.stopPropagation();
        drop.classList.add('dragover');
      });
    });
    ['dragleave', 'dragend'].forEach(function (type) {
      drop.addEventListener(type, function (e) {
        e.preventDefault();
        e.stopPropagation();
        drop.classList.remove('dragover');
      });
    });
    drop.addEventListener('drop', function (e) {
      e.preventDefault();
      e.stopPropagation();
      drop.classList.remove('dragover');
      var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) handleFile(key, file);
    });
  }
  function bindEvents() {
    bindDropzone('customer');
    bindDropzone('ours');

    el.profile.addEventListener('change', function () {
      if (isBusy()) {
        el.profile.value = state.profile.id;
        return;
      }
      state.profile = profiles[el.profile.value];
      state.compareToken++;
      ['customer', 'ours'].forEach(function (key) {
        var side = state[key];
        if (side.rawRows && !side.headerTouched) detectSideHeader(key);
      });
      state.keyError = '';
      state.result = null;
      state.filtered = [];
      resetView();
      renderAll();
      scheduleRecompute(0);
    });

    ['customer', 'ours'].forEach(function (key) {
      var ui = sideEls[key];
      ui.sheet.addEventListener('change', function (e) { switchSheet(key, e.target.value); });
      ui.header.addEventListener('change', function (e) { applyHeaderRow(key, e.target.value); });
      ui.header.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); }
      });
      ui.autoHeader.addEventListener('click', function () { autoHeader(key); });
    });

    el.keyOptions.addEventListener('change', function (e) {
      var input = e.target.closest('input[data-key-name]');
      if (!input) return;
      var info = getUnionInfo();
      if (!info) return;
      var name = input.dataset.keyName;
      var names = state.keyNames.slice();
      if (input.checked) {
        if (names.indexOf(name) === -1) names.push(name);
      } else {
        names = names.filter(function (n) { return n !== name; });
      }
      var order = new Map(info.union.map(function (c, i) { return [c.name, i]; }));
      names.sort(function (a, b) { return (order.get(a) || 0) - (order.get(b) || 0); });
      state.keyNames = names;
      state.keyMode = 'manual';
      state.keyInfo = null;
      state.keyError = '';
      syncKeySelection(info);
      scheduleRecompute(90);
    });
    el.keyPresets.addEventListener('click', function (e) {
      var btn = e.target.closest('button');
      if (!btn) return;
      var info = getUnionInfo();
      if (!info) return;
      if (btn.dataset.auto) {
        state.keyMode = 'auto';
        state.keyInfo = null;
        state.keyError = '';
      } else if (btn.__names) {
        state.keyNames = btn.__names.slice();
        state.keyMode = 'manual';
        state.keyInfo = { names: btn.__names.slice(), label: btn.textContent, source: 'preset' };
        state.keyError = '';
      }
      syncKeySelection(info);
      scheduleRecompute(0);
    });
    el.keyChips.addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-key]');
      if (!btn) return;
      var info = getUnionInfo();
      if (!info) return;
      state.keyNames = state.keyNames.filter(function (n) { return n !== btn.dataset.key; });
      state.keyMode = 'manual';
      state.keyInfo = null;
      state.keyError = '';
      syncKeySelection(info);
      scheduleRecompute(90);
    });

    el.rowFilters.addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-filter]');
      if (!btn) return;
      setFilter(btn.dataset.filter);
    });
    $all('[data-filter-card]').forEach(function (btn) {
      btn.addEventListener('click', function () { setFilter(btn.dataset.filterCard); });
    });
    el.searchInput.addEventListener('input', function () {
      clearTimeout(state.searchTimer);
      state.searchTimer = setTimeout(function () {
        state.view.search = el.searchInput.value;
        state.view.page = 1;
        renderDetail();
      }, 220);
    });
    el.searchInput.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        el.searchInput.value = '';
        state.view.search = '';
        state.view.page = 1;
        renderDetail();
      } else if (e.key === 'Enter') {
        e.preventDefault();
        state.view.search = el.searchInput.value;
        state.view.page = 1;
        renderDetail();
      }
    });
    el.diffColsOnly.addEventListener('change', function () {
      state.view.diffColsOnly = el.diffColsOnly.checked;
      state.view.page = 1;
      renderDetail();
    });
    el.sortSelect.addEventListener('change', function () { setSort(el.sortSelect.value, false); });
    el.sortDir.addEventListener('click', function () {
      state.view.dir = state.view.dir === 'asc' ? 'desc' : 'asc';
      state.view.page = 1;
      syncSortControls();
      renderDetail();
    });
    el.detailHead.addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-col]');
      if (!btn) return;
      setSort('column:' + btn.dataset.col, true);
    });
    el.detailBody.addEventListener('click', function (e) {
      var btn = e.target.closest('button.row-copy');
      if (!btn) return;
      copyRow(Number(btn.dataset.rowId));
    });
    el.colstatsBody.addEventListener('click', function (e) {
      var tr = e.target.closest('tr[data-col]');
      if (!tr) return;
      toggleColumnFocus(tr.dataset.col);
    });
    el.colstatsBody.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      var tr = e.target.closest('tr[data-col]');
      if (!tr) return;
      e.preventDefault();
      toggleColumnFocus(tr.dataset.col);
    });
    el.activeFilters.addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-clear]');
      if (!btn) return;
      if (btn.dataset.clear === 'column') {
        state.view.column = null;
        state.view.filter = 'all';
        state.view.page = 1;
        renderColumnStats(state.result);
        renderDetail();
      } else if (btn.dataset.clear === 'search') {
        el.searchInput.value = '';
        state.view.search = '';
        state.view.page = 1;
        renderDetail();
      }
    });

    el.pageFirst.addEventListener('click', function () { goPage(1); });
    el.pagePrev.addEventListener('click', function () { goPage(state.view.page - 1); });
    el.pageNext.addEventListener('click', function () { goPage(state.view.page + 1); });
    el.pageLast.addEventListener('click', function () {
      goPage(Math.max(1, Math.ceil(state.filtered.length / state.view.pageSize)));
    });
    el.pageInput.addEventListener('change', function () { goPage(Number(el.pageInput.value)); });
    el.pageSize.addEventListener('change', function () {
      state.view.pageSize = Number(el.pageSize.value) || 100;
      state.view.page = 1;
      renderPage();
    });
    el.exportBtn.addEventListener('click', exportXlsx);
    el.copyModalClose.addEventListener('click', closeCopyModal);
    el.copyModal.addEventListener('click', function (e) {
      if (e.target === el.copyModal) closeCopyModal();
    });
    document.addEventListener('dragover', function (e) {
      if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types, 'Files') !== -1) e.preventDefault();
    });
    document.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
        if (!(e.target && e.target.closest && e.target.closest('.dropzone'))) e.preventDefault();
      }
    });
    document.addEventListener('keydown', function (e) {
      if (e.key !== '/' || e.ctrlKey || e.metaKey || e.altKey) return;
      var t = e.target;
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
      e.preventDefault();
      el.searchInput.focus();
    });
  }
  function goPage(n) {
    var pages = Math.max(1, Math.ceil(state.filtered.length / state.view.pageSize));
    n = Math.floor(Number(n));
    if (!isFinite(n)) return;
    state.view.page = Math.max(1, Math.min(pages, n));
    renderPage();
    el.detailScroll.scrollTop = 0;
  }

  function init() {
    cacheEls();
    Object.values(profiles).forEach(function (profile) {
      el.profile.add(new Option(profile.label, profile.id));
    });
    el.profile.value = state.profile.id;
    resetView();
    bindEvents();
    renderAll();
    renderEngineStatus();
    renderNotices();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})(window.ReconcileCore, window.ReconcileProfiles, window.ReconcileWorkbook, window.ReconcileReport, window.XLSX);
