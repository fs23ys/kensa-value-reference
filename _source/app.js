(function () {
  'use strict';

  // ============ 定数 ============
  // 列名は指示書のExcel構造に合わせる。この配列以外に検査値名・疾患名はコード中に一切書かない。
  var LAB_COLUMNS = ['検査値名', '略号', '基準値', '一行要約', '患者説明例', '高値で考えること', '高値の時のアドバイス', '低値で考えること', '低値の時のアドバイス', '薬剤師補足', '疾患タグ', '検索キーワード', '参考ページ'];
  var DISEASE_COLUMNS = ['疾患名', '確認すべき検査値', '聞くこと', '声のかけ方', '見落としやすい点', '参考ページ'];
  var LAB_FULLNESS_COLS = ['基準値', '一行要約', '患者説明例', '高値で考えること', '低値で考えること'];
  var DISEASE_FULLNESS_COLS = ['聞くこと', '声のかけ方', '見落としやすい点'];
  var LAB_SHEET_NAME = '検査値';
  var DISEASE_SHEET_NAME = '疾患';
  var LS_WORKBOOK = 'labRefApp.workbook.v1';
  var LS_MEMOS = 'labRefApp.memos.v1';
  var LS_LAB_ORDER = 'labRefApp.order.lab.v1';
  var LS_DISEASE_ORDER = 'labRefApp.order.disease.v1';

  // ============ 状態 ============
  var state = {
    labRows: [],
    diseaseList: [],
    missingLabCols: [],
    missingDiseaseCols: [],
    fileName: '',
    loadedAt: '',
    hasData: false,
    query: '',
    listTab: 'lab', // 既定表示のタブ: 'lab' | 'disease'
    reorderMode: false,
    labOrder: [],      // ユーザーが並び替えた順のkey配列(検査値)
    diseaseOrder: [],  // 同上(疾患)
    selected: null, // {type:'lab', key} | {type:'disease', key}
    memos: [],
    labRawRows: [],      // 編集可能な生データ(検査値シート相当)。row.rawと同じ参照を保持
    diseaseRawRows: [],  // 編集可能な生データ(疾患シート相当)
    cardEditMode: false,
    editedLocally: false
  };

  var dragState = null; // {type:'lab'|'disease', key}

  // ============ ユーティリティ ============
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function normalize(str) {
    if (!str) return '';
    var s = String(str).toLowerCase();
    // 全角英数字 -> 半角
    s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    });
    s = s.replace(/　/g, ' ');
    // カタカナ -> ひらがな
    s = s.replace(/[ァ-ヶ]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0x60);
    });
    // 長音符・各種ハイフン類を除去(表記ゆれ吸収)
    s = s.replace(/[ー\-‐‑‒–—―ｰ]/g, '');
    s = s.replace(/\s+/g, '');
    return s;
  }

  function splitDelims(str) {
    if (!str) return [];
    return String(str).split(/[,，、\/\n]+/).map(function (t) { return t.trim(); }).filter(Boolean);
  }

  function fmtDateTime(iso) {
    try {
      var d = new Date(iso);
      var p = function (n) { return String(n).padStart(2, '0'); };
      return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
    } catch (e) { return iso; }
  }

  function dots(filled, total) {
    var s = '';
    for (var i = 0; i < total; i++) s += (i < filled ? '●' : '○');
    return s;
  }

  // ============ Excel読み込み・パース ============
  function sheetToRows(ws, expectedCols) {
    if (!ws) return { rows: [], missingCols: expectedCols.slice(), found: false };
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!aoa.length) return { rows: [], missingCols: expectedCols.slice(), found: true };
    var headerRow = aoa[0].map(function (h) { return String(h == null ? '' : h).trim(); });
    var colIndex = {};
    headerRow.forEach(function (h, i) { if (h && !(h in colIndex)) colIndex[h] = i; });
    var missingCols = expectedCols.filter(function (c) { return !(c in colIndex); });
    var rows = [];
    for (var r = 1; r < aoa.length; r++) {
      var arr = aoa[r];
      var obj = {};
      var hasAny = false;
      expectedCols.forEach(function (c) {
        var idx = colIndex[c];
        var v = '';
        if (idx !== undefined && arr[idx] !== undefined && arr[idx] !== null) {
          v = String(arr[idx]).trim();
        }
        obj[c] = v;
        if (v) hasAny = true;
      });
      if (hasAny) rows.push(obj);
    }
    return { rows: rows, missingCols: missingCols, found: true };
  }

  function findSheet(workbook, name) {
    var names = workbook.SheetNames || [];
    for (var i = 0; i < names.length; i++) {
      if (String(names[i]).trim() === name) return workbook.Sheets[names[i]];
    }
    return null;
  }

  function joinFields(row, cols) {
    return cols.map(function (c) { return row[c] || ''; }).join(' ');
  }

  function buildLabRow(raw, usedKeys) {
    var tags = splitDelims(raw['疾患タグ']);
    var filled = LAB_FULLNESS_COLS.reduce(function (n, c) { return n + (raw[c] ? 1 : 0); }, 0);
    var title = raw['検査値名'] || raw['略号'] || '(名称未記入)';
    // キーは行番号ではなく検査値名ベース(Excelの行順が変わっても並び替え設定が壊れないように)
    var base = 'lab-' + title;
    var key = base;
    var n = 2;
    while (usedKeys[key]) { key = base + '-' + n; n++; }
    usedKeys[key] = true;
    return {
      key: key,
      raw: raw,
      title: title,
      tags: tags,
      fullness: filled,
      searchBlob: normalize(joinFields(raw, LAB_COLUMNS))
    };
  }

  function buildDiseaseEntry(name, sheetRow, labRows) {
    var confirmLabs = labRows.filter(function (l) { return l.tags.indexOf(name) !== -1; });
    var filled = sheetRow ? DISEASE_FULLNESS_COLS.reduce(function (n, c) { return n + (sheetRow[c] ? 1 : 0); }, 0) : 0;
    var blob = normalize(name + (sheetRow ? ' ' + joinFields(sheetRow, DISEASE_COLUMNS) : ''));
    return {
      key: 'disease-' + name,
      name: name,
      sheetRow: sheetRow || null,
      confirmLabs: confirmLabs,
      fullness: filled,
      searchBlob: blob
    };
  }

  function rowsToAOA(columns, rows) {
    var aoa = [columns];
    rows.forEach(function (r) {
      aoa.push(columns.map(function (c) { return r[c] !== undefined ? r[c] : ''; }));
    });
    return aoa;
  }

  // labRawRows/diseaseRawRows(生データの配列)から表示用データを組み立てる。
  // 編集保存時にも同じ関数で再計算することで、パース直後と編集後で挙動を一致させる。
  function buildDerived(labRawRows, diseaseRawRows) {
    var usedKeys = {};
    var labRows = labRawRows.map(function (raw) { return buildLabRow(raw, usedKeys); });

    // 疾患一覧の自動生成: 検査値シートのタグ ∪ 疾患シートの疾患名
    var order = [];
    var map = {};
    labRows.forEach(function (l) {
      l.tags.forEach(function (t) {
        if (!map[t]) { map[t] = { name: t, sheetRow: null }; order.push(t); }
      });
    });
    diseaseRawRows.forEach(function (d) {
      var n = (d['疾患名'] || '').trim();
      if (!n) return;
      if (!map[n]) { map[n] = { name: n, sheetRow: d }; order.push(n); }
      else { map[n].sheetRow = d; }
    });
    var diseaseList = order.map(function (n) { return buildDiseaseEntry(n, map[n].sheetRow, labRows); });

    return { labRows: labRows, diseaseList: diseaseList };
  }

  function parseWorkbook(workbook) {
    var labWs = findSheet(workbook, LAB_SHEET_NAME);
    if (!labWs) {
      var err = new Error('「' + LAB_SHEET_NAME + '」シートが見つかりません。シート名を確認してください。');
      err.isUserError = true;
      throw err;
    }
    var labResult = sheetToRows(labWs, LAB_COLUMNS);
    var diseaseWs = findSheet(workbook, DISEASE_SHEET_NAME);
    var diseaseResult = diseaseWs ? sheetToRows(diseaseWs, DISEASE_COLUMNS) : { rows: [], missingCols: [], found: false };

    var derived = buildDerived(labResult.rows, diseaseResult.rows);

    return {
      labRawRows: labResult.rows,
      diseaseRawRows: diseaseResult.rows,
      labRows: derived.labRows,
      diseaseList: derived.diseaseList,
      missingLabCols: labResult.missingCols,
      missingDiseaseCols: diseaseWs ? diseaseResult.missingCols : [],
      diseaseSheetFound: !!diseaseWs
    };
  }

  function relatedLabs(row) {
    return state.labRows.filter(function (o) {
      return o !== row && o.tags.some(function (t) { return row.tags.indexOf(t) !== -1; });
    });
  }

  // ============ 永続化 ============
  function saveWorkbookToStorage(rawWorkbookJson, fileName) {
    try {
      localStorage.setItem(LS_WORKBOOK, JSON.stringify({
        fileName: fileName,
        loadedAt: new Date().toISOString(),
        labRaw: rawWorkbookJson.labRaw,
        diseaseRaw: rawWorkbookJson.diseaseRaw,
        diseaseSheetFound: rawWorkbookJson.diseaseSheetFound,
        editedLocally: !!rawWorkbookJson.editedLocally,
        bundledVersion: rawWorkbookJson.bundledVersion || null,
        bundledBuiltAt: rawWorkbookJson.bundledBuiltAt || null
      }));
    } catch (e) {
      console.warn('localStorage保存に失敗しました', e);
    }
  }

  // アプリ内編集の内容をブラウザだけに保存する(Excelファイルには一切書き込まない)
  function persistCurrentDataToStorage() {
    state.loadedAt = new Date().toISOString();
    state.editedLocally = true;
    saveWorkbookToStorage({
      labRaw: rowsToAOA(LAB_COLUMNS, state.labRawRows),
      diseaseRaw: rowsToAOA(DISEASE_COLUMNS, state.diseaseRawRows),
      diseaseSheetFound: true,
      editedLocally: true
    }, state.fileName);
  }

  function loadWorkbookFromStorage() {
    try {
      var s = localStorage.getItem(LS_WORKBOOK);
      if (!s) return null;
      return JSON.parse(s);
    } catch (e) { return null; }
  }

  function loadOrder(lsKey) {
    try {
      var s = localStorage.getItem(lsKey);
      return s ? JSON.parse(s) : [];
    } catch (e) { return []; }
  }
  function saveOrder(lsKey, arr) {
    try { localStorage.setItem(lsKey, JSON.stringify(arr)); } catch (e) {}
  }
  // 保存済みの並び順(キー配列)を適用する。並び順にない項目(Excelで新規追加された行など)は末尾に元の順で追加する。
  function applyOrder(items, orderArr) {
    if (!orderArr || !orderArr.length) return items.slice();
    var byKey = {};
    items.forEach(function (it) { byKey[it.key] = it; });
    var result = [];
    var used = {};
    orderArr.forEach(function (k) {
      if (byKey[k] && !used[k]) { result.push(byKey[k]); used[k] = true; }
    });
    items.forEach(function (it) {
      if (!used[it.key]) result.push(it);
    });
    return result;
  }

  function loadMemos() {
    try {
      var s = localStorage.getItem(LS_MEMOS);
      return s ? JSON.parse(s) : [];
    } catch (e) { return []; }
  }
  function saveMemos() {
    try { localStorage.setItem(LS_MEMOS, JSON.stringify(state.memos)); } catch (e) {}
  }

  // ============ DOM参照 ============
  var el = {};
  function cacheEls() {
    el.fileInput = document.getElementById('fileInput');
    el.btnLoad = document.getElementById('btnLoad');
    el.btnExportExcel = document.getElementById('btnExportExcel');
    el.btnUseBundled = document.getElementById('btnUseBundled');
    el.fileStatus = document.getElementById('fileStatus');
    el.warnBanner = document.getElementById('warnBanner');
    el.errorBanner = document.getElementById('errorBanner');
    el.searchInput = document.getElementById('searchInput');
    el.searchMeta = document.getElementById('searchMeta');
    el.paneList = document.getElementById('paneList');
    el.paneDetail = document.getElementById('paneDetail');
    el.memoBtn = document.getElementById('memoBtn');
    el.memoCount = document.getElementById('memoCount');
  }

  // ============ 描画: リスト ============
  function renderList() {
    var q = normalize(state.query);
    var html = '';

    if (!state.hasData) {
      el.paneList.innerHTML = '<div class="empty-hint">まだExcelが読み込まれていません。上部の「Excelを読み込む」ボタンからファイルを選択してください。</div>';
      el.searchMeta.textContent = '';
      return;
    }

    if (!q) {
      html += '<div class="list-tabs">' +
        '<button class="tab-btn' + (state.listTab === 'lab' ? ' active' : '') + '" data-tab="lab">検査値（' + state.labRows.length + '）</button>' +
        '<button class="tab-btn' + (state.listTab === 'disease' ? ' active' : '') + '" data-tab="disease">疾患（' + state.diseaseList.length + '）</button>' +
        '<button class="btn small reorder-toggle-btn" id="btnReorderToggle">' + (state.reorderMode ? '完了' : '⇅ 並び替え') + '</button>' +
        '</div>';
      if (state.reorderMode) {
        html += '<div class="reorder-hint">ドラッグ、または▲▼ボタンで順番を入れ替えられます。</div>';
      }
      if (state.listTab === 'disease') {
        var diseaseItems = applyOrder(state.diseaseList, state.diseaseOrder);
        html += '<ul class="item-list" id="itemListUl">' + diseaseItems.map(renderDiseaseListItem).join('') + '</ul>';
      } else {
        var labItems = applyOrder(state.labRows, state.labOrder);
        html += '<ul class="item-list" id="itemListUl">' + labItems.map(renderLabListItem).join('') + '</ul>';
      }
      el.paneList.innerHTML = html;
      el.searchMeta.textContent = '';
      bindListEvents();
      bindTabEvents();
      bindReorderEvents();
      return;
    }

    var diseaseMatches = state.diseaseList.filter(function (d) { return d.searchBlob.indexOf(q) !== -1; });
    var labMatches = state.labRows.filter(function (r) { return r.searchBlob.indexOf(q) !== -1; });
    var total = diseaseMatches.length + labMatches.length;
    el.searchMeta.textContent = total + ' 件ヒット';

    if (total === 0) {
      html += '<div class="no-result-box">' +
        '<div style="margin-bottom:10px;">「' + esc(state.query) + '」に一致するカードが見つかりませんでした。</div>' +
        '<button class="btn primary" id="btnRecordMemo">この検索語を未収載メモに記録する</button>' +
        '<div id="memoRecordedMsg" style="margin-top:8px;color:#276c3d;font-size:13px;"></div>' +
        '</div>';
      el.paneList.innerHTML = html;
      var btn = document.getElementById('btnRecordMemo');
      if (btn) btn.addEventListener('click', function () {
        recordMemo(state.query);
        document.getElementById('memoRecordedMsg').textContent = '「' + state.query + '」を未収載メモに記録しました。';
      });
      return;
    }

    if (diseaseMatches.length) {
      html += '<div class="section-heading">疾患（' + diseaseMatches.length + '）</div>';
      html += '<ul class="item-list">' + diseaseMatches.map(renderDiseaseListItem).join('') + '</ul>';
    }
    if (labMatches.length) {
      html += '<div class="section-heading">検査値（' + labMatches.length + '）</div>';
      html += '<ul class="item-list">' + labMatches.map(renderLabListItem).join('') + '</ul>';
    }
    el.paneList.innerHTML = html;
    bindListEvents();
  }

  function moveBtnsHtml(type, key) {
    return '<span class="move-btns">' +
      '<button class="move-btn" data-move="up" data-type="' + type + '" data-key="' + esc(key) + '" title="上へ">▲</button>' +
      '<button class="move-btn" data-move="down" data-type="' + type + '" data-key="' + esc(key) + '" title="下へ">▼</button>' +
      '</span>';
  }

  function renderLabListItem(row) {
    var selected = !state.reorderMode && state.selected && state.selected.type === 'lab' && state.selected.key === row.key;
    var cls = 'item-row lab-row' + (selected ? ' selected' : '');
    var handle = state.reorderMode ? '<span class="drag-handle">⋮⋮</span>' : '';
    var right = state.reorderMode ? moveBtnsHtml('lab', row.key) : ('<div class="dots" title="充実度">' + dots(row.fullness, LAB_FULLNESS_COLS.length) + '</div>');
    return '<li class="' + cls + '" data-type="lab" data-key="' + esc(row.key) + '"' + (state.reorderMode ? ' draggable="true"' : '') + '>' +
      handle +
      '<div class="item-main"><div class="name">' + esc(row.title) + '</div>' +
      (row.raw['略号'] ? '<div class="sub">' + esc(row.raw['略号']) + '</div>' : '') + '</div>' +
      right +
      '</li>';
  }

  function renderDiseaseListItem(d) {
    var selected = !state.reorderMode && state.selected && state.selected.type === 'disease' && state.selected.key === d.key;
    var cls = 'item-row disease-row' + (selected ? ' selected' : '');
    var handle = state.reorderMode ? '<span class="drag-handle">⋮⋮</span>' : '';
    var right = state.reorderMode ? moveBtnsHtml('disease', d.key) : ('<div class="dots" title="充実度">' + dots(d.fullness, DISEASE_FULLNESS_COLS.length) + '</div>');
    return '<li class="' + cls + '" data-type="disease" data-key="' + esc(d.key) + '"' + (state.reorderMode ? ' draggable="true"' : '') + '>' +
      handle +
      '<div class="item-main"><div class="name">' + esc(d.name) + '</div>' +
      '<div class="sub">確認すべき検査値 ' + d.confirmLabs.length + ' 件</div></div>' +
      right +
      '</li>';
  }

  function bindTabEvents() {
    el.paneList.querySelectorAll('.tab-btn').forEach(function (b) {
      b.addEventListener('click', function () {
        state.listTab = b.getAttribute('data-tab');
        renderList();
      });
    });
    var toggleBtn = document.getElementById('btnReorderToggle');
    if (toggleBtn) toggleBtn.addEventListener('click', function () {
      state.reorderMode = !state.reorderMode;
      renderList();
    });
  }

  function bindListEvents() {
    var rows = el.paneList.querySelectorAll('.item-row');
    rows.forEach(function (r) {
      r.addEventListener('click', function () {
        if (state.reorderMode) return; // 並び替え中は選択・遷移しない
        var type = r.getAttribute('data-type');
        var key = r.getAttribute('data-key');
        if (type === 'lab') selectLabByKey(key);
        else selectDiseaseByKey(key);
      });
    });
  }

  // ============ 並び替え(ドラッグ&ドロップ / ▲▼ボタン) ============
  function currentOrderKeys() {
    var items = state.listTab === 'disease' ? applyOrder(state.diseaseList, state.diseaseOrder) : applyOrder(state.labRows, state.labOrder);
    return items.map(function (it) { return it.key; });
  }
  function persistOrder(keys) {
    if (state.listTab === 'disease') {
      state.diseaseOrder = keys;
      saveOrder(LS_DISEASE_ORDER, keys);
    } else {
      state.labOrder = keys;
      saveOrder(LS_LAB_ORDER, keys);
    }
  }
  function moveKey(keys, key, dir) {
    var idx = keys.indexOf(key);
    if (idx === -1) return keys;
    var newIdx = idx + dir;
    if (newIdx < 0 || newIdx >= keys.length) return keys;
    var arr = keys.slice();
    var tmp = arr[idx];
    arr[idx] = arr[newIdx];
    arr[newIdx] = tmp;
    return arr;
  }
  function reorderByDrop(keys, draggedKey, targetKey, before) {
    var arr = keys.filter(function (k) { return k !== draggedKey; });
    var idx = arr.indexOf(targetKey);
    if (idx === -1) return keys;
    arr.splice(before ? idx : idx + 1, 0, draggedKey);
    return arr;
  }

  function bindReorderEvents() {
    if (!state.reorderMode) return;
    el.paneList.querySelectorAll('.move-btn').forEach(function (b) {
      b.addEventListener('click', function (e) {
        e.stopPropagation();
        var key = b.getAttribute('data-key');
        var dir = b.getAttribute('data-move') === 'up' ? -1 : 1;
        var keys = moveKey(currentOrderKeys(), key, dir);
        persistOrder(keys);
        renderList();
      });
    });

    var ul = document.getElementById('itemListUl');
    if (!ul) return;
    ul.querySelectorAll('.item-row').forEach(function (li) {
      li.addEventListener('dragstart', function (e) {
        dragState = { key: li.getAttribute('data-key') };
        li.classList.add('dragging');
        if (e.dataTransfer) { e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', li.getAttribute('data-key')); } catch (err) {} }
      });
      li.addEventListener('dragend', function () {
        li.classList.remove('dragging');
        ul.querySelectorAll('.item-row').forEach(function (x) { x.classList.remove('drag-over-top', 'drag-over-bottom'); });
        dragState = null;
      });
      li.addEventListener('dragover', function (e) {
        if (!dragState) return;
        e.preventDefault();
        var rect = li.getBoundingClientRect();
        var before = (e.clientY - rect.top) < rect.height / 2;
        ul.querySelectorAll('.item-row').forEach(function (x) { x.classList.remove('drag-over-top', 'drag-over-bottom'); });
        li.classList.add(before ? 'drag-over-top' : 'drag-over-bottom');
      });
      li.addEventListener('drop', function (e) {
        e.preventDefault();
        if (!dragState) return;
        var rect = li.getBoundingClientRect();
        var before = (e.clientY - rect.top) < rect.height / 2;
        var targetKey = li.getAttribute('data-key');
        if (targetKey !== dragState.key) {
          var keys = reorderByDrop(currentOrderKeys(), dragState.key, targetKey, before);
          persistOrder(keys);
        }
        dragState = null;
        renderList();
      });
    });
  }

  // ============ フィールド表示ヘルパー ============
  function fieldBlock(label, value, extraClass, badge) {
    var blank = !value;
    var labelHtml = esc(label) + (badge ? ' <span class="label-badge">' + esc(badge) + '</span>' : '');
    return '<div class="block ' + (extraClass || '') + '">' +
      '<div class="block-label">' + labelHtml + '</div>' +
      '<div class="block-body' + (blank ? ' blank' : '') + '">' + (blank ? '未記入' : esc(value)) + '</div>' +
      '</div>';
  }

  // ============ フィールド表示ヘルパー(高値/低値 + アドバイス併記) ============
  function hiloBlock(label, diagVal, adviceVal, cls) {
    var blank = !diagVal;
    var html = '<div class="block ' + cls + '">';
    html += '<div class="block-label">' + esc(label) + '</div>';
    html += '<div class="block-body' + (blank ? ' blank' : '') + '">' + (blank ? '未記入' : esc(diagVal)) + '</div>';
    if (adviceVal) {
      html += '<div class="advice-sub"><span class="advice-tag">アドバイス</span>' + esc(adviceVal) + '</div>';
    }
    html += '</div>';
    return html;
  }

  // ============ アプリ内編集(このブラウザだけに保存。Excelは書き換えない) ============
  var LAB_FIELD_TYPES = {
    '検査値名': 'input', '略号': 'input', '基準値': 'textarea-s',
    '一行要約': 'textarea-s', '患者説明例': 'textarea-l',
    '高値で考えること': 'textarea-s', '高値の時のアドバイス': 'textarea-s',
    '低値で考えること': 'textarea-s', '低値の時のアドバイス': 'textarea-s',
    '薬剤師補足': 'textarea-m', '疾患タグ': 'input', '検索キーワード': 'input', '参考ページ': 'input'
  };
  var DISEASE_FIELD_TYPES = {
    '疾患名': 'input', '聞くこと': 'textarea-m', '声のかけ方': 'textarea-m',
    '見落としやすい点': 'textarea-m', '参考ページ': 'input'
  };

  function editFieldHtml(col, value, type) {
    var val = esc(value || '');
    if (type === 'input') {
      return '<div class="edit-field"><label>' + esc(col) + '</label><input type="text" data-field="' + esc(col) + '" value="' + val + '"></div>';
    }
    var rows = type === 'textarea-l' ? 6 : (type === 'textarea-m' ? 4 : 2);
    return '<div class="edit-field"><label>' + esc(col) + '</label><textarea data-field="' + esc(col) + '" rows="' + rows + '">' + val + '</textarea></div>';
  }

  function readEditFields(container, target) {
    container.querySelectorAll('[data-field]').forEach(function (f) {
      target[f.getAttribute('data-field')] = f.value.trim();
    });
  }

  function renderLabEditForm(row) {
    var raw = row.raw;
    var html = '<button class="btn back-btn" id="backBtn">← 一覧に戻る</button>';
    html += '<div class="card edit-card">';
    html += '<div class="edit-card-title">検査値を編集</div>';
    html += '<div class="edit-note">この変更はこの端末のブラウザだけに保存されます。Excelファイルは書き換わりません。</div>';
    LAB_COLUMNS.forEach(function (col) {
      html += editFieldHtml(col, raw[col], LAB_FIELD_TYPES[col]);
    });
    html += '<div class="edit-actions"><button class="btn primary" id="btnSaveEdit">保存</button> <button class="btn" id="btnCancelEdit">キャンセル</button></div>';
    html += '</div>';
    el.paneDetail.innerHTML = html;
    document.getElementById('backBtn').addEventListener('click', hideDetailPaneMobile);
    document.getElementById('btnCancelEdit').addEventListener('click', function () {
      state.cardEditMode = false;
      renderLabDetail(row);
    });
    document.getElementById('btnSaveEdit').addEventListener('click', function () { saveLabEdit(row); });
    showDetailPane();
  }

  function saveLabEdit(row) {
    var raw = row.raw;
    var oldKey = row.key;
    readEditFields(el.paneDetail, raw);
    var derived = buildDerived(state.labRawRows, state.diseaseRawRows);
    state.labRows = derived.labRows;
    state.diseaseList = derived.diseaseList;
    state.cardEditMode = false;
    var newRow = state.labRows.filter(function (r) { return r.raw === raw; })[0];
    // 検査値名の変更でキーが変わった場合、並び替え設定(labOrder)内の古いキーを新しいキーに置き換え、
    // 並び順が崩れて末尾に飛ばされないようにする
    if (newRow && newRow.key !== oldKey) {
      state.labOrder = state.labOrder.map(function (k) { return k === oldKey ? newRow.key : k; });
      saveOrder(LS_LAB_ORDER, state.labOrder);
    }
    persistCurrentDataToStorage();
    updateFileStatus();
    renderList();
    if (newRow) selectLabByKey(newRow.key);
    else el.paneDetail.innerHTML = '<div class="empty-hint">左の一覧から選び直してください。</div>';
  }

  function renderDiseaseEditForm(entry) {
    var raw = entry.sheetRow || {};
    var html = '<button class="btn back-btn" id="backBtn">← 一覧に戻る</button>';
    html += '<div class="card edit-card">';
    html += '<div class="edit-card-title">疾患を編集</div>';
    html += '<div class="edit-note">この変更はこの端末のブラウザだけに保存されます。Excelファイルは書き換わりません。</div>';
    DISEASE_COLUMNS.forEach(function (col) {
      if (col === '確認すべき検査値') {
        html += '<div class="edit-field"><label>確認すべき検査値</label><div class="edit-readonly">自動生成のため編集できません（現在' + entry.confirmLabs.length + '件）</div></div>';
        return;
      }
      var val = col === '疾患名' ? (raw['疾患名'] || entry.name) : (raw[col] || '');
      html += editFieldHtml(col, val, DISEASE_FIELD_TYPES[col]);
    });
    html += '<div class="edit-actions"><button class="btn primary" id="btnSaveEdit">保存</button> <button class="btn" id="btnCancelEdit">キャンセル</button></div>';
    html += '</div>';
    el.paneDetail.innerHTML = html;
    document.getElementById('backBtn').addEventListener('click', hideDetailPaneMobile);
    document.getElementById('btnCancelEdit').addEventListener('click', function () {
      state.cardEditMode = false;
      renderDiseaseDetail(entry);
    });
    document.getElementById('btnSaveEdit').addEventListener('click', function () { saveDiseaseEdit(entry); });
    showDetailPane();
  }

  function saveDiseaseEdit(entry) {
    var raw = entry.sheetRow;
    var isNew = !raw;
    var oldKey = entry.key;
    if (isNew) {
      raw = {};
      DISEASE_COLUMNS.forEach(function (c) { raw[c] = ''; });
    }
    readEditFields(el.paneDetail, raw);
    if (!raw['疾患名']) raw['疾患名'] = entry.name;
    if (isNew) state.diseaseRawRows.push(raw);
    var derived = buildDerived(state.labRawRows, state.diseaseRawRows);
    state.labRows = derived.labRows;
    state.diseaseList = derived.diseaseList;
    state.cardEditMode = false;
    var newName = raw['疾患名'].trim();
    var newEntry = state.diseaseList.filter(function (d) { return d.name === newName; })[0];
    // 疾患名の変更でキーが変わった場合、並び替え設定(diseaseOrder)内の古いキーを新しいキーに置き換える
    if (newEntry && newEntry.key !== oldKey) {
      state.diseaseOrder = state.diseaseOrder.map(function (k) { return k === oldKey ? newEntry.key : k; });
      saveOrder(LS_DISEASE_ORDER, state.diseaseOrder);
    }
    persistCurrentDataToStorage();
    updateFileStatus();
    renderList();
    if (newEntry) selectDiseaseByKey(newEntry.key);
    else el.paneDetail.innerHTML = '<div class="empty-hint">左の一覧から選び直してください。</div>';
  }

  // ============ 描画: 検査値カード ============
  function renderLabDetail(row) {
    if (state.cardEditMode) { renderLabEditForm(row); return; }
    var raw = row.raw;
    var refBlank = !raw['基準値'];
    var summaryBlank = !raw['一行要約'];

    var html = '<button class="btn back-btn" id="backBtn">← 一覧に戻る</button>';
    html += '<div class="card">';
    html += '<div class="card-top"><span class="title">' + esc(row.title) + '</span>';
    if (raw['略号']) html += '<span class="abbr">' + esc(raw['略号']) + '</span>';
    html += '<button class="btn small edit-toggle-btn" id="btnEditToggle">✎ 編集</button>';
    html += '</div>';
    html += '<div class="card-refval' + (refBlank ? ' blank' : '') + '">' + (refBlank ? '基準値：未記入' : esc(raw['基準値'])) + '</div>';
    html += '<div class="ref-note">※施設により基準値が異なります。検査表の基準値欄を優先してください</div>';
    html += '<div class="summary-line' + (summaryBlank ? ' blank' : '') + '">' + (summaryBlank ? '一行要約：未記入' : esc(raw['一行要約'])) + '</div>';

    html += fieldBlock('患者説明例（そのまま声に出せます）', raw['患者説明例'], 'patient');

    html += '<div class="hilo-grid">';
    html += hiloBlock('▲ 高いとき', raw['高値で考えること'], raw['高値の時のアドバイス'], 'high');
    html += hiloBlock('▼ 低いとき', raw['低値で考えること'], raw['低値の時のアドバイス'], 'low');
    html += '</div>';

    html += fieldBlock('薬剤師補足', raw['薬剤師補足'], 'pharma', '患者さんには言わない');

    html += '<div class="block tag-row"><div class="block-label">疾患タグ</div>';
    if (row.tags.length) {
      html += row.tags.map(function (t) {
        return '<span class="chip disease-chip" data-nav-disease="' + esc(t) + '">' + esc(t) + '</span>';
      }).join('');
    } else {
      html += '<div class="block-body blank">未記入</div>';
    }
    html += '</div>';

    var rel = relatedLabs(row);
    html += '<div class="block related-list"><div class="block-label">関連する検査値（疾患タグが共通するもの・自動抽出）</div>';
    if (rel.length) {
      html += rel.map(function (r2) {
        return '<span class="chip" data-nav-lab="' + esc(r2.key) + '">' + esc(r2.title) + '</span>';
      }).join('');
    } else {
      html += '<div class="block-body blank">なし</div>';
    }
    html += '</div>';

    html += fieldBlock('参考ページ（書籍）', raw['参考ページ'], 'refpage');

    html += '</div>';
    el.paneDetail.innerHTML = html;
    bindDetailEvents();
    var editBtn = document.getElementById('btnEditToggle');
    if (editBtn) editBtn.addEventListener('click', function () {
      state.cardEditMode = true;
      renderLabDetail(row);
    });
    showDetailPane();
  }

  // ============ 描画: 疾患カード ============
  function renderToneBlock(text) {
    if (!text) return '<div class="block-body blank">未記入</div>';
    // 【見出し】を強調しつつ改行して表示
    var parts = String(text).split(/(?=【[^】]*】)/g).map(function (p) { return p.trim(); }).filter(Boolean);
    if (parts.length <= 1) return '<div class="block-body tone-block">' + esc(text) + '</div>';
    var html = '<div class="block-body tone-block">';
    parts.forEach(function (p) {
      var m = p.match(/^【([^】]*)】([\s\S]*)$/);
      if (m) {
        html += '<div><span class="tone-tag">' + esc(m[1]) + '</span>' + esc(m[2].trim()) + '</div>';
      } else {
        html += '<div>' + esc(p) + '</div>';
      }
    });
    html += '</div>';
    return html;
  }

  function renderQuestionBlock(text) {
    if (!text) return '<div class="block-body blank">未記入</div>';
    var qs = String(text).split(/[／\/\n]+/).map(function (t) { return t.trim(); }).filter(Boolean);
    if (qs.length <= 1) return '<div class="block-body">' + esc(text) + '</div>';
    return '<div class="block-body">' + qs.map(function (q) { return '<div class="qa-line">・' + esc(q) + '</div>'; }).join('') + '</div>';
  }

  function renderDiseaseDetail(entry) {
    if (state.cardEditMode) { renderDiseaseEditForm(entry); return; }
    var raw = entry.sheetRow || {};
    var html = '<button class="btn back-btn" id="backBtn">← 一覧に戻る</button>';
    html += '<div class="card">';
    html += '<div class="disease-header">' + esc(entry.name) + '<button class="btn small edit-toggle-btn" id="btnEditToggle">✎ 編集</button></div>';

    html += '<div class="block"><div class="block-label">確認すべき検査値（自動生成）</div>';
    if (entry.confirmLabs.length) {
      html += entry.confirmLabs.map(function (l) {
        return '<span class="chip" data-nav-lab="' + esc(l.key) + '">' + esc(l.title) + '</span>';
      }).join('');
    } else {
      html += '<div class="block-body blank">該当する検査値タグがまだ登録されていません</div>';
    }
    html += '</div>';

    html += '<div class="block"><div class="block-label">聞くこと</div>' + renderQuestionBlock(raw['聞くこと']) + '</div>';
    html += '<div class="block"><div class="block-label">声のかけ方</div>' + renderToneBlock(raw['声のかけ方']) + '</div>';
    html += fieldBlock('見落としやすい点', raw['見落としやすい点'], 'pharma', '自分用');
    html += fieldBlock('参考ページ（書籍）', raw['参考ページ'], 'refpage');

    html += '</div>';
    el.paneDetail.innerHTML = html;
    bindDetailEvents();
    var editBtn = document.getElementById('btnEditToggle');
    if (editBtn) editBtn.addEventListener('click', function () {
      state.cardEditMode = true;
      renderDiseaseDetail(entry);
    });
    showDetailPane();
  }

  function bindDetailEvents() {
    var backBtn = document.getElementById('backBtn');
    if (backBtn) backBtn.addEventListener('click', hideDetailPaneMobile);
    el.paneDetail.querySelectorAll('[data-nav-disease]').forEach(function (c) {
      c.addEventListener('click', function () { selectDiseaseByKey('disease-' + c.getAttribute('data-nav-disease')); });
    });
    el.paneDetail.querySelectorAll('[data-nav-lab]').forEach(function (c) {
      c.addEventListener('click', function () { selectLabByKey(c.getAttribute('data-nav-lab')); });
    });
  }

  function showDetailPane() {
    el.paneDetail.classList.add('mobile-active');
    el.paneList.classList.add('mobile-hidden');
    window.scrollTo(0, 0);
    el.paneDetail.scrollTop = 0;
  }
  function hideDetailPaneMobile() {
    el.paneDetail.classList.remove('mobile-active');
    el.paneList.classList.remove('mobile-hidden');
  }

  function selectLabByKey(key) {
    var row = state.labRows.filter(function (r) { return r.key === key; })[0];
    if (!row) return;
    state.cardEditMode = false;
    state.selected = { type: 'lab', key: key };
    renderLabDetail(row);
    highlightSelected();
  }
  function selectDiseaseByKey(key) {
    var d = state.diseaseList.filter(function (x) { return x.key === key; })[0];
    if (!d) return;
    state.cardEditMode = false;
    state.selected = { type: 'disease', key: key };
    renderDiseaseDetail(d);
    highlightSelected();
  }
  function highlightSelected() {
    el.paneList.querySelectorAll('.item-row').forEach(function (r) {
      var isSel = state.selected && r.getAttribute('data-key') === state.selected.key;
      r.classList.toggle('selected', !!isSel);
    });
  }

  // ============ 未収載メモ ============
  function recordMemo(term) {
    term = (term || '').trim();
    if (!term) return;
    var existing = state.memos.filter(function (m) { return m.term === term; })[0];
    if (existing) existing.recordedAt = new Date().toISOString();
    else state.memos.unshift({ term: term, recordedAt: new Date().toISOString() });
    saveMemos();
    updateMemoBadge();
  }
  function deleteMemo(term) {
    state.memos = state.memos.filter(function (m) { return m.term !== term; });
    saveMemos();
    updateMemoBadge();
    renderMemoPanel();
  }
  function updateMemoBadge() {
    if (state.memos.length) {
      el.memoCount.style.display = 'inline-block';
      el.memoCount.textContent = state.memos.length;
    } else {
      el.memoCount.style.display = 'none';
      el.memoCount.textContent = '';
    }
  }
  function csvEscape(v) {
    v = String(v == null ? '' : v);
    if (/[",\n]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
    return v;
  }
  function exportMemoCsv() {
    var lines = ['検索語,記録日時'];
    state.memos.forEach(function (m) {
      lines.push(csvEscape(m.term) + ',' + csvEscape(fmtDateTime(m.recordedAt)));
    });
    var csv = '﻿' + lines.join('\r\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = '未収載メモ.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function copyMemoText() {
    var text = state.memos.map(function (m) { return m.term; }).join('\n');
    function fallbackCopy() {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      document.body.removeChild(ta);
      return ok;
    }
    var msgEl = document.getElementById('memoCopyMsg');
    function showResult(ok) {
      if (msgEl) msgEl.textContent = ok ? 'コピーしました（' + state.memos.length + '件）。チャットに貼り付けてください。' : 'コピーに失敗しました。お手数ですが手入力をお願いします。';
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { showResult(true); }, function () { showResult(fallbackCopy()); });
    } else {
      showResult(fallbackCopy());
    }
  }

  function renderMemoPanel() {
    state.selected = null;
    var html = '<button class="btn back-btn" id="backBtn">← 一覧に戻る</button>';
    html += '<div class="memo-panel">';
    html += '<h2 style="margin-top:0;">未収載メモ（' + state.memos.length + '件）</h2>';
    html += '<p style="color:var(--text-muted);font-size:13px;">検索して見つからなかった項目の記録です。「コピー」で検査値名の一覧をコピーしてチャットに貼り付けるか、CSVに書き出して使ってください。</p>';
    html += '<button class="btn primary" id="btnCopyMemo" ' + (state.memos.length ? '' : 'disabled') + '>コピー</button> ';
    html += '<button class="btn" id="btnExportCsv" ' + (state.memos.length ? '' : 'disabled') + '>CSV書き出し</button>';
    html += '<div id="memoCopyMsg" style="margin-top:8px;color:#276c3d;font-size:13px;"></div>';
    html += '<div style="margin-top:14px;">';
    if (!state.memos.length) {
      html += '<div class="empty-hint">まだ記録はありません。</div>';
    } else {
      state.memos.forEach(function (m) {
        html += '<div class="memo-item"><div><div class="term">' + esc(m.term) + '</div><div class="time">' + esc(fmtDateTime(m.recordedAt)) + '</div></div>' +
          '<button class="btn small danger" data-del-memo="' + esc(m.term) + '">削除</button></div>';
      });
    }
    html += '</div></div>';
    el.paneDetail.innerHTML = html;
    document.getElementById('backBtn').addEventListener('click', hideDetailPaneMobile);
    var exportBtn = document.getElementById('btnExportCsv');
    if (exportBtn) exportBtn.addEventListener('click', exportMemoCsv);
    var copyBtn = document.getElementById('btnCopyMemo');
    if (copyBtn) copyBtn.addEventListener('click', copyMemoText);
    el.paneDetail.querySelectorAll('[data-del-memo]').forEach(function (b) {
      b.addEventListener('click', function () { deleteMemo(b.getAttribute('data-del-memo')); });
    });
    showDetailPane();
  }

  // ============ ファイル読み込み ============
  function handleFile(file) {
    hideBanner(el.errorBanner);
    hideBanner(el.warnBanner);
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var parsed = parseWorkbook(workbook);

        state.labRawRows = parsed.labRawRows;
        state.diseaseRawRows = parsed.diseaseRawRows;
        state.labRows = parsed.labRows;
        state.diseaseList = parsed.diseaseList;
        state.missingLabCols = parsed.missingLabCols;
        state.missingDiseaseCols = parsed.missingDiseaseCols;
        state.fileName = file.name;
        state.loadedAt = new Date().toISOString();
        state.hasData = true;
        state.editedLocally = false;
        state.selected = null;
        state.query = '';
        el.searchInput.value = '';

        // ローカル保存(生データをそのまま保持: 次回はExcel再選択なしで復元)
        saveWorkbookToStorage({
          labRaw: rowsToAOA(LAB_COLUMNS, state.labRawRows),
          diseaseRaw: rowsToAOA(DISEASE_COLUMNS, state.diseaseRawRows),
          diseaseSheetFound: parsed.diseaseSheetFound
        }, file.name);

        afterDataReady();
      } catch (err) {
        showBanner(el.errorBanner, err.isUserError ? err.message : ('Excelの読み込みに失敗しました: ' + err.message));
      }
    };
    reader.onerror = function () {
      showBanner(el.errorBanner, 'ファイルの読み込みに失敗しました。');
    };
    reader.readAsArrayBuffer(file);
  }

  function aoaToSheet(aoa) {
    return XLSX.utils.aoa_to_sheet(aoa);
  }

  function restoreFromStorage() {
    var saved = loadWorkbookFromStorage();
    if (!saved || !saved.labRaw) return false;
    try {
      var wb = { SheetNames: [], Sheets: {} };
      wb.SheetNames.push(LAB_SHEET_NAME);
      wb.Sheets[LAB_SHEET_NAME] = aoaToSheet(saved.labRaw);
      if (saved.diseaseRaw) {
        wb.SheetNames.push(DISEASE_SHEET_NAME);
        wb.Sheets[DISEASE_SHEET_NAME] = aoaToSheet(saved.diseaseRaw);
      }
      var parsed = parseWorkbook(wb);
      state.labRawRows = parsed.labRawRows;
      state.diseaseRawRows = parsed.diseaseRawRows;
      state.labRows = parsed.labRows;
      state.diseaseList = parsed.diseaseList;
      state.missingLabCols = parsed.missingLabCols;
      state.missingDiseaseCols = parsed.missingDiseaseCols;
      state.fileName = saved.fileName;
      state.loadedAt = saved.loadedAt;
      state.hasData = true;
      state.editedLocally = !!saved.editedLocally;
      return true;
    } catch (e) {
      console.warn('保存データの復元に失敗しました', e);
      return false;
    }
  }

  // アプリに内蔵された最新データ(assemble.js実行時点のinitialData.js内容)を読み込む
  function loadFromBundled(bundled) {
    try {
      var wb = { SheetNames: [], Sheets: {} };
      wb.SheetNames.push(LAB_SHEET_NAME);
      wb.Sheets[LAB_SHEET_NAME] = aoaToSheet(bundled.labRaw);
      wb.SheetNames.push(DISEASE_SHEET_NAME);
      wb.Sheets[DISEASE_SHEET_NAME] = aoaToSheet(bundled.diseaseRaw);
      var parsed = parseWorkbook(wb);
      state.labRawRows = parsed.labRawRows;
      state.diseaseRawRows = parsed.diseaseRawRows;
      state.labRows = parsed.labRows;
      state.diseaseList = parsed.diseaseList;
      state.missingLabCols = parsed.missingLabCols;
      state.missingDiseaseCols = parsed.missingDiseaseCols;
      state.fileName = '検査値リファレンス初期データ（アプリ内蔵・最新）';
      state.loadedAt = new Date().toISOString();
      state.hasData = true;
      state.editedLocally = false;
      saveWorkbookToStorage({
        labRaw: bundled.labRaw,
        diseaseRaw: bundled.diseaseRaw,
        diseaseSheetFound: true,
        editedLocally: false,
        bundledVersion: bundled.version,
        bundledBuiltAt: bundled.builtAt
      }, state.fileName);
      return true;
    } catch (e) {
      console.warn('内蔵データの読み込みに失敗しました', e);
      return false;
    }
  }

  function afterDataReady() {
    updateFileStatus();
    updateWarnBanner();
    renderList();
    el.paneDetail.innerHTML = '<div class="empty-hint">左の一覧から検査値または疾患を選んでください。</div>';
    hideDetailPaneMobile();
    el.searchInput.focus();
  }

  function updateFileStatus() {
    if (!state.hasData) {
      el.fileStatus.innerHTML = 'ファイル未読み込み';
      el.btnLoad.textContent = 'Excelを読み込む';
      el.btnExportExcel.style.display = 'none';
      el.btnUseBundled.style.display = 'none';
      return;
    }
    el.fileStatus.innerHTML = '<strong>' + esc(state.fileName) + '</strong>　検査値' + state.labRows.length + '件／疾患' + state.diseaseList.length + '件　（読込: ' + fmtDateTime(state.loadedAt) + '）' +
      (state.editedLocally ? '　<span style="color:#8a6a12;">※このブラウザ内で編集済み(Excel未反映)</span>' : '');
    el.btnLoad.textContent = 'Excelを読み込み直す';
    el.btnExportExcel.style.display = '';
    el.btnUseBundled.style.display = window.__BUNDLED__ ? '' : 'none';
  }

  // アプリ内編集も含めた現在のデータを新しい.xlsxとしてダウンロードする(他端末への手動反映用)
  function exportWorkbook() {
    // 並び替え(ドラッグ&ドロップ/▲▼)は表示順のみに反映され生データの順序は変わらないため、
    // 書き出し時は現在の表示順(labOrder/diseaseOrder)を適用してから出力する。
    var orderedLabRaw = applyOrder(state.labRows, state.labOrder).map(function (l) { return l.raw; });
    var orderedDiseaseRaw = applyOrder(state.diseaseList, state.diseaseOrder)
      .map(function (d) { return d.sheetRow; })
      .filter(Boolean);

    var wb = { SheetNames: [], Sheets: {} };
    wb.SheetNames.push(LAB_SHEET_NAME);
    wb.Sheets[LAB_SHEET_NAME] = aoaToSheet(rowsToAOA(LAB_COLUMNS, orderedLabRaw));
    wb.SheetNames.push(DISEASE_SHEET_NAME);
    wb.Sheets[DISEASE_SHEET_NAME] = aoaToSheet(rowsToAOA(DISEASE_COLUMNS, orderedDiseaseRaw));
    var d = new Date();
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    var fname = '検査値リファレンス_' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '_' + pad(d.getHours()) + pad(d.getMinutes()) + '.xlsx';
    XLSX.writeFile(wb, fname);
  }

  function showBanner(elm, msg) { elm.textContent = msg; elm.style.display = 'block'; }
  function hideBanner(elm) { elm.style.display = 'none'; elm.textContent = ''; }

  function updateWarnBanner() {
    var msgs = [];
    if (state.missingLabCols.length) {
      msgs.push('「' + LAB_SHEET_NAME + '」シートで見つからなかった列: ' + state.missingLabCols.join('、'));
    }
    if (state.missingDiseaseCols.length) {
      msgs.push('「' + DISEASE_SHEET_NAME + '」シートで見つからなかった列: ' + state.missingDiseaseCols.join('、'));
    }
    if (msgs.length) {
      showBanner(el.warnBanner, '⚠ 想定と異なる列構成です。' + msgs.join(' ／ ') + '（該当項目は未記入として扱われます）');
    } else {
      hideBanner(el.warnBanner);
    }
  }

  // ============ 初期化 ============
  function init() {
    cacheEls();
    state.memos = loadMemos();
    updateMemoBadge();
    state.labOrder = loadOrder(LS_LAB_ORDER);
    state.diseaseOrder = loadOrder(LS_DISEASE_ORDER);

    el.btnLoad.addEventListener('click', function () { el.fileInput.click(); });
    el.fileInput.addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
      e.target.value = '';
    });
    el.btnExportExcel.addEventListener('click', exportWorkbook);
    el.btnUseBundled.addEventListener('click', function () {
      if (!window.__BUNDLED__) return;
      var msg = state.editedLocally
        ? 'この端末で行った✎編集の内容が失われます。最新の内蔵データに更新しますか？'
        : '表示中のデータを、このHTMLに内蔵されている最新データで上書きします。よろしいですか？';
      if (!window.confirm(msg)) return;
      state.labOrder = [];
      state.diseaseOrder = [];
      saveOrder(LS_LAB_ORDER, []);
      saveOrder(LS_DISEASE_ORDER, []);
      loadFromBundled(window.__BUNDLED__);
      afterDataReady();
    });
    el.searchInput.addEventListener('input', function (e) {
      state.query = e.target.value;
      renderList();
    });
    el.memoBtn.addEventListener('click', renderMemoPanel);

    var bundled = window.__BUNDLED__;
    var saved = loadWorkbookFromStorage();
    // 「バージョンが違うから更新」ではなく「今のデータより新しい(builtAtが大きい)から更新」で判定する。
    // ブラウザのキャッシュ等で偶然古いHTMLが読み込まれても、保存済みの新しいデータが古いデータで
    // 上書きされて逆行してしまうことがないようにするため。
    var bundledIsNewer = !saved || !saved.bundledBuiltAt || (bundled && bundled.builtAt > saved.bundledBuiltAt);
    var useBundled = !!bundled && (!saved || !saved.labRaw || (!saved.editedLocally && saved.bundledVersion !== bundled.version && bundledIsNewer));

    if (useBundled) {
      // データを最新の内蔵データに更新する際は、古い並び替え設定(前のデータ構成に基づくキー順)も
      // 一緒にリセットする。そうしないと内蔵データ側で整えた並び順が、古い並び替え設定に上書きされてしまう。
      state.labOrder = [];
      state.diseaseOrder = [];
      saveOrder(LS_LAB_ORDER, []);
      saveOrder(LS_DISEASE_ORDER, []);
    }

    var restored = useBundled ? loadFromBundled(bundled) : restoreFromStorage();
    if (restored) {
      afterDataReady();
    } else {
      updateFileStatus();
      el.paneList.innerHTML = '<div class="empty-hint">まだExcelが読み込まれていません。上部の「Excelを読み込む」ボタンからファイルを選択してください。</div>';
      el.paneDetail.innerHTML = '<div class="empty-hint">左上の「Excelを読み込む」ボタンから、検査値・疾患シートを含むExcelファイルを選んでください。</div>';
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
