const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { LAB_COLUMNS, DISEASE_COLUMNS, labRows, diseaseRows } = require('./initialData');

const css = fs.readFileSync(path.join(__dirname, 'app.css'), 'utf8');
const appJs = fs.readFileSync(path.join(__dirname, 'app.js'), 'utf8');
const xlsxLib = fs.readFileSync(path.join(__dirname, 'node_modules/xlsx/dist/xlsx.core.min.js'), 'utf8');
const faviconBase64 = fs.readFileSync(path.join(__dirname, 'favicon.png')).toString('base64');

// initialData.js の現在の内容を、アプリ起動時の初期表示用データとして埋め込む(Excel未読み込みでも最新データが見える)
function rowsToAOA(columns, rows) {
  const aoa = [columns];
  rows.forEach(r => aoa.push(columns.map(c => (r[c] !== undefined ? r[c] : ''))));
  return aoa;
}
const bundledLabRaw = rowsToAOA(LAB_COLUMNS, labRows);
const bundledDiseaseRaw = rowsToAOA(DISEASE_COLUMNS, diseaseRows);
const bundledJson = JSON.stringify({ labRaw: bundledLabRaw, diseaseRaw: bundledDiseaseRaw });
const bundledVersion = crypto.createHash('sha1').update(bundledJson).digest('hex').slice(0, 12);
// builtAt: このビルドを実行した時刻。ブラウザが何らかの理由で古いキャッシュ済みHTMLを読み込んでしまった場合に、
// 「内容が違う」だけでなく「今持っているデータより新しいか」を判定できるようにするためのタイムスタンプ。
// これが無いと、古いHTMLが誤って読み込まれた際に新しいデータが古いデータで上書きされてしまう(逆行)。
const builtAt = Date.now();
const bundledScript = 'window.__BUNDLED__ = Object.assign(' + bundledJson + ', { version: ' + JSON.stringify(bundledVersion) + ', builtAt: ' + builtAt + ' });';

const bodyTop = `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>検査値リファレンス</title>
<link rel="icon" type="image/png" href="data:image/png;base64,${faviconBase64}">
<style>
${css}
</style>
</head>
<body>
<div id="app">
  <header class="topbar">
    <h1 class="brand">検査値リファレンス<span>KENSA REFERENCE</span></h1>
    <div class="file-controls">
      <button class="btn primary" id="btnLoad">Excelを読み込む</button>
      <input type="file" id="fileInput" accept=".xlsx,.xlsm,.xls" style="display:none;">
      <button class="btn" id="btnExportExcel" style="display:none;">Excelとして書き出す</button>
      <button class="btn" id="btnUseBundled" style="display:none;">最新の内蔵データに更新</button>
      <span class="file-status" id="fileStatus">ファイル未読み込み</span>
    </div>
    <div style="flex:1;"></div>
    <button class="btn memo-badge" id="memoBtn">未収載メモ<span class="memo-count" id="memoCount" style="display:none;"></span></button>
  </header>
  <div class="warn-banner" id="warnBanner" style="display:none;"></div>
  <div class="warn-banner" id="errorBanner" style="display:none;"></div>
  <div class="searchbar">
    <input type="text" id="searchInput" placeholder="検査値名・略号・症状・疾患名などで検索（例: alb / アルブミン / むくみ / 糖尿病）" autofocus autocomplete="off">
    <div class="search-meta" id="searchMeta"></div>
  </div>
  <main class="layout">
    <div class="pane-list" id="paneList"></div>
    <div class="pane-detail" id="paneDetail"></div>
  </main>
  <footer class="app-footer">個人用ツール／データの正本はExcelファイルです／本アプリは検査値の自動判定を行いません。判断は薬剤師が行ってください。</footer>
</div>
<script>
// ===== SheetJS (xlsx) — オフライン動作のためライブラリ本体をこのファイルに内包 =====
${xlsxLib}
</script>
<script>
${bundledScript}
</script>
<script>
${appJs}
</script>
</body>
</html>
`;

const outPath = process.argv[2];
fs.writeFileSync(outPath, bodyTop, 'utf8');
console.log('written:', outPath, 'size(KB):', (Buffer.byteLength(bodyTop, 'utf8') / 1024).toFixed(1));
