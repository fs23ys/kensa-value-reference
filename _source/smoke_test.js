const fs = require('fs');
const { JSDOM, VirtualConsole } = require('jsdom');

const ABS_HTML = process.env.APP_HTML;
const vc = new VirtualConsole();
const consoleErrors = [];
vc.on('jsdomError', (e) => consoleErrors.push('jsdomError: ' + e.message));

const dom = new JSDOM(fs.readFileSync(ABS_HTML, 'utf8'), {
  runScripts: 'dangerously',
  resources: 'usable',
  url: 'file:///' + ABS_HTML.replace(/\\/g, '/'),
  pretendToBeVisual: true,
  virtualConsole: vc,
});
const { window } = dom;
function wait(ms) { return new Promise((r) => setTimeout(r, ms)); }

function setValue(el, val) {
  var proto = window.HTMLInputElement.prototype;
  var setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
  setter.call(el, val);
  el.dispatchEvent(new window.Event('input', { bubbles: true }));
}

(async () => {
  await wait(800);

  const searchInput = window.document.getElementById('searchInput');
  setValue(searchInput, '白血球');
  await wait(200);

  const names = Array.from(window.document.querySelectorAll('li.item-row .name')).map(n => n.textContent.trim());
  console.log('"白血球"検索結果(タブ区別なし):', JSON.stringify(names));

  const meta = window.document.getElementById('searchMeta');
  console.log('searchMeta:', meta ? meta.textContent : '(none)');

  console.log('consoleErrors:', JSON.stringify(consoleErrors));
  process.exit(0);
})().catch((e) => {
  console.error('TEST SCRIPT ERROR:', e);
  process.exit(1);
});
