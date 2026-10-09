// ★ v0.38.23 — 거래명세서 품명·자동 글자 맞춤 + 출하 내역 품명·단가 회귀 (사용자 요청 2026-10-09)
// Usage: node _verify-tx23.js   (VELA_FILE=index-dev.html 지원)
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.resolve(__dirname, process.env.VELA_FILE || 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window; win.VELA_BACKEND = 'local';
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.matchMedia = q => ({ matches:false, media:q, addListener(){}, removeListener(){}, addEventListener(){}, removeEventListener(){} });
win.eval(scripts.join('\n;\n') + `;window.__X = s => eval(s);`);
const X = s => win.__X(s);
const R = []; const P = (n, ok, note) => { R.push({ n, ok }); console.log(`[${ok?'PASS':'FAIL'}] ${n}${note?'  — '+note:''}`); };
X(`seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='검증'; SESSION.userId='U-T';`);
const a = X(`(()=>{ window._shTab='list'; location.hash='#ship'; router();
  const t = [...document.querySelectorAll('#view table')].find(t => /품명/.test(t.querySelector('thead')?.textContent||'') && /단가/.test(t.querySelector('thead')?.textContent||''));
  const heads = t ? [...t.querySelectorAll('thead th')].map(x => x.textContent.trim()) : [];
  const sh = DB.all('shipments').sort((x,y)=>y.ts-x.ts)[0]; const so = DB.get('orders', sh.so_id); const it = DB.get('items', sh.itemCode) || {};
  const row = t && [...t.querySelectorAll('tbody tr')][0]; const cells = row ? [...row.children].map(c => c.textContent.trim()) : [];
  const price = so.price || it.price || 0;
  return { heads, nameOk: cells[3] === (it.name||''), priceOk: cells[5] === '₩' + price.toLocaleString(), amtOk: cells[6] === '₩' + ((sh.qty||0)*price).toLocaleString() && (sh.qty||0)*price === shipAmount(sh) }; })()`);
P('1 출하 내역: 열 = 일시·SO·품번·품명·수량·단가·금액·LOT · 품명 = 품번 마스터 · 단가 = 수주 단가(없으면 품번 단가) · 금액 = shipAmount 와 같음', JSON.stringify(a.heads.slice(0,8)) === JSON.stringify(['일시','SO','품번','품명','수량','단가','금액','LOT']) && a.nameOk && a.priceOk && a.amtOk, JSON.stringify(a));
const b = X(`(()=>{ let cap = null; const o = _openPrint; _openPrint = (title, body, css, opts) => { cap = { body, css, opts }; };
  try { printForm.tx([DB.all('shipments')[0].id]); } finally { _openPrint = o; }
  const d = new DOMParser().parseFromString('<div>' + cap.body + '</div>', 'text/html');
  const heads = [...d.querySelectorAll('.fit-tbl thead th')].slice(0,8).map(x => x.textContent.trim());
  const row = d.querySelector('.fit-tbl tbody tr'); const sh = DB.all('shipments')[0];
  return { heads, cols: row.children.length, code: row.children[0].textContent.trim() === sh.itemCode, name: row.children[1].textContent.trim() === ((DB.get('items', sh.itemCode)||{}).name||''),
    nowrap: /\\.fit-tbl th, \\.fit-tbl td\\{white-space:nowrap/.test(cap.css) && /table\\.fit-tbl\\{table-layout:fixed\\}/.test(cap.css), fit: cap.opts && cap.opts.fit, formKey: cap.opts && cap.opts.formKey }; })()`);
P('2 거래명세서: 품번·품명 칸 분리(8열) · 각 칸 내용 일치 · 줄바꿈 금지(nowrap) + 고정 폭 표 · 글자 맞춤 대상 지정', JSON.stringify(b.heads) === JSON.stringify(['품번','품명','LOT','수량','포장내역','단가','공급가액','파렛트']) && b.cols === 8 && b.code && b.name && b.nowrap && /fit-tbl td/.test(b.fit||'') && /party/.test(b.fit||'') && b.formKey === 'tx', JSON.stringify(b));
const c = X(`(()=>{ const src = printFitCells.toString(); let html = ''; const ow = window.open; window.open = () => ({ document:{ write(h){ html += h; }, close(){} } });
  try { _openPrintWindow('t', '<div class="page"><table class="fit-tbl"><tr><td>x</td></tr></table></div>', '', 'portrait', '.fit-tbl td'); } finally { window.open = ow; }
  return { inject: html.includes('const __fit = function printFitCells') && html.includes('__fit(document, __fitSel)'), minClamp: /Math\\.max\\(6, fs - 0\\.5\\)/.test(src), important: /'important'/.test(src), noFitOther: (() => { let h2=''; const ow2 = window.open; window.open = () => ({ document:{ write(h){ h2 += h; }, close(){} } }); try { _openPrintWindow('t','<div></div>','',null); } finally { window.open = ow2; } return !h2.includes('__fit'); })() }; })()`);
P('3 인쇄창에 글자 맞춤 함수 주입(인쇄 직전 실행) · 최소 6px · !important(인쇄 글자 규칙보다 우선) · 다른 양식은 무주입', c.inject && c.minClamp && c.important && c.noFitOther, JSON.stringify(c));
P('4 버전 v0.38.23', X(`APP_VERSION`) === 'v0.38.23', X(`APP_VERSION`));
if(R.length !== 4) R.push({ n:`항목 수 ${R.length} ≠ 기대 4`, ok:false });
const fail = R.filter(x => !x.ok).length;
console.log(`\nTX23 TOTAL ${R.length} · PASS ${R.length - fail} · FAIL ${fail}`);
process.exit(fail ? 1 : 0);
