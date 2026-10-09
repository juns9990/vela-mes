// 표준 시나리오 시뮬 회귀 (G1 증명) — CLAUDE.md §3-3 기준으로 재작성 (2026-07-26)
//  기준 수치(완전 불변이어야 함):
//   · 절단 보존식: 950(양품) + 30(불량) + 20(잔류) = 절단 투입 1,000
//   · 매출: 900개 × 3,000원 = ₩2,700,000
//  실행 경로는 화면과 동일한 코어 1벌(LOT.registerRecord · LOT.balance · shipAmount)만 사용.
// Usage: node _verify-sim-standard.js
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const target = path.join(__dirname, 'vela-mes-prototype.html');
const html = fs.readFileSync(target, 'utf8');

const results = [];
const record = (name, ok, detail) => {
  results.push({ name, ok, detail: detail || '' });
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${name}${detail ? `  — ${detail}` : ''}`);
};

// ── 부팅 (외부 스크립트 제거 · 스텁) ──
const stripped = html
  .replace(/<script\b[^>]*\bsrc=[^>]*><\/script>/g, '')
  .replace(/<link\s+rel="manifest"[^>]*>/g, '');
const virtualConsole = new VirtualConsole();
const dom = new JSDOM(stripped, { runScripts: 'outside-only', pretendToBeVisual: true, virtualConsole, url: 'http://localhost/' });
const win = dom.window;

win.XLSX = require(path.join(__dirname, 'vela-lib-xlsx.min.js'));
win.Html5Qrcode = function () { this.start = () => Promise.resolve(); this.stop = () => Promise.resolve(); };
win.Html5QrcodeScanner = function () { this.render = () => {}; this.clear = () => {}; };
win.qrcode = function () {
  return { addData: () => {}, make: () => {}, createSvgTag: () => '<svg/>', createDataURL: () => 'data:image/gif;base64,QRSTUB' };
};

const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const expose = `
;(function(){
  var names = ['DB','LOT','SESSION','shipAmount','getRoute','bomHasProc','fmtDate'];
  names.forEach(function(n){ try { window[n] = eval(n); } catch(e){} });
})();`;
try { win.eval(scripts.join('\n;\n') + expose); record('boot: eval inline scripts', true, `${scripts.length} block(s)`); }
catch (e) { record('boot: eval inline scripts', false, e.message); finish(); }

// ── 시나리오용 마스터 (표준 시나리오 전용 코드 — 기존 데이터와 충돌 없음) ──
const IT = 'SIM-STD-001';
win.DB.set('processes', 'CUT',  { code:'CUT',  name:'절단', seq:10, unit:'ea', type:'in', equipment:['절단기 1호'], inspect:[] });
win.DB.set('processes', 'SHIP', { code:'SHIP', name:'출하', seq:99, unit:'ea', type:'in', equipment:[], inspect:[] });
win.DB.set('items', IT, { code:IT, name:'표준시나리오품', mat:'SCM440', dia:50, weight:0.5, price:3000, cust:'C-SIM', box:50, safety:0, pack_qty:100, active:true });
win.DB.set('bom', IT, { code:IT, name:'표준시나리오품', rev:'A',
  material:{ type:'SCM440', dia:50, weight:0.5 },
  route:[ { seq:10, proc:'CUT', mode:'in' }, { seq:99, proc:'SHIP', mode:'in' } ] });

// ── 1) 생산LOT 1,000 생성 → 절단 950/30 등록 ──
const PLOT = 'PL-SIM-STD-1';
win.LOT.create({ plot_no: PLOT, itemCode: IT, materialLots: [], qty_initial: 1000 });
record('시나리오: 생산LOT 1,000 생성', win.DB.get('production_lots', PLOT).qty_initial === 1000);

win.LOT.registerRecord({ plot_no: PLOT, proc: 'CUT', equipment: '절단기 1호', good: 950, defect: 30, defect_reason: '치수불량' });

const good = win.LOT.totalGood(PLOT, 'CUT');
const input = win.LOT.totalInput(PLOT, 'CUT');
const remain = win.LOT.balance(PLOT, 'CUT');
record('보존식: 절단 양품 950', good === 950, `good=${good}`);
record('보존식: 절단 투입 980 (950+30)', input === 980, `input=${input}`);
record('보존식: 절단 잔류 20', remain === 20, `remain=${remain}`);
record('보존식: 950(양품)+30(불량)+20(잔류) = 절단 1,000', good + 30 + remain === 1000, `${good}+30+${remain}=${good + 30 + remain}`);

// ── 2) G1: 초과등록 차단 (잔량 20 에 21 등록 → 거부) ──
let blocked = false, blockMsg = '';
try { win.LOT.registerRecord({ plot_no: PLOT, proc: 'CUT', equipment: '절단기 1호', good: 21 }); }
catch (e) { blocked = true; blockMsg = e.message; }
record('G1: 초과등록 차단 (잔량 20 < 등록 21 → 거부)', blocked, blockMsg);
record('G1: 차단 후 수치 무변 (원장 보호)', win.LOT.balance(PLOT, 'CUT') === 20 && win.LOT.totalInput(PLOT, 'CUT') === 980);

// ── 3) 출하 900 × 3,000원 = ₩2,700,000 ──
record('시나리오: 출하 900 ≤ 직전공정 양품 950 (SHIP 잔량 검증)', win.LOT.balance(PLOT, 'SHIP') === 950, `ship-balance=${win.LOT.balance(PLOT, 'SHIP')}`);
win.DB.set('orders', 'SO-SIM-STD-1', { id:'SO-SIM-STD-1', itemCode: IT, qty: 900, price: 3000, cust:'C-SIM', ts: Date.now(), status:'open' });
win.DB.set('shipments', 'SH-SIM-STD-1', { id:'SH-SIM-STD-1', so_id:'SO-SIM-STD-1', itemCode: IT, plot_no: PLOT, qty: 900, ts: Date.now(), by:'SIM' });

const amt = win.shipAmount(win.DB.get('shipments', 'SH-SIM-STD-1'));
record('매출: 900개 × 3,000원 = ₩2,700,000 (shipAmount 1벌)', amt === 2700000, `₩${amt.toLocaleString()}`);

// ── 4) 단가 경로 회귀: SO 단가 우선 · 없으면 품번 표준단가 ──
win.DB.set('shipments', 'SH-SIM-STD-2', { id:'SH-SIM-STD-2', so_id:'SO-NONE', itemCode: IT, plot_no: PLOT, qty: 10, ts: Date.now(), by:'SIM' });
record('매출: SO 없으면 품번 표준단가 폴백 (10×3,000=30,000)', win.shipAmount(win.DB.get('shipments','SH-SIM-STD-2')) === 30000);

finish();

function finish() {
  const failed = results.filter(r => !r.ok);
  console.log('\n────────────────────────────────────────');
  console.log(`TOTAL ${results.length} · PASS ${results.length - failed.length} · FAIL ${failed.length}`);
  if (failed.length) { console.log('\nFAILED:'); failed.forEach(f => console.log(` - ${f.name}${f.detail ? `  (${f.detail})` : ''}`)); }
  process.exit(failed.length ? 1 : 0);
}
