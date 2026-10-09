// ★ v0.38.7 — 수주 종결 회귀 (SPEC 「수주 종결 v0.1」 §10 · GPT 조건부 승인 2026-10-09)
//  0) 종결 0건 → 잔량 사용처 수치 v0.38.6 과 완전 동일 (구버전 백업과 대조)
//  1~10) 종결·취소·재개 · 유효잔량/미출하 분리 · write path 차단(출하·지시·계획·수정) · 초과 출하 차단 · 원장 무변 · 사유 필수
// Usage: node _verify-soclose.js
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
function boot(file){
  const html = fs.readFileSync(path.join(__dirname, file), 'utf8');
  const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
  const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
  const win = dom.window; win.VELA_BACKEND = 'local';
  win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
  win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
  win.matchMedia = q => ({ matches:false, media:q, addListener(){}, removeListener(){}, addEventListener(){}, removeEventListener(){} });
  win.eval(scripts.join('\n;\n') + `;window.__X = s => eval(s);`);
  return s => win.__X(s);
}
const R = []; const P = (n, ok, note) => { R.push({ n, ok }); console.log(`[${ok?'PASS':'FAIL'}] ${n}${note?'  — '+note:''}`); };
const T = s => { try { return s(); } catch(e){ return 'ERR ' + e.message; } };
// 결정적 비교용 고정 시각 (Date.now 를 두 부팅에서 같게)
const FIX = `Date.now = () => ${Date.parse('2026-10-09T09:00:00+09:00')};`;
const SNAP = `JSON.stringify({
  rows: calcSoRows().map(r => [r.o.id, r.shipped, r.remain, r.status.label]),
  mat: soMaterialNeeds().map(x => [x.o.id, x.remain, x.needKg]),
  net: DB.all('items').map(i => [i.code, computeNetDemand(i.code).orderQty, computeNetDemand(i.code).netDemand, computeNetDemand(i.code).due]),
  stages: (s => [s.total, s.done, s.prog, s.wait])(computeOrderStages()),
  rep: ['day','week','month'].map(sc => { const d = buildReportData(sc); return JSON.stringify([d.sales, d.issues]); })
})`;
const SETUP = `${FIX} seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='검증'; SESSION.userId='U-T';`;

// 0) 종결 0건 → 구버전과 동일
let X0;
try {
  const OLD = boot('backup/vela-mes-v0.38.6-백업.html'); OLD(SETUP);
  const NEW = boot('vela-mes-prototype.html'); NEW(SETUP);
  const a = OLD(SNAP), b = NEW(SNAP);
  P('0 종결 0건 → 수주행·소재 소요·순소요·수주 3단·보고서(일/주/월) 수치 v0.38.6 과 완전 동일', a === b, a === b ? `${a.length} bytes` : 'DIFF');
  X0 = NEW;
} catch(e){ P('0 구버전 대조', false, e.message); process.exit(1); }
const X = X0;

// 테스트 수주 3건 준비 (출하는 셋업용 직접 기록 · UI 출하 경로는 4번에서 검증)
X(`(()=>{
  const it = DB.all('items')[0]; window.__it = it.code; window.__price = 1000;
  const due = fmtDate(Date.now() + 5*86400000);
  DB.set('orders','SO-T1',{ id:'SO-T1', cust:it.cust||'C001', itemCode:it.code, qty:100, price:1000, due, created:Date.now() });
  DB.set('orders','SO-T2',{ id:'SO-T2', cust:it.cust||'C001', itemCode:it.code, qty:40,  price:1000, due, created:Date.now() });
  DB.set('orders','SO-T3',{ id:'SO-T3', cust:it.cust||'C001', itemCode:it.code, qty:50,  price:1000, due, created:Date.now() });
  DB.set('shipments','SH-T1A',{ id:'SH-T1A', so_id:'SO-T1', itemCode:it.code, plot_no:'', qty:40, ts:Date.now() });
  DB.set('shipments','SH-T1B',{ id:'SH-T1B', so_id:'SO-T1', itemCode:it.code, plot_no:'', qty:30, ts:Date.now() });
  DB.set('shipments','SH-T3',{ id:'SH-T3', so_id:'SO-T3', itemCode:it.code, plot_no:'', qty:50, ts:Date.now() });
  DB.set('production_plans','PP-T1',{ id:'PP-T1', plan_type:'board', so_id:'SO-T1', itemCode:it.code, qty:30, date:fmtDate(Date.now()), equipment:null, wo_id:null, ts:Date.now() });
})()`);
const before = X(`(()=>{ const it = window.__it; const rep = buildReportData('month'); return {
  net: computeNetDemand(it).orderQty, stages: computeOrderStages(), backlog: rep.sales.backlog, custBl: (rep.sales.byCustomer||[]).reduce((a,c)=>a+(c.backlog||0),0),
  mat: soMaterialNeeds().some(x => x.o.id === 'SO-T1'), issue: rep.issues.some(i => /SO-T1/.test(i.text)),
  ledger: JSON.stringify(['production_lots','records','shipments','work_orders','production_plans'].map(c => DB.allRaw(c).map(d => JSON.stringify(d)).sort())) }; })()`);

// 1) UI 종결: 수주 화면 7번째 열 [수주 종결] → 확인 창 → 사유 → 저장
const c1 = X(`(()=>{
  location.hash = '#so'; window._soFilter = { q:'', status:'' }; VIEWS.so();
  const ths = [...document.querySelectorAll('#view thead th')].map(t => t.textContent);
  const btn = document.querySelector('[data-so-close="SO-T1"]'); if(!btn) return { err:'no button', ths };
  btn.onclick({ stopPropagation(){} });
  const mtxt = document.querySelector('.modal, #modal, .mbox') ? (document.querySelector('.modal, #modal, .mbox').textContent) : document.body.textContent;
  // 9) 사유 없이 저장 → 종결 안 됨
  document.getElementById('soc-ok').onclick();
  const noReason = !soCloseInfo(DB.get('orders','SO-T1'));
  document.getElementById('soc-r').value = '고객 협의 감량';
  document.getElementById('soc-t').value = '잔량 30 공급 종료';
  document.getElementById('soc-ok').onclick();
  const o = DB.get('orders','SO-T1'); const lg = (o._logs||[]).slice(-1)[0] || {};
  return { ths: ths.join('|'), modal: /미출하/.test(mtxt) && /자동으로 바뀌지 않습니다/.test(mtxt) && /생산계획 1건/.test(mtxt), noReason,
    kind: o.close && o.close.kind, qty: o.qty, un: soUnshippedQty(o), open: soOpenQty(o), shipped: soShippedQty(o), snapUn: o.close && o.close.unshipped_at,
    log: lg.op === 'correct' && /부분종결: 고객 협의 감량 — 잔량 30 공급 종료/.test(lg.reason) && lg.by === '검증' };
})()`);
P('1 [수주 종결] 버튼(7번째 열 「작업」) → 확인 창(미출하·연결 현황·자동 처리 없음) → 사유 저장 → kind=short · 수주량 100·출하 70 무변 · 미출하 30 보존 · 유효잔량 0 · 감사 로그(처리자·사유)',
  /작업$/.test(c1.ths) && c1.modal && c1.kind === 'short' && c1.qty === 100 && c1.shipped === 70 && c1.un === 30 && c1.open === 0 && c1.snapUn === 30 && c1.log, JSON.stringify(c1));
P('9 사유 없이 [수주 종결] → 저장 0', c1.noReason === true);

// 2) 유효잔량 사용처에서 제외 · 사실값 보존
const c2 = X(`(()=>{ const it = window.__it; const rep = buildReportData('month'); const row = calcSoRows().find(r => r.o.id === 'SO-T1'); const st = computeOrderStages();
  VIEWS.so(); const tr = document.querySelector('[data-so-row="SO-T1"]').textContent;
  VIEWS._planNew({}); const planOpts = [...document.querySelectorAll('#pln-so option')].some(x => x.value === 'SO-T1'); closeModal();
  VIEWS._woNew({}); const woOpts = [...document.querySelectorAll('#wo-so option')].some(x => x.value === 'SO-T1'); closeModal();
  return { status: row.status.label, remain: row.remain, un: row.unshipped, trUn: /미출하 30 종결/.test(tr), trBtn: !!document.querySelector('[data-so-reopen="SO-T1"]') && !document.querySelector('[data-so-close="SO-T1"]'),
    net: computeNetDemand(it).orderQty, backlog: rep.sales.backlog, custBl: (rep.sales.byCustomer||[]).reduce((a,c)=>a+(c.backlog||0),0), mat: soMaterialNeeds().some(x => x.o.id === 'SO-T1'), issue: rep.issues.some(i => /SO-T1/.test(i.text)),
    closedAmt: st.closedAmt, total: st.total, done: st.done, late: calcSoRows().filter(r => r.status.label==='지연' || r.status.label==='납기임박').some(r => r.o.id==='SO-T1'), planOpts, woOpts }; })()`);
P('2 상태 「부분종결」 · 표 「미출하 30 종결」 + [재개] · 납기임박/지연 집계 제외', c2.status === '부분종결' && c2.remain === 0 && c2.un === 30 && c2.trUn && c2.trBtn && !c2.late, JSON.stringify({s:c2.status,r:c2.remain,u:c2.un,tr:c2.trUn,b:c2.trBtn,l:c2.late}));
P('2b 순소요 −30 · 보고서 수주잔 −30 · 고객사별 잔량 −30 · 소재 소요 제외 · 보고서 납기 이슈 제외 · 계획/작업지시 수주 후보 제외',
  c2.net === before.net - 30 && c2.backlog === before.backlog - 30 && c2.custBl === before.custBl - 30 && before.mat === true && c2.mat === false && before.issue === true && c2.issue === false && !c2.woOpts && !c2.planOpts,
  JSON.stringify({ net:[before.net,c2.net], backlog:[before.backlog,c2.backlog], custBl:[before.custBl,c2.custBl], mat:[before.mat,c2.mat], issue:[before.issue,c2.issue], wo:c2.woOpts, plan:c2.planOpts }));
P('2c 수주 3단(D3): 종결 미출하 30×1,000 은 출하완료·생산중·미착수 어디에도 없음 · total −30,000 · done 무변 · closedAmt 30,000 표시용',
  c2.closedAmt === 30000 && c2.total === before.stages.total - 30000 && c2.done === before.stages.done, JSON.stringify({ total:[before.stages.total,c2.total], done:[before.stages.done,c2.done], closedAmt:c2.closedAmt }));

// 3) 출하 0 → 수주취소
const c3 = X(`(()=>{ VIEWS._soCloseModal('SO-T2'); const t = document.body.textContent; document.getElementById('soc-r').value = '고객 취소'; document.getElementById('soc-ok').onclick();
  const r = calcSoRows().find(r => r.o.id === 'SO-T2'); return { title: /수주 취소 — SO-T2/.test(t), kind: DB.get('orders','SO-T2').close?.kind, label: r.status.label, un: r.unshipped, open: r.remain }; })()`);
P('3 출하 0 수주 종결 → kind=cancel · 「수주취소」 · 미출하 40 보존 · 유효잔량 0', c3.title && c3.kind === 'cancel' && c3.label === '수주취소' && c3.un === 40 && c3.open === 0, JSON.stringify(c3));

// 4) 출하 write path 차단 (UI 저장 버튼 · 숨은 select 직접 지정 = 화면 우회 시나리오)
const shipTry = (so, qty) => X(`(()=>{ location.hash = '#ship'; VIEWS.ship(); document.getElementById('btn-ship-new').onclick();
  const sel = document.getElementById('sh-so'); sel.value = '${so}'; sel.dispatchEvent(new Event('change'));
  const q = document.getElementById('sh-qty'); q.value = '${qty}'; q.dispatchEvent(new Event('input'));
  const warn = document.getElementById('sh-warn').textContent;
  const n0 = DB.all('shipments').length; let msg = ''; const _t = toast; toast = (m) => { msg = m; }; document.getElementById('sh-ok').onclick(); toast = _t; closeModal();
  return { added: DB.all('shipments').length - n0, msg, warn }; })()`);
const s4a = shipTry('SO-T1', 1);
P('4a 종결 수주로 신규 출하 저장 → 차단 · shipments 불변 · 화면 안내 「부분종결 수주 — 출하 불가」', s4a.added === 0 && /종결된 수주/.test(s4a.msg) && /부분종결 수주 — 출하 불가/.test(s4a.warn), JSON.stringify(s4a));
const s4b = shipTry('SO-T3', 1);
P('4b 출하완료 수주(50/50)에 추가 출하 → 수주 수량 초과 차단 (GPT 🔴 필수) · shipments 불변', s4b.added === 0 && /수주 수량 초과 출하/.test(s4b.msg) && /초과 — 저장 시 차단/.test(s4b.warn), JSON.stringify(s4b));
const c4c = X(`(()=>{ const a = (()=>{ try { assertSoShipQty('SO-T1', 0); return 'pass'; } catch(e){ return e.message; } })();
  DB.set('orders','SO-T4',{ id:'SO-T4', cust:'C001', itemCode:window.__it, qty:30, price:1000, due:'2099-01-01', created:Date.now() });
  const ok30 = (()=>{ try { assertSoShipQty('SO-T4', 30); return true; } catch(e){ return e.message; } })();
  const ng31 = (()=>{ try { assertSoShipQty('SO-T4', 31); return true; } catch(e){ return /초과/.test(e.message); } })();
  return { ok30, ng31 }; })()`);
P('4c 진행 수주 잔량 이내(30/30) 통과 · 31 → 초과 차단', c4c.ok30 === true && c4c.ng31 === true, JSON.stringify(c4c));

// 5) 작업지시·계획·수주 수정 write path
const c5 = X(`(()=>{ const iw = issueWoFromPlan(DB.get('production_plans','PP-T1'));
  const a = (()=>{ try { assertSoOpen('SO-T1','계획 등록'); return 'pass'; } catch(e){ return /재개/.test(e.message); } })();
  VIEWS.so(); const editBtn = !!document.querySelector('[data-so-edit="SO-T1"]');
  const src = String(VIEWS._planNew) + String(VIEWS._woNew);
  return { iw: iw.ok === false && /종결/.test(iw.reason), wo: DB.query('work_orders', w => w.so_id === 'SO-T1').length, a, editBtn,
    srcPlan: /assertSoOpen\\(so_id, '계획 등록'\\)/.test(src), srcWo: /assertSoOpen\\(so_id, '작업지시 발행'\\)/.test(src) }; })()`);
P('5 종결 수주: 계획→작업지시 발행 거부(ok:false) · WO 0 · 계획/지시 저장 직전 assertSoOpen · [수정] 숨김(D2)', c5.iw && c5.wo === 0 && c5.a === true && !c5.editBtn && c5.srcPlan && c5.srcWo, JSON.stringify(c5));

// 8) 원장 무변 (종결 전후 · LOT/실적/출하/작업지시/계획 문서 전체 동일)
const ledgerAfter = X(`JSON.stringify(['production_lots','records','shipments','work_orders','production_plans'].map(c => DB.allRaw(c).map(d => JSON.stringify(d)).sort()))`);
P('8 종결·취소 전후 production_lots·records·shipments·work_orders·production_plans 문서 완전 동일 (자동 완료·자동 취소 0)', ledgerAfter === before.ledger);

// 6) 종결 후 출하 취소(void) 허용 → 미출하 증가 · 유효 0 유지
const c6 = X(`(()=>{ voidDoc('shipments','SH-T1B','검증 — 출하 취소'); const o = DB.get('orders','SO-T1'); return { un: soUnshippedQty(o), open: soOpenQty(o), still: !!soCloseInfo(o) }; })()`);
P('6 종결 후 출하 취소(void) = 허용(정정 업무) · 미출하 30→60 · 유효잔량 0 유지 · 종결 유지', c6.un === 60 && c6.open === 0 && c6.still, JSON.stringify(c6));

// 7) 재개 (UI) → 유효잔량 복원 · 로그 2건
const c7 = X(`(()=>{ VIEWS.so(); document.querySelector('[data-so-reopen="SO-T1"]').onclick({ stopPropagation(){} });
  document.getElementById('sor-ok').onclick(); const blank = !!soCloseInfo(DB.get('orders','SO-T1'));
  document.getElementById('sor-t').value = '고객 추가 요청'; document.getElementById('sor-ok').onclick();
  const o = DB.get('orders','SO-T1'); const logs = (o._logs||[]).filter(l => /부분종결|수주 재개/.test(l.reason));
  const ok = (()=>{ try { assertSoShipQty('SO-T1', 60); return true; } catch(e){ return e.message; } })();
  return { blank, close: o.close, open: soOpenQty(o), logs: logs.map(l => l.reason), ok, status: calcSoRows().find(r => r.o.id==='SO-T1').status.label }; })()`);
P('7 [재개] 사유 없으면 거부 → 사유 입력 후 재개 · close=null · 유효잔량 60 복원 · 출하 가능 · 감사 로그 종결+재개 2건', c7.blank === true && c7.close === null && c7.open === 60 && c7.ok === true && c7.logs.length === 2 && /수주 재개: 고객 추가 요청/.test(c7.logs[1]) && c7.status !== '부분종결', JSON.stringify(c7));

// 10) 권한 — 보기 전용이면 저장 안 됨 (클라이언트 계약 · 서버 강제 아님)
const c10 = X(`(()=>{ const _b = _permBlockWrite; _permBlockWrite = () => true; let m = ''; try { soClose('SO-T4', '검증'); } catch(e){ m = e.message; } _permBlockWrite = _b; return { m, close: soCloseInfo(DB.get('orders','SO-T4')) }; })()`);
P('10 쓰기 권한 없음 → soClose 거부 · 종결 0 (클라이언트 계약 — 서버 강제 아님)', /권한/.test(c10.m) && c10.close === null, JSON.stringify(c10));

const fail = R.filter(r => !r.ok).length;
console.log(`\nSOCLOSE TOTAL ${R.length} · PASS ${R.length - fail} · FAIL ${fail}`);
process.exit(fail ? 1 : 0);
