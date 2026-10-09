// ★ v0.38.15 — 등록 후 수정 1차 회귀 (SPEC 「등록 후 수정 v0.1」 · GPT D1~D4 · GPT 지정 테스트)
// Usage: node _verify-docedit.js
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
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
const T = fn => { try { return fn(); } catch(e){ return 'ERR ' + e.message; } };

X(`seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='검증'; SESSION.userId='U-T';`);
X(`(()=>{
  const m = _tsMonth(Date.now()); window.__m = m;
  const lm = _monthAdd(m, -1); window.__lm = lm;
  DB.set('purchase_orders','PO-T',{ id:'PO-T', vendor:'S001', due:'2026-12-31', items:[{kind:'mat', name:'SCM440', spec:'Ø55', unit:'kg', qty:100, price:1000}], ts:Date.now(), by:'검증' });
  DB.set('material_receipts','MR-T',{ id:'MR-T', heat:'H-T1', mat:'SCM440', dia:55, kg:60, vendor:'S001', po_id:'PO-T', po_item_key:'SCM440|Ø55', date:fmtDate(Date.now()), ts:Date.now(), by:'검증' });
  DB.set('material_receipts','MR-U',{ id:'MR-U', heat:'H-T2', mat:'SCM440', dia:55, kg:30, vendor:'S001', po_id:null, po_item_key:null, date:fmtDate(Date.now()), ts:Date.now(), by:'검증' });
  DB.set('purchase_receipts','PR-T',{ id:'PR-T', name:'장갑', spec:'L', unit:'ea', qty:5, price:100, vendor:'S001', po_id:null, po_item_key:null, date:fmtDate(Date.now()), ts:Date.now(), by:'검증' });
  const lot = DB.all('production_lots')[0];
  DB.set('os_dispatches','OD-T',{ id:'OD-T', plot_no:lot.plot_no, proc:'HEAT', vendor:'V001', qty:10, pallets:0, ts:Date.now() - 5*86400000, by:'검증' });
  DB.set('os_receipts','OR-T',{ id:'OR-T', dispatch_id:'OD-T', qty:0, good:0, defect:0, scrap:0, pallets_recovered:0, record_id:null, ts:Date.now() - 3*86400000 });
})()`);
const ledger = () => X(`JSON.stringify(['production_lots','records','shipments','work_orders','production_plans','material_lots','os_dispatches','os_receipts'].map(c => DB.allRaw(c).map(d => JSON.stringify(d)).sort()))`);
const L0 = ledger();

// 1 PO 100 → 80 (입고 60) 허용 · 2 100 → 50 차단
const p1 = X(`(()=>{ const po = DB.get('purchase_orders','PO-T'); const before = JSON.stringify(po);
  let ok80, ng50, after50;
  try { const patch = poEditValidate(po, { vendor:po.vendor, due:po.due, items:[{...po.items[0], qty:80}] }); DB.correct('purchase_orders','PO-T', patch, '발주 수정: 검증 80'); ok80 = true; } catch(e){ ok80 = e.message; }
  const st = calcPoStatus(DB.get('purchase_orders','PO-T'));
  const amt = DB.get('purchase_orders','PO-T').items.reduce((s,i)=>s+(+i.qty||0)*(+i.price||0),0);
  const snap = JSON.stringify(DB.get('purchase_orders','PO-T'));
  try { poEditValidate(DB.get('purchase_orders','PO-T'), { vendor:po.vendor, due:po.due, items:[{...po.items[0], qty:50}] }); ng50 = false; } catch(e){ ng50 = /이미 입고 60/.test(e.message); }
  after50 = JSON.stringify(DB.get('purchase_orders','PO-T')) === snap;
  const lg = (DB.get('purchase_orders','PO-T')._logs||[]).slice(-1)[0] || {};
  return { ok80, ord: st.totalOrdered, rcv: st.totalReceived, remain: st.items[0].remain, amt, ng50, after50, log: lg.op==='correct' && lg.by==='검증' && /검증 80/.test(lg.reason) && lg.prev && lg.next };
})()`);
P('1 PO 100→80 (입고 60) 허용 · 감사 로그(처리자·사유·변경 전/후)', p1.ok80 === true && p1.log, JSON.stringify(p1));
P('2 PO 100→50 (입고 60) 차단 · 실패 시 문서 불변', p1.ng50 === true && p1.after50 === true);
P('3 수정 전후 PO 금액·잔량 재계산 정확 (발주 80 · 입고 60 · 잔량 20 · 금액 80,000)', p1.ord === 80 && p1.rcv === 60 && p1.remain === 20 && p1.amt === 80000, JSON.stringify(p1));

// 4 마감 월: PO 단가(연결 입고 월 마감) · 소재 입고 날짜 · 공급처 — 차단
const p4 = X(`(()=>{
  DB.set('closings', _closingId('purchase', window.__m), { id:_closingId('purchase', window.__m), type:'purchase', month:window.__m, ts:Date.now() });
  const po = DB.get('purchase_orders','PO-T'); const s0 = JSON.stringify(po), m0 = JSON.stringify(DB.get('material_receipts','MR-U'));
  let price, date, vend, memo;
  try { poEditValidate(po, { vendor:po.vendor, due:po.due, items:[{...po.items[0], price:1200}] }); price = false; } catch(e){ price = /마감된 월/.test(e.message); }
  try { docEditSave('material_receipts','MR-U', { date: window.__lm + '-15' }, '날짜 정정'); date = false; } catch(e){ date = /마감된 월/.test(e.message); }
  try { docEditSave('material_receipts','MR-U', { vendor:'S002' }, '공급처 정정'); vend = false; } catch(e){ vend = /마감된 월/.test(e.message); }
  try { memo = docEditSave('material_receipts','MR-U', { memo:'비고만' }, '').ok === true; } catch(e){ memo = e.message; }
  const memoSaved = DB.get('material_receipts','MR-U').memo === '비고만';
  DB.set('closings', _closingId('purchase', window.__m), { void:{ ts:Date.now(), reason:'검증 해제' } });
  return { price, date, vend, memo, memoSaved, poSame: JSON.stringify(DB.get('purchase_orders','PO-T')) === s0 };
})()`);
P('4 마감 월 차단(D2): PO 단가(연결 입고 월 마감) · 소재 입고일 · 공급처 → 차단 / 비고는 허용', p4.price === true && p4.date === true && p4.vend === true && p4.memo === true && p4.memoSaved && p4.poSame, JSON.stringify(p4));

// 5 공급처 변경: PO 입고 연결 → 차단 · 연결 소재 입고 → 잠금 · 비연결 입고 → 허용 + 매입 마감 집계 거래처 일치
const p5 = X(`(()=>{
  const po = DB.get('purchase_orders','PO-T'); let poV, mrLinked, mrFree;
  try { poEditValidate(po, { vendor:'S002', due:po.due, items:po.items }); poV = false; } catch(e){ poV = /거래처 변경 불가/.test(e.message); }
  try { docEditSave('material_receipts','MR-T', { vendor:'S002' }, '정정'); mrLinked = false; } catch(e){ mrLinked = /발주 거래처 고정/.test(e.message); }
  try { mrFree = docEditSave('material_receipts','MR-U', { vendor:'S002' }, '공급처 정정').ok; } catch(e){ mrFree = e.message; }
  const b0 = calcClosing('purchase', window.__m).byParty['S003'];
  let prFree; try { prFree = docEditSave('purchase_receipts','PR-T', { vendor:'S003' }, '공급처 정정').ok; } catch(e){ prFree = e.message; }
  const b1 = calcClosing('purchase', window.__m).byParty['S003'];
  return { poV, mrLinked, mrFree, vendorNow: DB.get('material_receipts','MR-U').vendor, prFree, s003Before: b0 ? b0.amount : 0, s003After: b1 ? b1.amount : 0 };
})()`);
P('5 공급처: 입고 연결 PO 변경 차단 · 연결 소재 입고 잠금 · 비연결 입고 변경 허용 → 매입 집계 거래처 즉시 일치', p5.poV === true && p5.mrLinked === true && p5.mrFree === true && p5.vendorNow === 'S002' && p5.prFree === true && p5.s003After - p5.s003Before === 500, JSON.stringify(p5));

// 6 대상 변경(LOT·Heat·수주·수량) 차단 — 입력에 넣어도 무시 · 필드 정의에 없음
const p6 = X(`(()=>{
  const flag0 = getFlags().ship_qty_fix; setFlag('ship_qty_fix', false);   // ★ v0.38.21 — 출하 수량 정정(시험 기능)은 꺼진 상태가 v0.38.15 규칙
  const sh = DB.all('shipments')[0]; const sh0 = JSON.stringify(sh);
  const r1 = T(() => docEditValidate('shipments', sh.id, { qty: 1, plot_no:'X', so_id:'X' }));
  const r2 = T(() => docEditValidate('material_receipts', 'MR-T', { heat:'H-X', kg: 1 }));
  const keys = Object.keys(DOC_EDIT_SPEC).flatMap(c => DOC_EDIT_SPEC[c].fields({}).map(f => c + '.' + f.k));
  const forbidden = keys.filter(k => /\\.(qty|kg|good|defect|plot_no|heat|so_id|itemCode|proc|pallets|mat|dia|po_id)$/.test(k));
  setFlag('ship_qty_fix', true); const forbiddenOn = Object.keys(DOC_EDIT_SPEC).flatMap(c => DOC_EDIT_SPEC[c].fields({}).map(f => c + '.' + f.k)).filter(k => /\\.(qty|kg|good|defect|plot_no|heat|so_id|itemCode|proc|pallets|mat|dia|po_id)$/.test(k)); setFlag('ship_qty_fix', flag0);
  return { emptyShip: r1 && r1.empty === true, emptyMr: r2 && r2.empty === true, forbidden, forbiddenOn, same: JSON.stringify(DB.get('shipments', sh.id)) === sh0 };
  function T(f){ try { return f(); } catch(e){ return 'ERR ' + e.message; } }
})()`);
P('6 LOT·Heat·수주·수량·외주처 등 대상 변경 = 수정 경로 없음 (필드 정의 0 · 입력 무시 · 출하 문서 불변) · v0.38.21 출하 수량 정정 켜짐 = shipments.qty 1개만 추가 (대상 변경은 여전히 0)', p6.emptyShip && p6.emptyMr && p6.forbidden.length === 0 && JSON.stringify(p6.forbiddenOn) === '["shipments.qty"]' && p6.same, JSON.stringify(p6));

// 7 출하: 계산서 발행 월 차단 · 포장수량·비고 허용 · 날짜 허용(계산서 없는 월)
const p7 = X(`(()=>{
  const o = DB.all('orders')[0]; DB.set('shipments','SH-T',{ id:'SH-T', so_id:o.id, itemCode:o.itemCode, plot_no:'', qty:1, pallets:0, pack_qty_actual:200, ts:Date.now(), by:'검증' });
  const sh = DB.get('shipments','SH-T'); const cust = o.cust; const m = getClMonth(sh);
  const free = !getClosing('sales', m);
  const invId = 'TI-T'; DB.set('acct_invoices', invId, { id:invId, vendor:cust, month:m, amount:1, issue_date:m+'-28', ts:Date.now() });
  const day0 = (sh.date || fmtDate(sh.ts)).slice(8,10); const nd = m + '-' + (day0 === '10' ? '11' : '10');
  let inv; try { docEditSave('shipments', sh.id, { date: nd }, '날짜'); inv = false; } catch(e){ inv = /계산서 발행됨/.test(e.message) ? true : e.message; }
  let pack; try { pack = docEditSave('shipments', sh.id, { pack_qty_actual: '150', memo:'박스 교체' }, '').ok; } catch(e){ pack = e.message; }
  voidDoc('acct_invoices', invId, '검증');
  const qtySame = DB.get('shipments', sh.id).qty === sh.qty;
  return { free, inv, pack, packNow: DB.get('shipments', sh.id).pack_qty_actual, qtySame };
})()`);
P('7 출하: 계산서 발행 월 날짜 수정 차단(D2·D3) · 포장수량·비고 수정 허용 · 수량 불변', p7.free && p7.inv === true && p7.pack === true && p7.packNow === 150 && p7.qtySame, JSON.stringify(p7));

// 8 외주 반출일 > 첫 외주 입고 → 차단 · 외주 입고 = 비고만
const p8 = X(`(()=>{
  const d = DB.get('os_dispatches','OD-T');
  const first = DB.query('os_receipts', r => r.dispatch_id === d.id).map(r=>r.ts).sort((a,b)=>a-b)[0];
  let late; try { docEditSave('os_dispatches', d.id, { date: fmtDate(first + 5*86400000) }, '날짜'); late = false; } catch(e){ late = /첫 외주 입고/.test(e.message); }
  const orFields = DOC_EDIT_SPEC.os_receipts.fields({}).map(f=>f.k).join(',');
  return { late, orFields };
})()`);
P('8 외주 반출일은 첫 외주 입고보다 늦을 수 없음 · 외주 입고 = 비고만(수량·날짜는 실적 연동 2차)', p8.late === true && p8.orFields === 'memo', JSON.stringify(p8));

// 9 사유 필수(영향 필드) · 비고만이면 사유 없이 허용
const p9 = X(`(()=>{ const s0 = JSON.stringify(DB.get('purchase_receipts','PR-T')); let need; try { docEditSave('purchase_receipts','PR-T', { memo:'x', price:'777' }, ''); need = false; } catch(e){ need = /사유/.test(e.message); } const same = JSON.stringify(DB.get('purchase_receipts','PR-T')) === s0; let ok; try { ok = docEditSave('purchase_receipts','PR-T', { price:'777' }, '단가 정정').ok && DB.get('purchase_receipts','PR-T').price === 777; } catch(e){ ok = e.message; } return { need, same, ok }; })()`);
P('9 날짜·거래처·단가 변경은 사유 필수 (없으면 저장 0 · 문서 불변) · 사유 있으면 저장', p9.need === true && p9.same && p9.ok === true, JSON.stringify(p9));

// 10 원장 무변 — 속성 수정 전후 LOT·실적·재고·출하수량·외주 원장 (출하 pack/memo · 소재 입고 vendor/memo 제외 비교)
const L1 = X(`JSON.stringify(['production_lots','records','work_orders','production_plans','material_lots','os_receipts'].map(c => DB.allRaw(c).map(d => JSON.stringify(d)).sort()))`);
const L0b = JSON.stringify(JSON.parse(L0).filter((_, i) => [0,1,3,4,5,7].includes(i)));
P('10 원장 무변 — LOT·실적·작업지시·계획·소재 LOT·외주 입고 문서 수정 전후 동일', L1 === L0b);

// 11 UI — 목록 [수정] 버튼 · 발주 수정 창 저장 경로
const p11 = X(`(()=>{
  location.hash = '#po'; VIEWS.po(); const poBtn = !!document.querySelector('[data-doc-edit="purchase_orders|PO-T"]');
  VIEWS._poEdit('PO-T'); const rows = document.querySelectorAll('#poe-rows tr').length; const nameLocked = document.querySelector('[data-pe="0|name"]').disabled;
  const q = document.querySelector('[data-pe="0|qty"]'); q.value = '90'; q.oninput(); document.getElementById('poe-r').value = '창 저장 검증'; document.getElementById('poe-ok').onclick();
  const saved = DB.get('purchase_orders','PO-T').items[0].qty === 90;
  location.hash = '#ship'; VIEWS.ship(); const shBtn = document.querySelectorAll('[data-doc-edit^="shipments|"]').length > 0;
  VIEWS._docEdit('shipments', DB.all('shipments')[0].id); const shModal = /출하 수정/.test(document.body.textContent) && /취소\\] 후 재등록/.test(document.body.textContent); closeModal();
  return { poBtn, rows, nameLocked, saved, shBtn, shModal };
})()`);
P('11 UI: 발주·출하 목록 [수정] · 발주 수정 창(입고 연결 품목 이름 잠금) → 저장 90 · 출하 수정 창 「수량은 취소 후 재등록」 안내', p11.poBtn && p11.rows === 1 && p11.nameLocked && p11.saved && p11.shBtn && p11.shModal, JSON.stringify(p11));

const fail = R.filter(r => !r.ok).length;
console.log(`\nDOCEDIT TOTAL ${R.length} · PASS ${R.length - fail} · FAIL ${fail}`);
process.exit(fail ? 1 : 0);
