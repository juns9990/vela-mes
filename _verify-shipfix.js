// ★ v0.38.21/22 — 출하 수량 정정 회귀 (v0.38.22 = GPT 판정 반영: 서버 Gate · 단일 문서 · 필수 연결 · 정합성 표시) (SPEC 「등록 후 수정 v0.3」 · GPT D5~D10 · 기본 끔 기능 플래그 ship_qty_fix)
// Usage: node _verify-shipfix.js   (VELA_FILE=index-dev.html 지원)
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
const T = fn => { try { return fn(); } catch(e){ return 'ERR ' + e.message; } };
const E = s => X(`(()=>{ try { ${s}; return 'OK'; } catch(e){ return 'ERR ' + e.message; } })()`);
X(`seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='검증'; SESSION.userId='U-T';`);
const L0 = X(`JSON.stringify(['shipments','records','production_lots','orders'].map(c => DB.allRaw(c).length))`);
// 픽스처: 남은 양품이 넉넉한 LOT + 새 수주 + 이번 달 출하 40
const fx = X(`(()=>{
  const pick = DB.all('production_lots').map(l => { const r = getRoute(l.itemCode); const lp = r.length > 1 ? r[r.length-2] : r[0]; const bal = LOT.acceptedGood(l.plot_no, lp.proc); const sh = DB.query('shipments', x => x.plot_no === l.plot_no).reduce((a,x)=>a+(+x.qty||0),0); return { l, avail: bal - sh, bal, sh }; })
    .filter(x => x.avail >= 120 && !lotHoldInfo(x.l.plot_no).held).sort((a,b) => b.avail - a.avail)[0];
  const lot = pick.l; const cust = DB.all('orders')[0].cust;
  DB.set('orders','SO-FX',{ id:'SO-FX', cust, itemCode:lot.itemCode, qty:100, price:3000, due:'2099-12-31', created:Date.now() });
  DB.set('shipments','SH-FX',{ id:'SH-FX', so_id:'SO-FX', itemCode:lot.itemCode, plot_no:lot.plot_no, qty:40, pallets:1, pack_qty_actual:200, ts:Date.now(), by:'검증' });
  window.__fx = { plot: lot.plot_no, avail: pick.avail, bal: pick.bal, item: lot.itemCode, cust };
  return window.__fx; })()`);
const A = () => X(`({ qty: DB.get('shipments','SH-FX').qty, hist: (DB.get('shipments','SH-FX').qty_hist||[]).length, amt: shipAmount(DB.get('shipments','SH-FX')), soShip: soShippedQty(DB.get('orders','SO-FX')), lotShip: DB.query('shipments', x => x.plot_no === window.__fx.plot).reduce((a,x)=>a+(+x.qty||0),0) })`);

// F0 기능 꺼짐 (기본)
const f0 = X(`(()=>{ const demoOn = getFlags().ship_qty_fix; setFlag('ship_qty_fix', false); const flagDefault = defaultFlags().ship_qty_fix || EXPERIMENTAL_FLAGS.indexOf('ship_qty_fix') < 0; const fields = DOC_EDIT_SPEC.shipments.fields(DB.get('shipments','SH-FX')).map(f=>f.k); const v = docEditValidate('shipments','SH-FX',{ qty:'30' });
  let thr = ''; try { shipQtyFixCheck('SH-FX', 30); } catch(e){ thr = e.message; } return { demoOn, flagDefault, fields, patch: v.patch, empty: v.empty, thr, deps: JSON.stringify(FEATURE_DEPS.ship_qty_fix) }; })()`);
P('F0 실데이터 기본 = 꺼짐(시험 기능 · 풀옵션 프리셋 제외 · 데모만 켜짐) · 끄면 수량 칸 없음 · 수량 입력 무시 · 검사 함수 거부 · 의존성 ship', f0.demoOn === true && f0.flagDefault === false && !f0.fields.includes('qty') && f0.empty === true && /꺼져/.test(f0.thr) && f0.deps === '{"needs":["ship"],"breaks":[]}', JSON.stringify(f0));
// G1 (GPT v0.38.21 ①) 실데이터(데모 아님) = 플래그를 켜도 저장 차단 · 수량 칸 없음
X(`setFlag('ship_qty_fix', true); 1`);
const g1 = X(`(()=>{ const demo = isDemoMode(); const fields = DOC_EDIT_SPEC.shipments.fields(DB.get('shipments','SH-FX')).map(f=>f.k); const locked = DOC_EDIT_SPEC.shipments.locked(DB.get('shipments','SH-FX')).map(x=>x.join(':')).join('|');
  let e=''; try { shipQtyFixCheck('SH-FX', 30); } catch(x){ e = x.message; } const v = docEditValidate('shipments','SH-FX',{ qty:'30' }); return { demo, fields, locked, e, empty: v.empty, gate: SERVER_GATES.ship_qty_fix }; })()`);
P('G1 실데이터(데모 아님): 플래그 켜도 수량 칸 없음 · 저장 차단(서버 Gate) · 입력 무시 · SERVER_GATES=false', g1.demo === false && !g1.fields.includes('qty') && /데모에서만/.test(g1.locked) && /서버 권한·동시성/.test(g1.e) && g1.empty === true && g1.gate === false, JSON.stringify(g1));
X(`sessionStorage.setItem('vela_demo','1'); 1`);   // 이후 = 데모(시험) 환경
const a0 = A();
// F1 감소 정정
const r1 = E(`docEditSave('shipments','SH-FX',{ qty:'30' }, '수량 오타')`);
const a1 = A();
const lg = X(`(()=>{ const d = DB.get('shipments','SH-FX'); const l = (d._logs||[]).filter(x => x.op === 'correct').pop() || {}; return { prev: l.prev && l.prev.qty, next: l.next && l.next.qty, by: l.by, reason: l.reason, h: d.qty_hist[0] }; })()`);
P('F1 감소 정정 40 → 30: 같은 문서 · 매출 −30,000 · 수주 출하 −10 · LOT 출하 −10 · 출하 문서 수 그대로', r1 === 'OK' && a1.qty === 30 && a1.amt - a0.amt === -30000 && a0.soShip - a1.soShip === 10 && a0.lotShip - a1.lotShip === 10, JSON.stringify({ a0, a1 }));
P('F2 감사 이력: _logs 전·후 수량·처리자·사유 + qty_hist(추가 전용) 1건', lg.prev === 40 && lg.next === 30 && lg.by === '검증' && /수량 정정 40 → 30 ea/.test(lg.reason) && /수량 오타/.test(lg.reason) && lg.h.from === 40 && lg.h.to === 30 && lg.h.by === '검증', JSON.stringify(lg));
P('F3 사유 없으면 거부 · 0·음수·소수·빈값 거부', /사유/.test(E(`docEditSave('shipments','SH-FX',{ qty:'31' }, '')`)) && ['0','-5','12.5','abc'].every(v => /정수|확인/.test(E(`docEditSave('shipments','SH-FX',{ qty:'${v}' }, 'x')`))) && A().qty === 30, '');
// F4 증가: 수주 초과 / LOT 초과 / 정상
const e4a = E(`docEditSave('shipments','SH-FX',{ qty:'101' }, '증가')`);
X(`DB.set('orders','SO-FX',{ qty: 100000 }); 1`);
const e4b = E(`docEditSave('shipments','SH-FX',{ qty:'${fx.avail + 41}' }, '증가')`);
X(`DB.set('orders','SO-FX',{ qty: 100 }); 1`);
const e4c = E(`docEditSave('shipments','SH-FX',{ qty:'60' }, '증가 정정')`);
P('F4 증가 정정: 수주 초과 차단(최대 표시) · LOT 남은 양품 초과 차단 · 범위 안 30 → 60 저장', /수주 수량 초과/.test(e4a) && /최대 100/.test(e4a) && /LOT 남은 양품 초과/.test(e4b) && e4c === 'OK' && A().qty === 60, JSON.stringify({ e4a, e4b, e4c }));
// F5 HOLD: 증가 차단 · 감소 허용
const h = X(`(()=>{ const n = qcRegisterNc({ lot_id: window.__fx.plot, item_code: window.__fx.item, defect_type: (DB.all('defect_types')[0]||{}).code || 'D', qty: 1, hold: true, hold_reason:'검증 HOLD' }); window.__hn = n.id;
  let up = 'OK'; try { docEditSave('shipments','SH-FX',{ qty:'61' }, 'HOLD 중 증가'); } catch(e){ up = e.message; }
  let down = 'OK'; try { docEditSave('shipments','SH-FX',{ qty:'59' }, 'HOLD 중 감소'); } catch(e){ down = e.message; }
  voidDoc('nc_records', n.id, '검증'); return { up, down, qty: DB.get('shipments','SH-FX').qty }; })()`);
P('F5 품질 HOLD LOT: 증가 정정 차단 · 감소 정정은 허용', /HOLD/.test(h.up) && h.down === 'OK' && h.qty === 59, JSON.stringify(h));
// F6 (GPT P0) 단일 문서 쓰기: NC·사고 문서 무변 · NC 상세는 qty_hist 파생 표시 · 판정 대기 LOT 증가 정정 차단
const q = X(`(()=>{
  const n = qcRegisterNc({ lot_id: window.__fx.plot, item_code: window.__fx.item, defect_type: (DB.all('defect_types')[0]||{}).code || 'D', qty: 1, hold: false });
  DB.set('nc_records', n.id, { pending_out: [{ kind:'ship', ref_id:'SH-FX', qty:59, ts:Date.now(), by:'검증' }] });
  const nc0 = JSON.stringify(DB.get('nc_records', n.id));
  docEditSave('shipments','SH-FX',{ qty:'55' }, '추적 정정');
  const same = JSON.stringify(DB.get('nc_records', n.id)) === nc0;
  let up = 'OK'; try { docEditSave('shipments','SH-FX',{ qty:'56' }, '판정 대기 중 증가'); } catch(e){ up = e.message; }
  VIEWS._qcDetail(n.id); const tl = document.getElementById('modal').textContent; closeModal();
  voidDoc('nc_records', n.id, '검증');
  return { same, up, tl: /출하 수량 정정 · SH-FX · 59 → 55ea \\(현재 유효 55ea\\)/.test(tl), qty: DB.get('shipments','SH-FX').qty }; })()`);
P('F6 (GPT P0) 정정 = 출하 문서 1건만 쓰기 · NC 문서 무변 · NC 상세는 qty_hist 파생 「59 → 55ea (현재 유효 55ea)」 · 판정 대기 LOT 증가 정정 차단', q.same && /판정 대기·동시성 보류/.test(q.up) && q.tl && q.qty === 55, JSON.stringify(q));
// G2 (GPT P1) 필수 연결 누락 차단
const g2 = X(`(()=>{ const out = {};
  DB.set('shipments','SH-FXN',{ id:'SH-FXN', so_id:'SO-FX', itemCode:window.__fx.item, plot_no:'', qty:1, pallets:0, ts:Date.now(), by:'검증' });
  try { shipQtyFixCheck('SH-FXN', 2); out.noPlot = 'OK'; } catch(e){ out.noPlot = e.message; }
  const it = DB.get('orders','SO-FX').itemCode; DB.set('orders','SO-FX',{ itemCode:'' });
  try { shipQtyFixCheck('SH-FX', 54); out.noItem = 'OK'; } catch(e){ out.noItem = e.message; }
  DB.set('orders','SO-FX',{ itemCode: it }); voidDoc('shipments','SH-FXN','검증'); return out; })()`);
P('G2 (GPT P1) 수주·LOT 연결 비어 있음 · 품번 비어 있음 → 정정 차단 (불일치 검사 우회 없음)', /연결이 비어/.test(g2.noPlot) && /품번이 비어/.test(g2.noItem), JSON.stringify(g2));
// F7 마감 월 · 계산서 월 · 종결 수주 · 취소 문서
const m = X(`_tsMonth(Date.now())`);
const e7a = X(`(()=>{ DB.set('closings', _closingId('sales','${m}'), { id:_closingId('sales','${m}'), type:'sales', month:'${m}', ts:Date.now() }); let e='OK'; try { docEditSave('shipments','SH-FX',{ qty:'54' }, '마감 중'); } catch(x){ e = x.message; } DB.set('closings', _closingId('sales','${m}'), { void:{ ts:Date.now(), reason:'검증 해제' } }); return e; })()`);
const e7b = X(`(()=>{ DB.set('acct_invoices','TI-FX',{ id:'TI-FX', vendor: window.__fx.cust, month:'${m}', amount:1, issue_date:'${m}-28', ts:Date.now() }); let e='OK'; try { docEditSave('shipments','SH-FX',{ qty:'54' }, '계산서'); } catch(x){ e = x.message; } voidDoc('acct_invoices','TI-FX','검증'); return e; })()`);
const e7c = X(`(()=>{ soClose('SO-FX', '검증 종결'); let e='OK'; try { docEditSave('shipments','SH-FX',{ qty:'54' }, '종결'); } catch(x){ e = x.message; } soReopen('SO-FX', '검증 재개'); return e; })()`);
P('F7 차단: 판매 마감 월 · 계산서 발행 월 · 종결 수주 (해제 후 원장 무변)', /마감/.test(e7a) && /계산서/.test(e7b) && /종결|재개/.test(e7c) && A().qty === 55, JSON.stringify({ e7a, e7b, e7c }));
// F8 다른 곳 변경 감지 · 취소 문서 · 정정 횟수 상한
const e8a = X(`(()=>{ const u = DB.get('shipments','SH-FX')._updated; DB.set('shipments','SH-FX',{ memo:'다른 단말' }); let e='OK'; try { docEditSave('shipments','SH-FX',{ qty:'54' }, '경합', { expectUpdated: u }); } catch(x){ e = x.message; } return e; })()`);
const e8b = X(`(()=>{ DB.set('shipments','SH-FX2',{ id:'SH-FX2', so_id:'SO-FX', itemCode:window.__fx.item, plot_no:window.__fx.plot, qty:1, pallets:0, ts:Date.now(), by:'검증' }); voidDoc('shipments','SH-FX2','검증'); let e='OK'; try { shipQtyFixCheck('SH-FX2', 2); } catch(x){ e = x.message; } return e; })()`);
const e8c = X(`(()=>{ const out = []; for(let i=0;i<4;i++){ try { docEditSave('shipments','SH-FX',{ qty:String(54-i) }, '반복 ' + i); out.push('OK'); } catch(x){ out.push(x.message); } } return { out, hist: DB.get('shipments','SH-FX').qty_hist.length }; })()`);
P('F8 창을 연 뒤 다른 곳 변경 → 저장 거부 · 취소된 출하 정정 불가 · 정정 5회 초과 → [취소] 후 재등록 안내', /다른 곳/.test(e8a) && /취소된/.test(e8b) && e8c.hist === 5 && /5회/.test(e8c.out[e8c.out.length-1]), JSON.stringify({ e8a, e8b, e8c }));
// F9 납기준수 영향 기록 (수량 감소로 준수 → 일부 출하)
const o9 = X(`(()=>{ const K = s => Date.parse(s + '+09:00');
  DB.set('orders','SO-FX9',{ id:'SO-FX9', cust:window.__fx.cust, itemCode:window.__fx.item, qty:10, price:1, due: _kstDate(Date.now() - 86400000*2) });
  DB.set('shipments','SH-FX9',{ id:'SH-FX9', so_id:'SO-FX9', itemCode:window.__fx.item, plot_no:window.__fx.plot, qty:10, pallets:0, ts: Date.now() - 86400000*3, by:'검증' });
  const im = otdShipPatchImpact('SH-FX9', { qty: 8 }); let e='OK'; try { docEditSave('shipments','SH-FX9',{ qty:'8' }, '수량 오타'); } catch(x){ e = x.message; }
  const l = (DB.get('shipments','SH-FX9')._logs||[]).filter(x=>x.op==='correct').pop() || {};
  return { im, e, reason: l.reason }; })()`);
P('F9 수량 정정이 납기준수 판정을 바꾸면 미리 알림 + 감사 사유에 「납기준수 영향」 기록', o9.im && o9.im.changed && /^준수/.test(o9.im.before) && o9.e === 'OK' && /납기준수 영향: SO-FX9 준수/.test(o9.reason || ''), JSON.stringify(o9));
// F10 화면: 수정 창 미리보기 · 오류 표시 · 목록 ✎ 정정
const ui = X(`(()=>{ VIEWS._docEdit('shipments','SH-FX9'); const inp = document.querySelector('[data-de="qty"]'); const has = !!inp;
  inp.value = '7'; inp.dispatchEvent(new window.Event('input')); const t1 = document.getElementById('de-qfix').textContent;
  inp.value = '999'; inp.dispatchEvent(new window.Event('input')); const t2 = document.getElementById('de-qfix').textContent;
  const locked = document.querySelector('.modal, #modal').textContent; closeModal();
  window._shipVoidView = { voidView:false }; window._shTab = 'list'; location.hash = '#ship'; router();
  const row = [...document.querySelectorAll('#view tr')].find(tr => /SO-FX9/.test(tr.textContent));
  return { has, t1, t2, badge: !!row && /✎ 정정/.test(row.textContent), noFixed: !/출하 수량 · 고정/.test(locked) }; })()`);
P('F10 수정 창: 수량 칸 · 「8 → 7 ea」 미리보기(매출·수주 잔량·LOT·명세서 재발행·거래처 통보·파렛트 안내) · 초과는 「저장 불가」 · 목록 ✎ 정정', ui.has && /8 → 7 ea/.test(ui.t1) && /매출/.test(ui.t1) && /거래명세서/.test(ui.t1) && /알려/.test(ui.t1) && /파렛트/.test(ui.t1) && /저장 불가/.test(ui.t2) && ui.badge && ui.noFixed, JSON.stringify(ui));
// F10b 기존 원장이 이미 초과(레거시·데모)인 LOT: 감소 정정은 허용 + 경고 기록 · 증가는 차단
const ov = X(`(()=>{ const K = Date.now();
  DB.set('shipments','SH-FXO',{ id:'SH-FXO', so_id:'SO-FX', itemCode:window.__fx.item, plot_no:window.__fx.plot, qty: window.__fx.bal + 5, pallets:0, ts:K, by:'검증' });   // 원장 초과 상태 재현
  let up='OK', down='OK'; try { docEditSave('shipments','SH-FXO',{ qty:String(window.__fx.bal + 6) }, '증가'); } catch(e){ up = e.message; }
  try { docEditSave('shipments','SH-FXO',{ qty:String(window.__fx.bal + 1) }, '감소'); } catch(e){ down = e.message; }
  const l = (DB.get('shipments','SH-FXO')._logs||[]).filter(x=>x.op==='correct').pop() || {}; voidDoc('shipments','SH-FXO','검증');
  return { up, down, reason: l.reason }; })()`);
P('F10b 기존 초과 원장(레거시): 증가 정정 차단 · 감소 정정 허용 + 「⚠ 정정 후에도 … 초과 — 확인 필요」 감사 기록', /초과/.test(ov.up) && ov.down === 'OK' && /⚠ .*확인 필요/.test(ov.reason || ''), JSON.stringify(ov));
// G3 (GPT ②) 원장 정합성 화면: LOT 출하 초과 = 감사 화면 잠금 단락에 계속 표시
const g3 = X(`(()=>{ DB.set('shipments','SH-FXO2',{ id:'SH-FXO2', so_id:'SO-FX', itemCode:window.__fx.item, plot_no:window.__fx.plot, qty: window.__fx.bal + 7, pallets:0, ts:Date.now(), by:'검증' });
  const ov = lotShipOverage(window.__fx.plot); location.hash = '#audit'; router(); const s = secList().find(x => x.key === 'audit:audit-ship-overage');
  const listed = s && s.card.textContent.includes(window.__fx.plot); voidDoc('shipments','SH-FXO2','검증'); return { over: ov.over > 0, lockBy: s && s.lockBy, listed }; })()`);
P('G3 (GPT ②) 감사 › 「원장 정합성 · LOT 출하 초과」 단락(안전 잠금)에 미해결 초과 LOT 표시 · 파생(저장 0)', g3.over && g3.lockBy === 'mark' && g3.listed, JSON.stringify(g3));
// F11 원장 무변 (출하 정정은 새 문서를 만들지 않음)
const L1 = X(`JSON.stringify(['shipments','records','production_lots','orders'].map(c => DB.allRaw(c).length))`);
const b = JSON.parse(L0), a = JSON.parse(L1);
P('F11 출하 문서 = 픽스처 6건만 증가(정정으로 새 출하 0) · 실적·LOT 무변', a[0] - b[0] === 6 && a[1] === b[1] && a[2] === b[2] && a[3] - b[3] === 2, L0 + ' → ' + L1);
P('버전 v0.38.23', X(`APP_VERSION`) === 'v0.38.23', X(`APP_VERSION`));
if(R.length !== 17) R.push({ n:`항목 수 ${R.length} ≠ 기대 17`, ok:false });
const fail = R.filter(x => !x.ok).length;
console.log(`\nSHIPFIX TOTAL ${R.length} · PASS ${R.length - fail} · FAIL ${fail}`);
process.exit(fail ? 1 : 0);
