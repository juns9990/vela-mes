// ★ v0.35.10 (9-B) — 품질 판정 상태 머신 회귀 (GPT 정책 심사 §11 필수 30항목 · 원장 무변)
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window;
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.prompt = () => 'QA 사유';
win.VELA_BACKEND = 'local';   // 데모·테스트 모드와 동일(local) — 드롭존 인라인 폴백 검증용
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const body = `
;(function(){
  const R = window.__T = [];
  const P = (n, ok, note) => R.push({ n, ok, note: String(note||'') });
  const tryRun = f => { try { f(); return null; } catch(e){ return e.message; } };
  const T = s => s.replace(/\\s+/g,' ');
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa';
    const dt = DB.all('defect_types').find(d => d.active !== false).code;
    // 준비: 완성 LOT (출하 대기) — 실적 있는 마지막 공정 실적 1건 골라 부적합 연결
    const lot = DB.all('production_lots').find(l => { const nx = LOT.nextProcess(l.plot_no); return nx && nx.code === 'SHIP' && !heldPlotSet().has(l.plot_no) && !lockedPlotSet().has(l.plot_no); });
    const PLOT = lot.plot_no; const route = getRoute(lot.itemCode); const lastP = route[route.length-2].proc;
    const shipped0 = DB.query('shipments', x => x.plot_no === PLOT).reduce((s,r)=>s+(r.qty||0),0);
    const avail0 = LOT.totalGood(PLOT, lastP) - shipped0;
    // 실적에 불량 40 넣기 (보정 경로) → 실적 연결 NC 상한 검증용
    const rec = DB.query('records', r => r.plot_no === PLOT && r.proc === lastP)[0];
    LOT.correctRecord(rec.id, { good: rec.good - 40, defect: 40, rework: 0 }, 'QA 불량 40');
    const avail = LOT.totalGood(PLOT, lastP) - shipped0;
    // ── 상태 전환 ──
    const n1 = qcRegisterNc({ lot_id: PLOT, record_id: rec.id, item_code: lot.itemCode, process: lastP, defect_type: dt, qty: 30, hold: false });
    P('1 현장 NC(hold=false) → 판정 대기 · LOT 미차단', ncIsPending(DB.get('nc_records', n1.id)) && !lotHoldInfo(PLOT).held && pendingPlotSet().has(PLOT), n1.id);
    P('1b 판정 대기 LOT 은 실적 등록 가능 (registerRecord 미차단)', tryRun(() => assertLotNotHeld(PLOT, '실적 등록')) === null, '');
    const seedHold = DB.all('nc_records').find(n => n.status === '접수' && n.hold === true && n.id.startsWith('NC-2026-DA'));
    P('1c 시드 접수 NC 는 hold:true 명시 = HOLD 시연 케이스 유지 (U1·K1-3 시나리오 보존)', !!seedHold && ncIsHold(seedHold) && heldPlotSet().has(seedHold.lot_id), seedHold && seedHold.id);
    const legacy = { id:'NC-LEGACY', status:'접수', lot_id: PLOT, qty: 1 };
    P('1d hold 필드 없는 기존 NC(실DB 업데이트 전 등록분) = 판정 대기 (사용자 판정: 기본 전부 HOLD 아님)', ncIsPending(legacy) && !ncIsHold(legacy), '');
    ncSetHold(n1.id, true, '품질팀 판단');
    P('2 품질팀 HOLD 체크 → 실적·출하·반출 3중 차단 (기존 Gate)', lotHoldInfo(PLOT).held && ['실적 등록','출하','외주 반출'].every(a => /품질 HOLD/.test(tryRun(() => assertLotNotHeld(PLOT, a)) || '')), '');
    ncSetHold(n1.id, false, '오등록 확인');
    P('2b HOLD 해제 → 판정 대기 복귀', !lotHoldInfo(PLOT).held && ncIsPending(DB.get('nc_records', n1.id)), '');
    // ── 격리 계산 ──
    const n2 = qcRegisterNc({ lot_id: PLOT, record_id: rec.id, item_code: lot.itemCode, process: lastP, defect_type: dt, qty: 10, hold: false });   // 30+10 ≤ 실적 불량 40
    P('4 실적 NC 30+10 / 실적 불량 40 → 격리 40 (실적 상한)', lotQuarantineQty(PLOT, avail) === 40, lotQuarantineQty(PLOT, avail));
    const e40 = tryRun(() => qcRegisterNc({ lot_id: PLOT, record_id: rec.id, item_code: lot.itemCode, defect_type: dt, qty: 30, hold: false }));
    P('4b 같은 실적 초과 신고(30 추가 → 70 > 40)는 등록 단계에서 차단 (ncSumByRecord)', typeof e40 === 'string' && /초과/.test(e40), e40);
    const n3 = qcRegisterNc({ lot_id: PLOT, item_code: lot.itemCode, defect_type: dt, qty: 500, hold: false });   // LOT 단위 NC
    P('5 LOT NC 500 / 가용 ' + avail + ' → LOT NC 는 가용 상한 · 합계도 가용 상한', lotQuarantineQty(PLOT, avail) === avail, lotQuarantineQty(PLOT, avail));
    voidDoc('nc_records', n3.id, 'QA');
    const n4 = qcRegisterNc({ lot_id: PLOT, item_code: lot.itemCode, defect_type: dt, qty: 25, hold: false });
    P('6 실적 NC(40) + LOT NC(25) 혼합 → 65 (가용 이내)', lotQuarantineQty(PLOT, avail) === 65 && 65 < avail, lotQuarantineQty(PLOT, avail));
    DB.set('nc_records', n4.id, { status: '재작업' });
    P('7 재작업 중 NC → 격리 유지', lotQuarantineQty(PLOT, avail) === 65, lotQuarantineQty(PLOT, avail));
    DB.set('nc_records', n4.id, { status: '완료' }); DB.set('nc_records', n2.id, { status: '특채' });
    P('8 처리 완료(완료·특채) → 격리에서 제외 (남은 판정 대기 30만)', lotQuarantineQty(PLOT, avail) === 30, lotQuarantineQty(PLOT, avail));
    // ── 출하/반출 ──
    const so = DB.all('orders').find(o => o.itemCode === lot.itemCode) || DB.all('orders')[0];
    const allowed = avail - 30;
    const ship = (qty) => { const id = 'SH-QA-' + Date.now() + '-' + Math.random().toString(36).slice(2,5); const q = assertQuarantine(PLOT, qty, '출하', LOT.totalGood(PLOT, lastP) - DB.query('shipments', x=>x.plot_no===PLOT).reduce((s,r)=>s+(r.qty||0),0)); DB.set('shipments', id, { id, so_id: so.id, itemCode: lot.itemCode, plot_no: PLOT, qty, ts: Date.now(), by:'qa' }); ncRecordPendingOut(q.pending, 'ship', id, qty); return id; };
    const e9 = tryRun(() => ship(allowed - 5));
    P('9 판정 대기 → 허용 수량(가용−격리) 이내 출하 성공', e9 === null, e9 || ('allowed ' + allowed));
    const e10 = tryRun(() => assertQuarantine(PLOT, 6, '출하', 5 + 30));   // 남은 가용 35 · 격리 30 → 허용 5
    P('10 허용 초과 출하 → 차단 (문구에 격리·가능 수량)', typeof e10 === 'string' && /판정 대기 격리/.test(e10) && /가능 5ea/.test(e10), e10);
    const os1 = tryRun(() => assertQuarantine(PLOT, 3, '외주 반출', 35));
    P('11 판정 대기 → 허용 이내 반출 통과', os1 === null, os1 || '');
    const os2 = tryRun(() => assertQuarantine(PLOT, 6, '외주 반출', 35));
    P('12 허용 초과 반출 → 차단 (출하와 동일 규칙 · 완화 없음)', typeof os2 === 'string' && /외주 반출 가능 5ea/.test(os2), os2);
    // ── 최종 write path 재검증 ──
    const n5 = qcRegisterNc({ lot_id: PLOT, item_code: lot.itemCode, defect_type: dt, qty: 5, hold: false });   // 격리 35 → 허용 0
    P('13 출하 직전 격리 증가 → 같은 요청(3ea) 재검증 시 차단', typeof tryRun(() => assertQuarantine(PLOT, 3, '출하', 35)) === 'string', '');
    P('14 반출 직전 격리 증가 → 차단', typeof tryRun(() => assertQuarantine(PLOT, 3, '외주 반출', 35)) === 'string', '');
    voidDoc('nc_records', n5.id, 'QA');
    ncSetHold(n1.id, true, '직전 HOLD');
    P('15/16 출하·반출 직전 HOLD 전환 → assertLotNotHeld 차단', /품질 HOLD/.test(tryRun(() => assertLotNotHeld(PLOT, '출하'))||'') && /품질 HOLD/.test(tryRun(() => assertLotNotHeld(PLOT, '외주 반출'))||''), '');
    ncSetHold(n1.id, false, '복귀');
    P('17 동시성 판정은 별도 Gate 그대로 (v0.36.0: assertWithinAccepted → assertQuarantine 순서 · 출하 핸들러 내)', typeof assertWithinAccepted === 'function' && /assertWithinAccepted\\(plot, qty, '출하'[\\s\\S]*?assertQuarantine\\(plot, qty, '출하'/.test(Array.from(document.querySelectorAll('script')).map(s=>s.textContent).join('')), '');
    // ── UI/감지 ──
    location.hash = '#ship'; router(); document.getElementById('btn-ship-new').onclick();
    const soSel = document.getElementById('sh-so'); const opt = [...soSel.options].find(o => o.dataset.i === lot.itemCode); if(opt){ soSel.value = opt.value; soSel.dispatchEvent(new window.Event('change')); }
    const lotSel = document.getElementById('sh-lot'); lotSel.value = PLOT; lotSel.dispatchEvent(new window.Event('change'));
    const warn = T(document.getElementById('sh-warn').textContent);
    P('18 출하 창 판정 대기 경고 문구 (NC n건 · 격리 m · 출하 가능 k)', /판정 대기 NC 1건/.test(warn) && /격리 30ea/.test(warn) && /출하 가능 5ea/.test(warn), warn.slice(0,120));
    P('18b LOT 옵션 라벨 ⚠ 판정 대기', /판정 대기/.test([...lotSel.options].find(o => o.value === PLOT).textContent), '');
    closeModal();
    DB.set('nc_records', n1.id, { ts: Date.now() - 1.2 * 86400000 });
    P('19 판정 지체 1일 → 🟡 (기본 warn=1)', ncPendingLevel(DB.get('nc_records', n1.id)) === 1, '');
    DB.set('nc_records', n1.id, { ts: Date.now() - 3.2 * 86400000 });
    P('20 판정 지체 3일 → 🔴 (기본 alert=3)', ncPendingLevel(DB.get('nc_records', n1.id)) === 2, '');
    DB.set('settings','quality', { pending_warn_days: 2, pending_alert_days: 7 });
    P('21 설정값 변경 반영 (alert 7 → 3일은 🟡)', ncPendingLevel(DB.get('nc_records', n1.id)) === 1 && qualityPolicy().alert === 7, '');
    DB.set('settings','quality', { pending_warn_days: 1, pending_alert_days: 3 });
    location.hash = '#dash'; _lockScanTs = 0; VIEWS.dash();
    P('21b 대시보드 「오늘 이상」에 판정 지체 집계 (alert 이상만)', /판정 지체 1/.test(T(document.querySelector('#view .dkpi.issue, #view .dkpi.ok').textContent)), T(document.querySelectorAll('#view .dkpi')[3].textContent).slice(0,100));
    location.hash = '#qc'; router();
    P('21c 품질 목록 배지 🔴 판정 지체 n일', /판정 지체 3일/.test(document.getElementById('view').textContent), '');
    // ── Audit ──
    const nc1 = DB.get('nc_records', n1.id);
    P('22/23 HOLD 설정·해제 이력 (hold_log on/off · by · reason) + _logs', nc1.hold_log.filter(h => h.on).length >= 2 && nc1.hold_log.filter(h => !h.on).length >= 2 && nc1.hold_log.every(h => h.by && h.ts) && (nc1._logs||[]).length >= 4, nc1.hold_log.length + ' hold_log · ' + (nc1._logs||[]).length + ' logs');
    P('24 판정 대기 중 출하 기록 (NC pending_out kind=ship · ref · qty)', (nc1.pending_out||[]).some(o => o.kind === 'ship' && /^SH-QA-/.test(o.ref_id) && o.qty === allowed - 5), JSON.stringify(nc1.pending_out||[]).slice(0,120));
    ncRecordPendingOut([nc1], 'os', 'OD-QA-1', 3);
    P('25 판정 대기 중 반출 기록 (kind=os)', (DB.get('nc_records', n1.id).pending_out||[]).some(o => o.kind === 'os' && o.ref_id === 'OD-QA-1'), '');
    P('25b 출하 원장 구조 무변 — shipments 에 pending_nc 필드 없음', DB.all('shipments').every(x => !('pending_nc' in x)) && !/pending_nc/.test(Array.from(document.querySelectorAll('script')).map(s=>s.textContent).join('')), '');
    VIEWS._qcDetail(n1.id);
    const tl = T(document.getElementById('modal').textContent);
    P('25c NC 상세 타임라인에 HOLD 설정/해제·판정 대기 중 출하 표시 + [로트 불량 HOLD] 버튼', /로트 불량 HOLD 설정/.test(tl) && /HOLD 해제 → 불량 수량만 격리/.test(tl) && /판정 대기 중 출하/.test(tl) && !!document.getElementById('nc-hold-on'), '');
    closeModal();
    // ── 기존 자산 승계 ──
    P('26 U1 승계 — HOLD 오류문에 NC 번호·해결 방법 (hold=true 조건)', (() => { ncSetHold(n1.id, true, 'U1'); const m = tryRun(() => assertLotNotHeld(PLOT, '실적 등록')) || ''; ncSetHold(n1.id, false, 'U1 복귀'); return m.includes(n1.id) && /판정\\(재작업\\/특채\\/폐기\\)/.test(m); })(), '');
    P('27 K1-3 승계 — 특채 처리 → HOLD 해제 (시드 NC)', (() => { const h = DB.all('nc_records').find(n => ncIsHold(n) && n.lot_id); if(!h) return false; const before = lotHoldInfo(h.lot_id).held; DB.set('nc_records', h.id, { status: '특채', handle:{ kind:'deviation', ts:Date.now(), by:'qa' } }); const after = lotHoldInfo(h.lot_id).held; DB.set('nc_records', h.id, { status:'접수', handle:null }); return before && !after; })(), '');
    P('28 v2.2 K2 승계 — HOLD LOT 실적 등록 차단 (registerRecord)', (() => { const h = DB.all('nc_records').find(n => ncIsHold(n) && n.lot_id); const e = tryRun(() => LOT.registerRecord({ plot_no: h.lot_id, proc: getRoute(DB.get('production_lots', h.lot_id).itemCode)[0].proc, equipment:'', good:1, type:'in', shift:'day' })); return typeof e === 'string' && /품질 HOLD/.test(e); })(), '');
    P('29 등록 창 = 「판정 범위」 선택칸(부분 불량 기본 · 로트 불량) — 체크박스 없음 · 두 경로 모두 기본 부분 불량', (() => { VIEWS._qcRegister({}); const sc = document.getElementById('nc-scope'); const a = sc && sc.value === 'partial' && !document.getElementById('nc-hold') && /수량만 격리/.test(document.getElementById('nc-hold-note').textContent) && document.getElementById('nc-hold-reason-wrap').style.display === 'none'; closeModal(); VIEWS._qcRegister({ record_id: rec.id, lot_id: PLOT, item_code: lot.itemCode, defect_qty_source: 40, already_qty: 0 }); const b = document.getElementById('nc-scope').value === 'partial'; closeModal(); return a && b; })(), '');
    P('29a 로트 불량 선택 → 사유 칸 표시(필수) · 사유 없이 등록 시 차단', (() => { VIEWS._qcRegister({}); const sc = document.getElementById('nc-scope'); sc.value = 'lot'; sc.onchange(); const shown = document.getElementById('nc-hold-reason-wrap').style.display !== 'none' && /LOT 전체 차단/.test(document.getElementById('nc-hold-note').textContent); document.getElementById('nc-lot').value = PLOT; document.getElementById('nc-qty').value = 2; const n0 = DB.all('nc_records').length; window.__t=[]; const _t=window.toast; window.toast=(m,k)=>window.__t.push(m); document.getElementById('nc-save').onclick(); window.toast=_t; const blocked = DB.all('nc_records').length === n0 && window.__t.some(m => /사유/.test(m)); closeModal(); return shown && blocked; })(), '');
    P('29b 등록 시 hold 미지정 → 판정 대기 (qcRegisterNc 기본값 false)', (() => { const n = qcRegisterNc({ lot_id: PLOT, item_code: lot.itemCode, defect_type: dt, qty: 1 }); const ok = ncIsPending(DB.get('nc_records', n.id)) && DB.get('nc_records', n.id).hold === false; voidDoc('nc_records', n.id, 'QA'); return ok; })(), '');
    P('29c 로트 불량(hold:true+사유) 등록 → 즉시 HOLD + hold_reason + hold_log', (() => { const n = qcRegisterNc({ lot_id: PLOT, item_code: lot.itemCode, defect_type: dt, qty: 1, hold: true, hold_reason: 'heat 혼입 의심' }); const d = DB.get('nc_records', n.id); const ok = ncIsHold(d) && d.hold_reason === 'heat 혼입 의심' && d.hold_log.length === 1 && d.hold_log[0].reason === 'heat 혼입 의심'; voidDoc('nc_records', n.id, 'QA'); return ok; })(), '');
    P('29d [📷 등록 + 사진 첨부] 버튼 + NC 상세 [📎 사진·파일] → 첨부 모듈 1벌(nc_records attachments)', (() => { VIEWS._qcRegister({}); const hasBtn = !!document.getElementById('nc-save-att'); closeModal(); VIEWS._attFilesModal('nc_records', n1.id); const modalOk = !!document.getElementById('dwf-link-add'); document.getElementById('dwf-link-name').value = '현장 사진'; document.getElementById('dwf-link-url').value = 'https://example.com/photo.jpg'; document.getElementById('dwf-link-add').onclick(); const att = _drawAttList(DB.get('nc_records', n1.id)); closeModal(); VIEWS._qcDetail(n1.id); const detailBtn = /사진·파일 1/.test(document.getElementById('nc-att').textContent); closeModal(); return hasBtn && modalOk && att.length === 1 && att[0].kind === 'link' && detailBtn; })(), '');
    P('29e NC 상세 정보에 판정 범위 행 (부분 불량 n ea 격리 / 로트 불량 + 사유)', (() => { VIEWS._qcDetail(n1.id); const t = T(document.getElementById('modal').textContent); closeModal(); return /판정 범위/.test(t) && /부분 불량 — 30ea 격리/.test(t); })(), '');
    P('30 원장 불변 — records/lots 총량 (NC·hold·pending_out 은 nc_records 만 변경)', DB.all('production_lots').length === 32 && DB.all('records').every(r => typeof r.good === 'number') && !DB.all('records').some(r => 'hold' in r || 'pending_out' in r), '');
    P('버전 v0.38.22', APP_VERSION === 'v0.38.22', APP_VERSION);
    // ── 드롭존 (사용자 요청 2026-09-10: 부적합 등록·도면 등록 폼에 드래그 업로드 · 저장 후 attUploadFiles 1벌) ──
    P('29f 부적합 등록 폼에 드롭존 + 도면 등록 폼에 드롭존 (attPendingDropZone 1벌)', (() => { VIEWS._qcRegister({}); const a = !!document.querySelector('#nc-drop .att-dz'); closeModal(); VIEWS._drawNewModal(); const b = !!document.querySelector('#dw-drop .att-dz'); closeModal(); return a && b && typeof attPendingDropZone === 'function' && typeof attUploadFiles === 'function'; })(), '');
    P('29g 폼에 파일 드롭 → 대기 목록 표시 → 등록 시 문서 생성 후 업로드 (로컬 모드 = 인라인 ≤400KB)', (() => { VIEWS._qcRegister({}); const dz = document.querySelector('#nc-drop .att-dz'); const f = new window.File([new Uint8Array(1200)], 'crack.jpg', { type:'image/jpeg' }); const dt = { files:[f], types:['Files'] }; const ev = new window.Event('drop', { bubbles:true, cancelable:true }); ev.dataTransfer = dt; dz.ondrop(ev); const pendingShown = /crack\.jpg/.test(document.querySelector('#nc-drop .att-dz-list').textContent); document.getElementById('nc-lot').value = PLOT; document.getElementById('nc-qty').value = 1; const n0 = DB.all('nc_records').length; document.getElementById('nc-save').onclick(); const nc = DB.all('nc_records').sort((a,b)=>b.ts-a.ts)[0]; window.__ncDrop = nc.id; return pendingShown && DB.all('nc_records').length === n0 + 1; })(), '');
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
  // 업로드는 비동기 → ★ v0.38.18 고정 600ms 대기 → 첨부 반영까지 폴링(최대 15초) — 병렬 실행·CPU 부하 시 간헐 FAIL(29h·29i) 원인 제거 (GPT Gate: 테스트 안정성)
  const __t0 = Date.now();
  const __waitAtt = () => { const d = DB.get('nc_records', window.__ncDrop); return (d && _drawAttList(d).length > 0) || Date.now() - __t0 > 15000; };
  const __poll = setInterval(() => { if(!__waitAtt()) return; clearInterval(__poll);
    try {
      const nc = DB.get('nc_records', window.__ncDrop); const att = _drawAttList(nc);
      P('29h 저장 후 첨부 반영 — kind=inline · data URL · 이름 유지', att.length === 1 && att[0].kind === 'inline' && att[0].name === 'crack.jpg' && (att[0].data||'').length > 100, JSON.stringify(att.map(a=>({k:a.kind,n:a.name,s:a.size}))));
      const big = new window.File([new Uint8Array(500*1024)], 'big.jpg', { type:'image/jpeg' });
      attUploadFiles('nc_records', nc.id, [big]).then(n => { P('29i 로컬 모드 400KB 초과 → 거부(건수 0) · 서버 모드 경로 무변', n === 0 && _drawAttList(DB.get('nc_records', nc.id)).length === 1, 'n=' + n); window.__done = true; });
    } catch(e){ window.__T.push({ n:'FATAL2', ok:false, note: e.stack || e.message }); window.__done = true; }
  }, 50);
})();`;
win.eval(scripts.join('\n;\n') + body);
const finish = () => {
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nHOLD3 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 47) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 47 (중도 종료·누락)`);
process.exit((fail || out.length !== 47) ? 1 : 0);
};
let waited = 0; const tick = () => { if(win.__done || waited > 250){ if(!win.__done) (win.__T = win.__T || []).push({ n:'TIMEOUT 비동기 첨부 검사 25초 초과', ok:false }); finish(); } else { waited++; setTimeout(tick, 100); } }; tick();
