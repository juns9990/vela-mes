// ★ v0.35.5 — 쓰기 경로 자체 탐지 잠금 회귀 (QV-1b O-2 실브라우저 FAIL #3: LOT 추적 미방문 상태에서 후공정 등록 통과)
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window;
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const body = `
;(function(){
  const R = window.__T = [];
  const P = (n, ok, note) => R.push({ n, ok, note: String(note||'') });
  const tryRun = f => { try { f(); return null; } catch(e){ return e.message; } };
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa';
    const PLOT = 'PL-2026-D002';
    const route = getRoute(DB.get('production_lots', PLOT).itemCode).map(s => s.proc);
    const i0 = route.findIndex(pr => LOT.balance(PLOT, pr) > 0);
    const p0 = route[i0], p1 = route[i0 + 1];
    // 실브라우저 재현: 탭A 150 정상 등록 + 오프라인 150 도착 (사고 기록 없음 · LOT 추적 미방문)
    LOT.registerRecord({ plot_no:PLOT, proc:p0, equipment:'프레스 1호', good:150, type:'in', shift:'day' });
    DB.set('records', 'R-OFFLINE', { id:'R-OFFLINE', plot_no:PLOT, proc:p0, equipment:'프레스 2호', type:'in', good:150, defect:0, rework:0, scrap:0, rework_returned:0, ts:Date.now(), by:'OP-099', worker:{name:'작업자B'} });
    const cnt0 = DB.all('records').length, ship0 = DB.all('shipments').length, os0 = DB.all('os_dispatches').length;
    // ★ v0.36.0 (Q6-0) 승계 — 잠금(LOT 전체 차단) → 보류(초과분만 · 인정분 진행) · 사고 문서는 Audit 의미로 유지
    P('준비: 단조 보류 100 · 잔량 0 · 사고 기록 0 (화면 미방문 상태)', LOT.pendingQty(PLOT, p0) === 100 && LOT.balance(PLOT, p0) === 0 && DB.all('concurrency_incidents').length === 0, 'pend ' + LOT.pendingQty(PLOT, p0) + ' · bal ' + LOT.balance(PLOT, p0) + ' · inc ' + DB.all('concurrency_incidents').length);
    const accP0 = LOT.acceptedGood(PLOT, p0);
    P('준비: 단조 인정양품 = 가용(부풀림 없음) · 보고양품 = 인정 + 100', accP0 === LOT.totalGood(PLOT, p0) - 100 && LOT.balance(PLOT, p1) === accP0 - LOT.totalInput(PLOT, p1), 'acc ' + accP0 + ' · tot ' + LOT.totalGood(PLOT, p0));
    // T1 ★ 후공정(열처리) 등록 → 인정분 안이면 허용 + 쓰기 경로에서 사고 기록 자동 생성 (잠금이었다면 차단되던 것)
    const e1 = tryRun(() => LOT.registerRecord({ plot_no:PLOT, proc:p1, equipment:'', good:10, type: DB.get('processes',p1)?.type==='out'?'out':'in', shift:'day' }));
    P('T1 후공정 등록 → 인정분 안 허용 (LOT 추적 미방문 · 잠금 아님)', e1 === null, e1 || '');
    P('T1 등록과 동시에 사고 기록 생성 (멱등 1건 · pending 100)', DB.all('concurrency_incidents').filter(c=>c.plot_no===PLOT && c.status==='open').length === 1 && DB.all('concurrency_incidents')[0].pending === 100 && (DB.all('concurrency_incidents')[0].conflict_records||[]).some(x => x.id === 'R-OFFLINE' && x.pending === 100 && x.accepted === 50), JSON.stringify(DB.all('concurrency_incidents')[0]?.conflict_records));
    P('T1 records +1 (인정분 흐름 유지)', DB.all('records').length === cnt0 + 1, cnt0 + ' → ' + DB.all('records').length);
    // T2 사고 공정(단조) 등록은 잔량 0 → 초과 등록 차단 (Lock 의 수량 차단 = balance 가 흡수)
    const e2 = tryRun(() => LOT.registerRecord({ plot_no:PLOT, proc:p0, equipment:'', good:10, type:'in', shift:'day' }));
    P('T2 사고 공정 등록 → 잔량 0 초과 등록 차단', typeof e2 === 'string' && e2.includes('초과 등록 차단') && e2.includes('잔량 0ea'), e2);
    // T3 출하·반출 write path 재검증 — 인정 가용 밖 차단(보류 안내) · 안이면 통과 + 사고 목록 반환
    const e3 = tryRun(() => assertWithinAccepted(PLOT, 1, '출하', 0)), e4 = tryRun(() => assertWithinAccepted(PLOT, 30, '외주 반출', 20));
    const ok5 = assertWithinAccepted(PLOT, 5, '출하', 10);
    P('T3 출하·반출 인정 밖 차단 (보류 안내 · 같은 사고번호)', typeof e3 === 'string' && e3.includes('보류 100') && typeof e4 === 'string' && e4.includes('외주 반출 가능 20') && ok5.incidents.length === 1 && DB.all('concurrency_incidents').length === 1, [e3, e4].join(' | '));
    P('T3 count 불변 (records·shipments·os_dispatches)', DB.all('records').length === cnt0 + 1 && DB.all('shipments').length === ship0 && DB.all('os_dispatches').length === os0, '');
    // T4 오탐 없음: 정상 LOT 은 사고 아님 · 사고 생성 안 됨
    const healthy = DB.all('production_lots').find(l => l.plot_no !== PLOT && !lotHasPending(l.plot_no));
    P('T4 정상 LOT 사고 아님 · 사고 미생성 (오탐 0)', healthy && lotConflictInfo(healthy.plot_no).open === false && DB.all('concurrency_incidents').length === 1, healthy && healthy.plot_no);
    // T5 lockedPlotSet(사고 LOT 집합) 도 화면 방문 없이 최신 (목록·선택지 표시용 — 표시는 ⚠ 보류)
    P('T5 lockedPlotSet 에 사고 LOT 포함', lockedPlotSet().has(PLOT), '');
    // T6 판정 → 정정 → J1′ 재검증 해제 → 등록 계속 (시스템은 승자를 고르지 않는다)
    const ci = DB.all('concurrency_incidents').find(c => c.plot_no === PLOT && c.status === 'open');
    const eRes0 = tryRun(() => resolveIncident(ci.id));
    P('T6 정정 전 해제 거부 (J1′ · 보류 미해소)', typeof eRes0 === 'string' && eRes0.includes('J1′ 재검증 실패') && eRes0.includes('보류 100'), eRes0);
    const eJ = tryRun(() => recordIncidentJudgment(ci.id, 'A', '작업자B 오입력 — 150 → 50'));
    P('T6 관리자 판정 Case A 기록 (원장 무변 · 사고 open 유지)', eJ === null && DB.get('concurrency_incidents', ci.id).resolution?.case === 'A' && DB.get('concurrency_incidents', ci.id).status === 'open' && DB.get('records','R-OFFLINE').good === 150, eJ || '');
    const eCorr = tryRun(() => LOT.correctRecord('R-OFFLINE', { good:50, defect:0, rework:0 }, '검증 정정'));
    P('T6 감소 보정(150→50) 허용 → 보류 0', eCorr === null && LOT.pendingQty(PLOT, p0) === 0 && LOT.balance(PLOT, p0) === 0, eCorr || ('pend ' + LOT.pendingQty(PLOT, p0)));
    const eRes1 = tryRun(() => resolveIncident(ci.id));
    P('T6 J1′ 재검증 통과 → 해제', eRes1 === null && lotConflictInfo(PLOT).open === false && DB.get('concurrency_incidents', ci.id).status === 'resolved', eRes1 || '');
    const e7 = tryRun(() => LOT.registerRecord({ plot_no:PLOT, proc:p1, equipment:'', good:10, type: DB.get('processes',p1)?.type==='out'?'out':'in', shift:'day' }));
    P('T6 해제 후 후공정 등록 계속', e7 === null && DB.all('records').length === cnt0 + 2, e7 || '');
    // T7 UI: PC 실적 등록 모달 — 사고(보류) 안내줄·옵션 라벨 배선 · 파비콘 내장
    P('T7 등록 모달 사고 안내 배선 (⚠ 보류 · 인정분 안에서 등록)', /⚠ 동시성 사고 — 초과 \\$\\{lk\\.pending/.test(document.documentElement.outerHTML), '');
    P('T7 등록 모달 LOT 옵션에 사고 라벨', /동시성 사고 \\(인정분 진행\\)/.test(document.documentElement.outerHTML), '');
    P('T7 파비콘 data URI 내장 (콘솔 404 제거)', /<link rel="icon" href="data:image\\/svg\\+xml/.test(document.documentElement.outerHTML), '');
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nLOCK-LIVE TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 18) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 18 (중도 종료·누락)`);
process.exit((fail || out.length !== 18) ? 1 : 0);
