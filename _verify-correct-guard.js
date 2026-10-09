// ★ v0.35.4 — 실적 보정 잔량 규율 회귀 (QV-1b 실브라우저 발견: 실적현황 [수정] 초과 보정이 저장되던 결함)
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
    const i0 = route.findIndex(pr => LOT.balance(PLOT, pr) > 0);   // 현재 진행 위치 (D002 = FORGE 잔량 200)
    const p0 = route[i0], p1 = route[i0 + 1];
    // 준비: 첫 공정 실적 100 (잔량 200 → 100), 둘째 공정 실적 60 (가용 = 첫 공정 누적 양품)
    const bal0 = LOT.balance(PLOT, p0);
    const r0 = LOT.registerRecord({ plot_no:PLOT, proc:p0, equipment:'E', good:100, type:'in', shift:'day' });
    const g1before = LOT.totalGood(PLOT, p0);
    const r1 = LOT.registerRecord({ plot_no:PLOT, proc:p1, equipment:'E', good:60, type:'in', shift:'day' });
    const bal0after = LOT.balance(PLOT, p0), bal1after = LOT.balance(PLOT, p1);
    P('준비: 등록 2건 성공 · 첫 공정 잔량 = 이전−100', bal0after === bal0 - 100 && bal1after >= 0, 'bal0 ' + bal0 + '→' + bal0after + ' · bal1 ' + bal1after);
    // T1 실브라우저 재현: 초과 보정 (100 → 잔량+100+1) 차단
    const over = bal0after + 100 + 1;
    const e1 = tryRun(() => LOT.correctRecord(r0.id, { good: over, defect:0, rework:0 }, 'qa'));
    P('T1 잔량 초과 보정 차단 (등록과 동일 규율)', typeof e1 === 'string' && e1.includes('보정 차단') && e1.includes('잔량 초과'), e1);
    P('T1 차단 후 원장 무변 (레코드 수량·잔량 그대로)', DB.get('records', r0.id).good === 100 && LOT.balance(PLOT, p0) === bal0after, '');
    // T2 한도 내 증가 보정 허용 (정확히 잔량만큼)
    const e2 = tryRun(() => LOT.correctRecord(r0.id, { good: 100 + bal0after, defect:0, rework:0 }, 'qa'));
    P('T2 한도 내 증가 보정 허용 (잔량 0 까지)', e2 === null && LOT.balance(PLOT, p0) === 0, e2 || ('bal ' + LOT.balance(PLOT, p0)));
    // T3 하류 음수 유발 보정 차단: 첫 공정 양품을 둘째 공정 소비(60) 미만으로 줄이면 둘째 공정 잔량 음수 → 차단
    const downLimit = g1before - 100 + 60;   // 둘째 공정 가용 = 첫 공정 누적 양품 → r0.good 이 (60 − 기존 양품) 미만이면 음수
    const tooLow = Math.max(0, (60 - (g1before - 100)) - 1);
    const e3 = (60 - (g1before - 100)) > 0
      ? tryRun(() => LOT.correctRecord(r0.id, { good: tooLow, defect:0, rework:0 }, 'qa'))
      : 'skip';
    P('T3 하류 공정 음수 유발 감소 보정 차단', e3 === 'skip' || (typeof e3 === 'string' && e3.includes('보정 차단')), e3 === 'skip' ? '기존 양품이 충분해 재현 불가 — 건너뜀' : e3);
    // T4 사고 LOT(이미 음수) 개선 보정 허용 · 악화 보정 차단
    //   주입: 둘째 공정에 가용을 넘는 실적 도착 (다른 단말 오프라인 등록 = Gate 시나리오)
    const availP1 = LOT.balance(PLOT, p1) + LOT.totalInput(PLOT, p1);
    DB.set('records', 'R-QA-OFFLINE', { id:'R-QA-OFFLINE', plot_no:PLOT, proc:p1, equipment:'E', type:'in', good: availP1, defect:0, rework:0, scrap:0, rework_returned:0, ts:Date.now(), by:'qa' });
    // ★ v0.36.0 (Q6-0) 승계: 잔량 음수 → 보류(pending) 양수 · 잔량은 0 (J1′)
    const negPend = LOT.pendingQty(PLOT, p1);
    P('T4 준비: 둘째 공정 보류 양수 (사고 상태 · 잔량 0)', negPend > 0 && LOT.balance(PLOT, p1) === 0, 'pend1 ' + negPend + ' · bal1 ' + LOT.balance(PLOT, p1));
    const e4a = tryRun(() => LOT.correctRecord('R-QA-OFFLINE', { good: availP1 + 10, defect:0, rework:0 }, 'qa'));
    P('T4-a 사고 LOT 악화 보정(증가) 차단', typeof e4a === 'string' && e4a.includes('보정 차단'), e4a);
    const e4b = tryRun(() => LOT.correctRecord('R-QA-OFFLINE', { good: Math.max(availP1 - 10, 0), defect:0, rework:0 }, 'qa'));
    P('T4-b 사고 LOT 개선 보정(감소 · 아직 보류) 허용 — Gate ③ 정정 경로 유지', e4b === null && LOT.pendingQty(PLOT, p1) === negPend - 10, e4b || ('pend1 ' + LOT.pendingQty(PLOT, p1)));
    const e4c = tryRun(() => LOT.correctRecord('R-QA-OFFLINE', { good: 0, defect:0, rework:0 }, 'qa'));
    P('T4-c 사고 LOT 완전 정정(0) 허용 → 보류 0', e4c === null && LOT.pendingQty(PLOT, p1) === 0 && LOT.balance(PLOT, p1) >= 0, e4c || ('pend1 ' + LOT.pendingQty(PLOT, p1)));
    // T5 드라이런 무부작용: correctionCheck 만 호출 → 레코드·잔량 불변 · 로그 미증가
    const logsBefore = (DB.get('records', r0.id)._logs || []).length, balB = LOT.balance(PLOT, p0);
    const chk = LOT.correctionCheck(r0.id, { good: 99999 });
    P('T5 correctionCheck 드라이런 — 판정만 · 원장/로그 불변', chk.ok === false && chk.violations.length >= 1 && (DB.get('records', r0.id)._logs || []).length === logsBefore && LOT.balance(PLOT, p0) === balB, JSON.stringify(chk.violations[0]));
    // T6 UI 경로 배선: 실적현황 [수정] 핸들러가 LOT.correctRecord 를 경유 (DB.correct 직접 호출 0)
    P('T6 _perfEdit → LOT.correctRecord 배선 (DB.correct 직접 호출 제거)', /LOT\\.correctRecord\\(recId/.test(String(VIEWS._perfEdit)) && !/DB\\.correct\\('records', recId/.test(String(VIEWS._perfEdit)), '');
    P('T6 보정 모달 잔량 힌트 표기', /잔량 초과 보정 차단/.test(String(VIEWS._perfEdit)), '');
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nCORRECT-GUARD TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 12) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 12 (중도 종료·누락)`);
process.exit((fail || out.length !== 12) ? 1 : 0);
