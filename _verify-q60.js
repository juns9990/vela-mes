// ★ v0.36.0 — Q6-0 「초과분 보류/격리」 회귀 (설계서 v1.1 §10 · GPT FINAL PASS 2026-09-10 · 17항목 + 원장 무변)
//   원칙 검증: 인정/보류 = 파생(미저장) · ts ASC→id ASC · clamp 상한 · 양품 우선 차감 · 하류 연쇄 · Case A/B/C · Lock(b) 유지 · 강제 인정 경로 0
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window;
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.prompt = () => '검증 사유';
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const body = `
;(function(){
  const R = window.__T = [];
  const P = (n, ok, note) => R.push({ n, ok, note: String(note||'') });
  const tryRun = f => { try { f(); return null; } catch(e){ return e.message; } };
  const rq = r => (+r.good||0)+(+r.defect||0)+(+r.rework||0)+(+r.scrap||0)+(+r.transferred||0);
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa';
    const snap0 = { rec: DB.all('records').length, lots: DB.all('production_lots').length, ship: DB.all('shipments').length, os: DB.all('os_dispatches').length,
      recSum: DB.all('records').reduce((s,r)=>s+rq(r),0) };
    const heat = DB.all('material_lots')[0]?.heat || 'H2604-101';
    const PL = 'PL-Q60-A';   // 라우트 HF-3301: CUT → FORGE → HEAT(out) → SHOT → SHIP
    LOT.create({ plot_no:PL, itemCode:'HF-3301', materialLots:[heat], qty_initial:200 });
    LOT.registerRecord({ plot_no:PL, proc:'CUT', equipment:'절단기 1호', good:200, type:'in', shift:'day' });
    const T0 = Date.now();
    // 1 정상
    LOT.registerRecord({ plot_no:PL, proc:'FORGE', equipment:'프레스 1호', good:150, type:'in', shift:'day', worker:{name:'A'} });
    const s1 = LOT.acceptedSplit(PL,'FORGE');
    P('1 정상: 절단 200 → 단조 150 → 잔량 50 · 보류 0 · NORMAL', LOT.balance(PL,'FORGE') === 50 && LOT.pendingQty(PL,'FORGE') === 0 && s1.length === 1 && s1[0].status === 'NORMAL' && s1[0].accepted === 150, JSON.stringify(s1.map(x=>[x.accepted,x.pending,x.status])));
    // 2 충돌 주입 — B 단말 150 도착 (ts 늦음)
    DB.set('records', 'R-Q60-B', { id:'R-Q60-B', plot_no:PL, proc:'FORGE', equipment:'프레스 2호', type:'in', good:150, defect:0, rework:0, scrap:0, rework_returned:0, ts:T0+5000, by:'B', worker:{name:'B'} });
    const s2 = LOT.acceptedSplit(PL,'FORGE');
    const rowA = s2.find(x=>x.rec.worker?.name==='A'), rowB = s2.find(x=>x.rec.id==='R-Q60-B');
    P('2 충돌: 단조 A150+B150 / 실물 200 → 인정 200(A150·B50) · 보류 100 · B=PARTIAL_PENDING · 잔량 0 (음수 없음)', LOT.acceptedGood(PL,'FORGE') === 200 && LOT.pendingQty(PL,'FORGE') === 100 && rowA.accepted === 150 && rowB.accepted === 50 && rowB.pending === 100 && rowB.status === 'PARTIAL_PENDING' && LOT.balance(PL,'FORGE') === 0, JSON.stringify(s2.map(x=>[x.rec.id,x.accepted,x.pending])));
    // 3 열처리 가용 = 200 (300 아님) · 201 차단
    const e3 = tryRun(() => LOT.registerRecord({ plot_no:PL, proc:'HEAT', equipment:'외주', vendor:'V-HT', good:201, type:'out', shift:'day' }));
    P('3 열처리 가용 = 인정양품 200 (보고 300 아님) · 201 등록 → 초과 등록 차단', LOT.balance(PL,'HEAT') === 200 && typeof e3 === 'string' && /초과 등록 차단/.test(e3) && /잔량 200ea/.test(e3), e3);
    // 4 열처리 200 등록 성공 (Lock 이었다면 차단)
    const e4 = tryRun(() => LOT.registerRecord({ plot_no:PL, proc:'HEAT', equipment:'외주', vendor:'V-HT', good:200, type:'out', shift:'day' }));
    P('4 열처리 200 등록 성공 — 인정분 흐름 유지 (LOT 전체 잠금 아님)', e4 === null && LOT.balance(PL,'HEAT') === 0 && LOT.balance(PL,'SHOT') === 200, e4 || '');
    // 6 사고 문서 자동 생성(멱등) — 쓰기 경로에서 이미 생성됨
    const cis = DB.query('concurrency_incidents', c => c.plot_no === PL && c.status === 'open');
    detectConcurrencyIncidents(PL); detectConcurrencyIncidents();
    const cis2 = DB.query('concurrency_incidents', c => c.plot_no === PL && c.status === 'open');
    P('6 사고 문서 자동 생성 (쓰기 경로 · FORGE 1건 · pending 100 · 충돌 실적 = R-Q60-B 인정 50/보류 100) · 재탐지 멱등', cis.length === 1 && cis[0].proc === 'FORGE' && cis[0].pending === 100 && cis[0].conflict_records.length === 1 && cis[0].conflict_records[0].id === 'R-Q60-B' && cis[0].conflict_records[0].accepted === 50 && cis2.length === 1, JSON.stringify(cis.map(c=>[c.id,c.pending,c.conflict_records])));
    const CI = cis[0].id;
    // 5 출하: 인정분 범위 가능 + pending_log · 초과는 assertWithinAccepted 차단 (UI 값 무시)
    LOT.registerRecord({ plot_no:PL, proc:'SHOT', equipment:'쇼트 1호', good:200, type:'in', shift:'day' });
    const e5a = tryRun(() => assertWithinAccepted(PL, 201, '출하', LOT.acceptedGood(PL,'SHOT')));
    const ok5 = assertWithinAccepted(PL, 200, '출하', LOT.acceptedGood(PL,'SHOT'));
    ciRecordPendingOut(ok5.incidents, 'ship', 'SH-Q60-1', 200);
    P('5 출하: 인정 안(200) 통과 + 사고 pending_log 기록 · 201 은 보류 안내로 차단', typeof e5a === 'string' && /출하 가능 200/.test(e5a) && /보류 100/.test(e5a) && ok5.allowed === 200 && ok5.incidents.length === 1 && DB.get('concurrency_incidents', CI).pending_log.length === 1 && DB.get('concurrency_incidents', CI).pending_log[0].ref_id === 'SH-Q60-1', e5a);
    // 10 연쇄: 별도 LOT — 열처리가 부풀린 양을 먼저 소비한 경우 → 열처리 보류 자동
    const PL2 = 'PL-Q60-C';
    LOT.create({ plot_no:PL2, itemCode:'HF-3301', materialLots:[heat], qty_initial:200 });
    LOT.registerRecord({ plot_no:PL2, proc:'CUT', equipment:'절단기 1호', good:200, type:'in', shift:'day' });
    LOT.registerRecord({ plot_no:PL2, proc:'FORGE', equipment:'프레스 1호', good:150, type:'in', shift:'day', worker:{name:'A'} });
    DB.set('records', 'R-Q60-C-B', { id:'R-Q60-C-B', plot_no:PL2, proc:'FORGE', equipment:'프레스 2호', type:'in', good:150, defect:0, rework:0, scrap:0, rework_returned:0, ts:T0+5000, by:'B' });
    DB.set('records', 'R-Q60-C-H', { id:'R-Q60-C-H', plot_no:PL2, proc:'HEAT', equipment:'외주', vendor:'V-HT', type:'out', good:250, defect:0, rework:0, scrap:0, rework_returned:0, ts:T0+9000, by:'H' });   // 잠금 전 부풀린 300 을 보고 250 등록(다른 단말)
    P('10 연쇄 보류: 단조 보류 100 → 열처리 가용 200 → 열처리 250 중 인정 200 · 보류 50 자동 (재계산 절차 없음)', LOT.pendingQty(PL2,'FORGE') === 100 && LOT.acceptedGood(PL2,'FORGE') === 200 && LOT.pendingQty(PL2,'HEAT') === 50 && LOT.acceptedGood(PL2,'HEAT') === 200 && LOT.balance(PL2,'SHOT') === 200, 'F ' + LOT.pendingQty(PL2,'FORGE') + ' H ' + LOT.pendingQty(PL2,'HEAT'));
    // 7 Case A: A 150→50 보정 → 단조 보류 0 · 열처리 보류 50 유지(열처리 자체 초과) → 단조 사고 resolve 가능 · 열처리 사고는 open
    detectConcurrencyIncidents(PL2);
    const ciF = DB.query('concurrency_incidents', c => c.plot_no === PL2 && c.proc === 'FORGE' && c.status === 'open')[0];
    const ciH = DB.query('concurrency_incidents', c => c.plot_no === PL2 && c.proc === 'HEAT' && c.status === 'open')[0];
    const recA2 = DB.query('records', r => r.plot_no === PL2 && r.proc === 'FORGE' && r.worker?.name === 'A')[0];
    const eJ = tryRun(() => recordIncidentJudgment(ciF.id, 'A', 'A 오입력 150→50'));
    const eResBefore = tryRun(() => resolveIncident(ciF.id));
    const eC = tryRun(() => LOT.correctRecord(recA2.id, { good:50, defect:0, rework:0 }, 'Case A'));
    const eRes = tryRun(() => resolveIncident(ciF.id));
    P('7 Case A: 판정 기록(open 유지) → 정정 전 해제 거부 → correctRecord 150→50 → 단조 보류 0 · 열처리 보류 50 유지 → 단조 사고 resolve · 열처리 사고 open', eJ === null && DB.get('concurrency_incidents', ciF.id).resolution.case === 'A' && typeof eResBefore === 'string' && /J1′/.test(eResBefore) && eC === null && LOT.pendingQty(PL2,'FORGE') === 0 && LOT.pendingQty(PL2,'HEAT') === 50 && eRes === null && DB.get('concurrency_incidents', ciF.id).status === 'resolved' && !!ciH && DB.get('concurrency_incidents', ciH.id).status === 'open', [eJ, eResBefore, eC, eRes].join(' | '));
    // 8 Case B: 열처리 250 은 중복(잘못) → void → 열처리 보류 0 → resolve
    const eJB = tryRun(() => recordIncidentJudgment(ciH.id, 'B', '열처리 중복 등록'));
    voidDoc('records', 'R-Q60-C-H', 'Case B 중복');
    const eResH = tryRun(() => resolveIncident(ciH.id));
    P('8 Case B: void → 열처리 보류 0 → resolve 가능', eJB === null && LOT.pendingQty(PL2,'HEAT') === 0 && eResH === null && DB.get('concurrency_incidents', ciH.id).status === 'resolved', [eJB, eResH].join(' | '));
    // 9 Case C: 원 LOT 단조 사고 — 보류 유지 · 지체 감지 → 대시보드 「오늘 이상」
    const eJC = tryRun(() => recordIncidentJudgment(CI, 'C', '실물 확인 필요'));
    DB.set('concurrency_incidents', CI, { ts: Date.now() - 4*86400000 });   // 4일 경과 (조치 3일)
    const ciC = DB.get('concurrency_incidents', CI);
    P('9 Case C: 사유 기록 · status open 유지 · 판정 지체 4일 = 조치 단계(2) · 정책값 settings/concurrency 기본 1/3', eJC === null && ciC.status === 'open' && ciC.resolution.case === 'C' && ciPendingLevel(ciC) === 2 && concurrencyPolicy().warn === 1 && concurrencyPolicy().alert === 3, JSON.stringify(ciC.resolution));
    // 13 순서: 동일 ts → id ASC · 배열 순서 무관
    const PL3 = 'PL-Q60-T';
    LOT.create({ plot_no:PL3, itemCode:'HF-3301', materialLots:[heat], qty_initial:100 });
    LOT.registerRecord({ plot_no:PL3, proc:'CUT', equipment:'절단기 1호', good:100, type:'in', shift:'day' });
    const tsSame = T0 + 777;
    DB.set('records', 'R-Q60-T-Z', { id:'R-Q60-T-Z', plot_no:PL3, proc:'FORGE', type:'in', good:80, defect:0, rework:0, scrap:0, rework_returned:0, ts:tsSame, by:'Z' });
    DB.set('records', 'R-Q60-T-A', { id:'R-Q60-T-A', plot_no:PL3, proc:'FORGE', type:'in', good:80, defect:0, rework:0, scrap:0, rework_returned:0, ts:tsSame, by:'A' });
    const s13 = LOT.acceptedSplit(PL3,'FORGE');
    P('13 순서: 동일 ts → id ASC (R-…-A 전량 인정 · R-…-Z 20 인정/60 보류) · 도착(삽입) 순서 무관', s13[0].rec.id === 'R-Q60-T-A' && s13[0].accepted === 80 && s13[1].rec.id === 'R-Q60-T-Z' && s13[1].accepted === 20 && s13[1].pending === 60, JSON.stringify(s13.map(x=>[x.rec.id,x.accepted,x.pending])));
    // 14 결정론: 같은 원장 → 100회 동일 · 메모 무효화 후 동일 · 렌더 전후 동일
    const ref = JSON.stringify(LOT.acceptedSplit(PL3,'FORGE').map(x=>[x.rec.id,x.accepted,x.pending]));
    let same = true; for(let i=0;i<100;i++){ if(JSON.stringify(LOT.acceptedSplit(PL3,'FORGE').map(x=>[x.rec.id,x.accepted,x.pending])) !== ref) same = false; }
    LOT._invalidate(); const afterInv = JSON.stringify(LOT.acceptedSplit(PL3,'FORGE').map(x=>[x.rec.id,x.accepted,x.pending]));
    location.hash = '#lot'; router(); const afterRender = JSON.stringify(LOT.acceptedSplit(PL3,'FORGE').map(x=>[x.rec.id,x.accepted,x.pending]));
    P('14 결정론: 100회 호출 동일 · 메모 무효화 후 동일 · 렌더 전후 동일', same && afterInv === ref && afterRender === ref, ref);
    // 15 부분 인정 상한: 전 LOT 전수
    let bound = true, viol = '';
    DB.all('production_lots').forEach(l => getRoute(l.itemCode).forEach(r => { const sp = LOT.acceptedSplit(l.plot_no, r.proc); const pa = LOT._prevAvail(l.plot_no, r.proc); const sum = sp.reduce((s,x)=>s+x.accepted,0); sp.forEach(x => { if(!(x.accepted >= 0 && x.accepted <= x.reported && x.accepted + x.pending === x.reported)){ bound = false; viol = l.plot_no+'/'+r.proc+'/'+x.rec.id; } }); if(!pa.skip && sum > pa.avail){ bound = false; viol = l.plot_no+'/'+r.proc+' sum'; } if(LOT.balance(l.plot_no, r.proc) < 0){ bound = false; viol = 'neg '+l.plot_no; } }));
    P('15 상한 전수: 0 ≤ accepted_i ≤ reported_i ∧ accepted+pending = reported ∧ Σ accepted ≤ avail ∧ balance ≥ 0 (시드+검증 LOT 전체)', bound, viol);
    // 16 Lock(b) 유지: 인정 0 공정 등록 차단 + 사고 문서 open + lotConflictInfo.incidents + 대시보드 오늘 이상
    const e16 = tryRun(() => LOT.registerRecord({ plot_no:PL, proc:'FORGE', equipment:'프레스 1호', good:1, type:'in', shift:'day' }));
    const info16 = lotConflictInfo(PL);
    location.hash = '#dash'; router();
    const dashTxt = document.getElementById('view').textContent.replace(/\\s+/g,' ');
    P('16 Lock(b) 유지: 사고 공정 등록 → 잔량 0 차단 · 사고 open · lotConflictInfo.incidents 1 · pending 100 · 대시보드 「동시성 사고 n」', typeof e16 === 'string' && /잔량 0ea/.test(e16) && info16.open && info16.incidents.length === 1 && info16.pending === 100 && lockedPlotSet().has(PL) && /동시성 사고 \\d+/.test(dashTxt), e16);
    // 17 저장 필드 0: records 전수에 accepted*/pending* 키 없음 · 소스 정적
    const keys = new Set(); DB.allRaw('records').forEach(r => Object.keys(r).forEach(k => { if(/accepted|pending/i.test(k)) keys.add(k); }));
    const src = Array.from(document.querySelectorAll('script')).map(s=>s.textContent).join('');
    P('17 인정/보류 저장 필드 0 (records 키 검사 + 소스: DB.set(\\'records\\' 에 accepted 없음)', keys.size === 0 && !/DB\\.set\\('records'[^\\n]*accepted/.test(src), [...keys].join(','));
    // 11 강제 인정 경로 없음 · 원장 총량: records 총량은 검증 등록분만큼만 증가(자동 수정 0) · lots/shipments/os 는 검증 생성분만
    const added = DB.all('records').filter(r => /Q60/.test(r.plot_no));
    const recSumNow = DB.all('records').reduce((s,r)=>s+rq(r),0);
    const addedSum = added.reduce((s,r)=>s+rq(r),0);
    P('11 강제 인정 버튼/경로 0 (소스 grep) · 원장 무변: 기존 records 총량 불변 (증가분 = 검증 LOT 실적뿐) · shipments/os 불변', !/강제 인정|forceAccept|force_accept/.test(src) && recSumNow - addedSum === snap0.recSum && DB.all('shipments').length === snap0.ship && DB.all('os_dispatches').length === snap0.os && DB.all('production_lots').length === snap0.lots + 3, 'Δ ' + (recSumNow - addedSum - snap0.recSum));
    // 12 표준 시뮬·기존 하네스 승계는 별도 파일(_verify-sim-standard·2tab·lock-live·correct-guard) — 여기서는 J1′ 함수 자체
    const v = verifyLotInvariant(PL2);
    P('12 J1′ verifyLotInvariant: PL-Q60-C 전 공정 err 0 · pending 0 → ok', v.ok && v.checks.every(c => c.err === 0 && c.pending === 0), JSON.stringify(v.checks.map(c=>[c.proc,c.accepted,c.pending,c.bal])));
    P('버전 v0.38.23', APP_VERSION === 'v0.38.23', APP_VERSION);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nQ60 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 18) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 18 (중도 종료·누락)`);
process.exit((fail || out.length !== 18) ? 1 : 0);
