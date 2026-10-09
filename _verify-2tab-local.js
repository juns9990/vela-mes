// ★ v0.35.3 — 로컬 어댑터 문서 단위 쓰기·탭 동기화 회귀 (QV-1b O절 실브라우저 FAIL → 수정 검증)
// 탭 = 독립 JSDOM 창. 실제 브라우저의 localStorage 공유·storage 이벤트는 스냅샷 전달 + StorageEvent 수동 dispatch 로 재현.
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const KEY = 'vela_mes_demo_v1';
function boot(preLS){
  const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
  const win = dom.window;
  win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
  win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
  win.sessionStorage.setItem('vela_demo','1');
  if(preLS) win.localStorage.setItem(KEY, preLS);
  win.eval(scripts.join('\n;\n') + `;(function(){ if(!DB.all('production_lots').length) seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='qa';
    window.__X = (code) => eval(code); })();`);
  return win;
}
const tick = () => new Promise(r => setTimeout(r, 0));
// 다른 탭의 쓰기를 이 탭에 전달 (브라우저가 하는 일을 수동으로)
function pushStorage(toWin, newValue){
  const ev = new toWin.StorageEvent('storage', { key: KEY, newValue, storageArea: toWin.localStorage });
  toWin.dispatchEvent(ev);
}
const R = []; const P = (n, ok, note) => R.push({ n, ok, note: String(note||'') });
const PLOT = 'PL-2026-D002', PROC = 'FORGE';
const REG = (w, q) => w.__X(`(function(){ try { LOT.registerRecord({plot_no:'${PLOT}',proc:'${PROC}',equipment:'프레스 1호',good:${q},type:'in',shift:'day',source:'manual_pc'}); return null; } catch(e){ return e.message; } })()`);

(async () => {
  try {
    // ── S0 시드 정합: 마이크로태스크 합류 후 localStorage = 메모리 전량
    const A = boot(null);
    await tick();
    const lsA0 = JSON.parse(A.localStorage.getItem(KEY));
    const recMem = A.__X(`DB.all('records').length`), recLS = Object.values(lsA0.records||{}).filter(x=>!x.void).length;
    P('S0 시드 후 localStorage 문서 수 = 메모리 (records)', recMem === recLS && recMem === 84, `mem ${recMem} · ls ${recLS}`);
    const memCols = A.__X(`(function(){ const s=DB.snapshot(); return Object.keys(s).filter(c=>Object.keys(s[c]).length).length; })()`);
    P('S0 시드 후 localStorage 컬렉션 수 = 메모리(문서 1개 이상)', Object.keys(lsA0).filter(c=>Object.keys(lsA0[c]).length).length === memCols, `ls ${Object.keys(lsA0).length} · mem ${memCols}`);
    const S0 = A.localStorage.getItem(KEY);

    // ── S1 예방 (같은 기기 탭 2개 · storage 이벤트 동기화): 탭B 는 잔량 50 을 보고 150 을 차단해야 한다
    const B = boot(S0);
    P('S1 탭B 부팅 잔량 200', B.__X(`LOT.balance('${PLOT}','${PROC}')`) === 200, '');
    P('S1 탭A 150 등록 성공', REG(A, 150) === null, '');
    await tick();
    const SA = A.localStorage.getItem(KEY);
    P('S1 탭A 쓰기 후 localStorage records 85', Object.keys(JSON.parse(SA).records).length === 85, '');
    pushStorage(B, SA);
    P('S1 storage 이벤트 수신 → 탭B 메모리 병합 (records 85 · 잔량 50)', B.__X(`DB.all('records').length`) === 85 && B.__X(`LOT.balance('${PLOT}','${PROC}')`) === 50, `rec ${B.__X(`DB.all('records').length`)} bal ${B.__X(`LOT.balance('${PLOT}','${PROC}')`)}`);
    const errB = REG(B, 150);
    P('S1 탭B 150 등록 → 초과 등록 차단 (예방)', typeof errB === 'string' && errB.includes('초과 등록 차단') && errB.includes('잔량 50'), errB);
    B.localStorage.setItem(KEY, SA);            // 브라우저의 공유 localStorage 현재값 (A 가 쓴 것)
    P('S1 탭B 50 등록은 허용', REG(B, 50) === null, '');
    await tick();
    const SB = B.localStorage.getItem(KEY);
    P('S1 두 탭 실적 모두 보존 (records 86)', Object.keys(JSON.parse(SB).records).length === 86, '');
    pushStorage(A, SB);
    P('S1 탭A 역방향 병합 → 잔량 0', A.__X(`LOT.balance('${PLOT}','${PROC}')`) === 0 && A.__X(`DB.all('records').length`) === 86, '');
    P('S1 정합: 사고 0 · 잠금 없음', A.__X(`(detectConcurrencyIncidents()||[]).length`) === 0 && A.__X(`lockedPlotSet().has('${PLOT}')`) === false, '');

    // ── S2 진짜 동시 도착 (동기화 전에 둘 다 씀 = 다른 단말·오프라인): 두 문서 모두 원장에 남고 Gate 가 잠근다
    const A2 = boot(S0), B2 = boot(S0);
    P('S2 탭A 150 등록', REG(A2, 150) === null, '');
    await tick();
    const SA2 = A2.localStorage.getItem(KEY);
    // 탭B 는 아직 storage 이벤트를 못 받은 상태에서 150 등록 (자기 메모리 잔량 200) → 자기 문서만 RMW 로 추가
    B2.localStorage.setItem(KEY, SA2);            // 브라우저의 공유 localStorage 현재값 (A 가 쓴 것)
    P('S2 탭B(stale) 150 등록 성공 (자기 메모리 기준 잔량 200)', REG(B2, 150) === null, '');
    await tick();
    const SB2 = B2.localStorage.getItem(KEY);
    const n2 = Object.keys(JSON.parse(SB2).records).length;
    P('S2 ★ 유실 없음 — localStorage records 86 (구 방식 85)', n2 === 86, `records ${n2}`);
    const C = boot(SB2);
    const balC = C.__X(`LOT.pendingQty('${PLOT}','${PROC}')`), incC = C.__X(`(detectConcurrencyIncidents()||[]).length + DB.query('concurrency_incidents', c => c.plot_no === '${PLOT}' && c.status === 'open').length`), lockC = C.__X(`lockedPlotSet().has('${PLOT}')`);   // v0.35.7: 부팅 렌더(대시보드 오늘 이상)가 먼저 탐지하므로 incC = 신규+기존 합
    P('S2 새 부팅 진실: 보류 100(잔량 0) · 사고 ≥1 · 사고 LOT 집합 true (★ v0.36.0 Q6-0 승계)', balC === 100 && C.__X(`LOT.balance('${PLOT}','${PROC}')`) === 0 && incC >= 1 && lockC === true, `pend ${balC} inc ${incC} lock ${lockC}`);
    pushStorage(A2, SB2);
    P('S2 탭A 도 storage 병합으로 보류 100 인지 (새로고침 없이)', A2.__X(`LOT.pendingQty('${PLOT}','${PROC}')`) === 100, '');
    A2.localStorage.setItem(KEY, SB2);
    A2.__X(`detectConcurrencyIncidents()`);   // = 탭A 가 LOT 추적 진입 (O-1) — 렌더 시 자동 탐지
    P('S2 탭A LOT 추적 진입 → 사고 생성·사고 LOT 집합 true', A2.__X(`lockedPlotSet().has('${PLOT}')`) === true && A2.__X(`DB.all('concurrency_incidents').length`) === 1, '');
    const errLocked = REG(A2, 10);
    P('S2 사고 공정 신규 실적 → 잔량 0 초과 등록 차단 (Lock 의 수량 차단을 balance 가 흡수)', typeof errLocked === 'string' && errLocked.includes('초과 등록 차단') && errLocked.includes('잔량 0ea'), errLocked);
    P('S2 차단 시도 후 records 불변 (O-3)', A2.__X(`DB.all('records').length`) === 86, '');

    // ── S3 restore 전체 쓰기 경로 + remove 문서 단위 삭제
    const D = boot(S0);
    const snap = D.__X(`DB.snapshot()`);
    D.__X(`DB.set('orders','SO-TMP-1',{id:'SO-TMP-1',cust:'C001',itemCode:'HF-3301',qty:1,price:1,due:'2026-12-31',created:Date.now()})`);
    await tick();
    P('S3 set 후 localStorage 에 SO-TMP-1', !!JSON.parse(D.localStorage.getItem(KEY)).orders['SO-TMP-1'], '');
    D.__X(`DB.restore(${JSON.stringify(snap)})`);
    await tick();
    P('S3 restore(전체 쓰기) 후 SO-TMP-1 사라짐 · records 84', !JSON.parse(D.localStorage.getItem(KEY)).orders['SO-TMP-1'] && Object.keys(JSON.parse(D.localStorage.getItem(KEY)).records).length === 84, '');
    D.__X(`DB.set('items','ZZ-TMP',{id:'ZZ-TMP',name:'tmp'})`); await tick();
    const okRm = D.__X(`DB.remove('items','ZZ-TMP')`); await tick();
    P('S3 remove 문서 단위 삭제 반영 (타 문서 무손상)', okRm === true && !JSON.parse(D.localStorage.getItem(KEY)).items['ZZ-TMP'] && Object.keys(JSON.parse(D.localStorage.getItem(KEY)).items).length === D.__X(`DB.allRaw('items').length`), '');
    P('S3 flush() 동기 강제 반영', (()=>{ D.__X(`DB.set('orders','SO-TMP-2',{id:'SO-TMP-2',cust:'C001',itemCode:'HF-3301',qty:1,price:1,due:'2026-12-31',created:Date.now()}); DB._adapter.flush();`); return !!JSON.parse(D.localStorage.getItem(KEY)).orders['SO-TMP-2']; })(), '');
    // ── S4 자기 쓰기 에코 무시 · 모달 열림 시 리렌더 보류 (데이터는 병합)
    const E = boot(S0);
    E.__X(`window.__rc = 0; const _r = router; window.router = function(){ window.__rc++; };`);
    pushStorage(E, E.localStorage.getItem(KEY));
    P('S4 자기 쓰기 에코(동일 문자열) 무시 — 리렌더 0', E.__X(`window.__rc`) === 0, '');
    E.__X(`document.getElementById('modal-bg').classList.add('show')`);
    pushStorage(E, SB2);
    P('S4 모달 열림 중 수신 → 데이터 병합(86) · 리렌더 보류(0)', E.__X(`DB.all('records').length`) === 86 && E.__X(`window.__rc`) === 0, '');
    E.__X(`document.getElementById('modal-bg').classList.remove('show')`);
    pushStorage(E, SB);
    P('S4 모달 닫힌 뒤 수신 → 리렌더 1', E.__X(`window.__rc`) === 1, '');
  } catch(e){ R.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
  let fail = 0;
  R.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
  console.log(`\n2TAB TOTAL ${R.length} · PASS ${R.length-fail} · FAIL ${fail}`);
  // ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
  if(R.length !== 26) console.log(`[FAIL] 항목 수 ${R.length} ≠ 기대 26 (중도 종료·누락)`);
  process.exit((fail || R.length !== 26) ? 1 : 0);
})();
