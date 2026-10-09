// ★ v0.35.6 — 실적 ID 충돌 방지 + 파일럿 Gate UI 문구 회귀 (QV-1b 관찰 #1~#3·#5 · GPT 파일럿 전 필수)
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
    const p0 = route.find(pr => LOT.balance(PLOT, pr) > 0);
    // T1 ★ 같은 ms 에 같은 공정·LOT 실적 2건 → records +2 (v0.35.5 까지는 id 동일 → 병합 → +1)
    const realNow = Date.now; const FIX = realNow(); Date.now = () => FIX;
    const c0 = DB.all('records').length;
    const e1 = tryRun(() => { LOT.registerRecord({ plot_no:PLOT, proc:p0, equipment:'프레스 1호', good:10, type:'in', shift:'day' });
                              LOT.registerRecord({ plot_no:PLOT, proc:p0, equipment:'프레스 2호', good:10, type:'in', shift:'day' }); });
    Date.now = realNow;
    const added = DB.all('records').filter(r => r.plot_no===PLOT && r.proc===p0 && r.good===10);
    P('T1 같은 ms 2건 등록 → records +2 (무음 덮어쓰기 없음)', e1===null && DB.all('records').length === c0 + 2 && added.length === 2, e1 || (c0 + ' → ' + DB.all('records').length));
    P('T1 두 id 서로 다름 · 접두 R-{proc}-{plot}-{ts} 유지', added.length===2 && added[0].id !== added[1].id && added.every(r => r.id.startsWith('R-' + p0 + '-' + PLOT + '-' + FIX + '-')), added.map(r=>r.id).join(' | '));
    P('T1 잔량 = 두 건 모두 투입 반영', LOT.balance(PLOT, p0) === 200 - 20, 'bal ' + LOT.balance(PLOT, p0));
    // T2 newRecId 존재 시 재생성 (강제 충돌)
    const pre = 'R-X-' + PLOT; const rnd = Math.random; let calls = 0; Math.random = () => { calls++; return calls <= 1 ? 0.5 : rnd(); };
    const fixed = pre + '-' + Date.now() + '-' + (0.5).toString(36).slice(2,6);
    DB.set('records', fixed, { id:fixed, plot_no:PLOT, proc:'X', good:0 });
    const id2 = newRecId(pre); Math.random = rnd;
    P('T2 newRecId 는 이미 있는 id 를 피해 재생성', id2 !== fixed && !DB.get('records', id2), id2);
    // T3 MERGE/SPLIT/FORK/STOP 도 헬퍼 경유 (정적)
    const src = String(LOT.registerRecord) + String(LOT.merge||'') + String(LOT.split||'') + String(LOT.fork||'');
    P('T3 registerRecord 가 newRecId 사용 · Date.now 직접 조합 0', /newRecId\\(/.test(String(LOT.registerRecord)) && !/R-\\$\\{proc\\}-\\$\\{plot_no\\}-\\$\\{Date\\.now/.test(String(LOT.registerRecord)), '');
    P('T3 R-STOP 도 헬퍼 경유', /newRecId\\(\`R-STOP-/.test(document.documentElement.outerHTML) || /newRecId\\(\`R-STOP-/.test(String(VIEWS._perfStop||'')) , '');
    const allJs = Array.from(document.querySelectorAll('script')).map(s=>s.textContent).join('\\n');
    P('T3 Date.now() 직접 조합 실적 id 0건 (R-…-\${Date.now()})', !/\`R-[^\`]*\\$\\{Date\\.now\\(\\)\\}\`/.test(allJs), '');
    // T4 출하 창: 잠긴 LOT 선택 시 ⛔ 차단 안내만 · 완성품 경고 혼재 없음
    DB.set('records', 'R-OFFLINE-QA', { id:'R-OFFLINE-QA', plot_no:PLOT, proc:p0, equipment:'', type:'in', good:400, defect:0, rework:0, scrap:0, rework_returned:0, ts:Date.now(), by:'OP-099' });
    P('T4 준비: D002 사고(보류) 상태', lotConflictInfo(PLOT).open === true && lotConflictInfo(PLOT).pending > 0, '');
    location.hash = '#ship'; router();
    const e4 = tryRun(() => document.getElementById('btn-ship-new').onclick());
    const lotSel = document.getElementById('sh-lot');
    // v0.35.9: 출하 창 LOT 목록은 선택 수주 품번으로 좁혀짐 → 같은 품번 수주를 먼저 선택
    const _soSel = document.getElementById('sh-so'); const _item = DB.get('production_lots', PLOT).itemCode; const _so = [..._soSel.options].find(o => o.dataset.i === _item); if(_so){ _soSel.value = _so.value; _soSel.dispatchEvent(new window.Event('change')); }
    let t4 = false, t4b = false;
    if(lotSel){ lotSel.value = PLOT; lotSel.dispatchEvent(new window.Event('change')); const w = document.getElementById('sh-warn'); t4 = !!w && /⚠ 동시성 사고/.test(w.innerHTML) && /보류/.test(w.innerHTML) && !/완성품재고/.test(w.innerHTML); t4b = /동시성 사고 \\(인정분만 출하 가능\\)/.test(lotSel.innerHTML); }
    P('T4 출하 창 사고 LOT → ⚠ 보류 안내 (인정분 출하 가능 · 완성품 경고 없음) — v0.36.0 승계', t4, e4 || (lotSel ? '' : 'sh-lot 없음'));
    P('T4 출하 창 LOT 옵션 라벨 사고', t4b, '');
    closeModal();
    // T5 반출 창: 대상 표·LOT 옵션에 잠금 표시
    location.hash = '#os'; router();
    const e5 = tryRun(() => document.getElementById('btn-disp').onclick());
    const odl = document.getElementById('od-l'); const modalHtml = document.getElementById('modal') ? document.getElementById('modal').innerHTML : '';
    P('T5 반출 창 LOT 옵션 사고 라벨 (인정분만)', !!odl && /동시성 사고 \\(인정분만\\)/.test(odl.innerHTML), e5 || '');
    P('T5 반출 대상 표 ⚠ 사고 표시 (해당 LOT 후보가 있을 때)', !/반출 대상 실적/.test(modalHtml) || !new RegExp(PLOT).test(modalHtml) || /⚠ 사고/.test(modalHtml), '');
    closeModal();
    // T6 NC 상세: 오등록 취소 라벨 + 안내 줄 · 처리 창 존재
    const nc = DB.all('nc_records').find(n => n.status === '접수');
    if(nc){ VIEWS._qcDetail ? VIEWS._qcDetail(nc.id) : null; }
    const mh = document.getElementById('modal') ? document.getElementById('modal').innerHTML : '';
    P('T6 NC 상세 [오등록 취소] 라벨 + 판정→[처리] 안내', /오등록 취소/.test(mh) && /판정\\(특채·폐기·반품·재작업\\)/.test(mh), nc ? '' : '접수 NC 없음');
    P('T6 오등록 취소 prompt 취소 시 무동작 (코드)', /if\\(reason === null\\) return;/.test(allJs), '');
    P('버전 ≥ v0.35.6 · 2026-09-07', /^v0\\.35\\.([6-9]|[1-9][0-9])$|^v0\\.3[6-9]/.test(APP_VERSION) && APP_DATE >= '2026-09-07', APP_VERSION + ' ' + APP_DATE);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nRECID TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 15) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 15 (중도 종료·누락)`);
process.exit((fail || out.length !== 15) ? 1 : 0);
