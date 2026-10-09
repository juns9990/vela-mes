// ★ v0.35.7 — 대시보드 v3(A안) + 메뉴 축소(매일/관리) 회귀 (QV-1b D′7 FAIL 1/3 수정 · 담당자 "복잡·정신없다" 대응)
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
  const T = s => s.replace(/\\s+/g,' ');
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa';
    // ── 대시보드 ──
    location.hash = '#dash'; VIEWS.dash();
    const v = document.getElementById('view');
    const tiles = [...v.querySelectorAll('.dkpi')];
    P('D1 스트립 정확히 4칸 · 라벨 = 총 수주잔/당월 매출/오늘 생산/오늘 이상', tiles.length === 4 && ['총 수주잔','당월 매출','오늘 생산','오늘 이상'].every((l,i)=>tiles[i].querySelector('.lab').textContent.includes(l)), tiles.map(t=>t.querySelector('.lab').textContent.trim()).join('|'));
    const st = calcSoFulfillment();
    const t0 = T(tiles[0].textContent);
    P('D2 총 수주잔 = calcSoFulfillment().total (수주 화면 띠와 같은 값 · 오차 0)', t0.includes('₩' + Math.round(st.total).toLocaleString()), t0.slice(0,80));
    P('D2 띠 카드 총 수주 금액 = 스트립 총 수주잔 (한 화면 두 숫자 일치)', T(v.querySelector('.stage-legend').textContent).includes('₩' + st.total.toLocaleString()), '');
    const sales = calcSales(DB.all('shipments'));
    P('D3 당월 매출 = calcSales (출하 화면과 계산 1벌)', sales.monthSales > 0 ? T(tiles[1].textContent).includes('₩' + Math.round(sales.monthSales).toLocaleString()) : T(tiles[1].textContent).includes('수집중'), '');
    const dayGood = DB.all('records').filter(r => r.ts >= new Date(new Date().setHours(0,0,0,0)).getTime()).reduce((s,r)=>s+(r.good||0),0);
    P('D4 오늘 생산 = 오늘 records 양품 합', T(tiles[2].querySelector('.val').textContent).startsWith(dayGood.toLocaleString()), tiles[2].querySelector('.val').textContent);
    // 오늘 이상 = 잠금 + HOLD + 부적합 대기 + 납기 지연
    // v0.35.10(9-B) 승계: 부적합 대기(당월 접수) → 판정 지체(alert 일수 이상 판정 대기 NC)
    const expIssue = lockedPlotSet().size + heldPlotSet().size + DB.query('nc_records', n => ncIsPending(n) && ncPendingLevel(n) >= 2).length + calcSoRows().filter(r=>r.status.label==='지연').length;
    const t3 = T(tiles[3].textContent);
    P('D5 오늘 이상 = 잠금+HOLD+판정 지체(alert↑)+납기 지연 합 (9-B 승계)', expIssue ? t3.includes(expIssue + '건') : t3.includes('이상 없음'), 'exp ' + expIssue + ' · ' + t3.slice(0,90));
    P('D5 이상 있으면 amber 타일(.issue) / 없으면 .ok', expIssue ? tiles[3].classList.contains('issue') : tiles[3].classList.contains('ok'), '');
    // 잠금 1건 추가 → 이상 +1 · 클릭 목적지 = 변경 이력(사고 카드)
    const PLOT = 'PL-2026-D002';
    const p0 = getRoute(DB.get('production_lots', PLOT).itemCode).map(s=>s.proc).find(pr => LOT.balance(PLOT, pr) > 0);
    DB.set('records', 'R-OFFLINE-QA', { id:'R-OFFLINE-QA', plot_no:PLOT, proc:p0, equipment:'', type:'in', good:999, defect:0, rework:0, scrap:0, rework_returned:0, ts:Date.now(), by:'OP-099' });
    _lockScanTs = 0;   // lockedPlotSet 400ms 캐시 무효화 (실사용은 다음 렌더에서 반영)
    VIEWS.dash();
    const tiles2 = [...document.querySelectorAll('#view .dkpi')];
    P('D6 사고 발생 → 오늘 이상 +1 · 「동시성 사고 1」 표기 · 클릭 → #audit (v0.36.0 승계)', T(tiles2[3].textContent).includes((expIssue+1) + '건') && tiles2[3].textContent.includes('동시성 사고 1') && (tiles2[3].getAttribute('onclick')||'').includes('#audit'), T(tiles2[3].textContent).slice(0,100));
    P('D7 카드 2장만 (수주 대비 띠 + 생산 추세) · 이동 항목 미노출', document.querySelectorAll('#view .card').length === 2 && !document.getElementById('view').innerHTML.includes('kgrid') && !document.getElementById('view').innerHTML.includes('cust5y'), '');
    P('D8 5초 규격 — 화면 내 통화/수량 대형 숫자(.val) 4개뿐', document.querySelectorAll('#view .dkpi .val').length === 4, '');
    // 추세 토글 동작
    const e1 = tryRun(() => document.querySelector('#trend-seg [data-r="m"]').onclick(new window.Event('click')));
    P('D9 추세 주/월/년 토글 동작', e1 === null && document.querySelector('#trend-seg [data-r="m"]').classList.contains('on'), e1 || '');
    // ── 메뉴 축소 ──
    renderNav();
    const nav = document.getElementById('nav');
    const top = [...nav.children].filter(el => el.matches('button[data-page]')).map(b => b.dataset.page);
    P('N1 매일 메뉴 = 대시보드 + 7 (수주·생산계획·실적 등록·실적현황·LOT·품질·출하) — 메뉴 순서 유지', JSON.stringify(top) === JSON.stringify(['dash','so','plan','scan','perf','lot','qc','ship']), top.join(','));
    // ★ v0.37.7 승계 — 「관리」 접이식 폐지: 나머지는 사이드바에 없고 「전체 메뉴」 창에 그룹별로 있음 (기능 숨김 0)
    P('N2 사이드바 = 내 메뉴만 + 「전체 메뉴」 버튼 (관리 접이식·그룹 라벨·아이콘 0)', !nav.querySelector('.nav-more') && !nav.querySelector('.nav-grp') && !nav.querySelector('button[data-page] svg') && !!nav.querySelector('[data-allmenu]'), nav.querySelectorAll('button[data-page]').length + '개');
    nav.querySelector('[data-allmenu]').onclick();
    const am = document.getElementById('modal');
    P('N2 전체 메뉴 창 = 나머지 전부 · 그룹 라벨 ≥4', am.querySelectorAll('[data-am]').length >= 18 && am.querySelectorAll('.am-grp-lbl').length >= 4, am.querySelectorAll('[data-am]').length + '개');
    P('N3 기능 수 불변 — 전체 메뉴 창 항목 수 = canSee 결과와 동일 (숨김 0)', am.querySelectorAll('[data-am]').length === getMenuOrder().map(id=>PAGES.find(p=>p.id===id)).filter(p=>p && getFlags()[p.flag] && (!MANAGER_ONLY.includes(p.id) || true)).length, am.querySelectorAll('[data-am]').length);
    P('N4 전체 메뉴 창 열림', document.getElementById('modal-bg').classList.contains('show'), '');
    closeModal();
    P('N4 닫기 → 닫힘 · vela_nav_more 미사용', !document.getElementById('modal-bg').classList.contains('show') && localStorage.getItem('vela_nav_more') === null, '');
    location.hash = '#masters'; router();
    P('N5 내 메뉴 밖 화면으로 이동 → 「전체 메뉴」 버튼 활성 + 화면 이름 힌트', nav.querySelector('[data-allmenu]').classList.contains('on') && (nav.querySelector('[data-allmenu] .nav-cur')||{}).textContent.includes('기준정보'), (nav.querySelector('[data-allmenu]')||{}).textContent);
    location.hash = '#dash'; router();
    P('N5 내 메뉴 화면으로 복귀 → 힌트 해제', !nav.querySelector('[data-allmenu]').classList.contains('on') && !nav.querySelector('[data-allmenu] .nav-cur'), '');
    // 설정 데이터로 매일/관리 조정 (업체별 — 하드코딩 아님)
    DB.set('settings','menu', { primary: ['dash','so','os'] }); renderNav();
    const top2 = [...nav.children].filter(el => el.matches('button[data-page]')).map(b => b.dataset.page);
    P('N6 settings/menu.primary 로 매일 메뉴 조정 가능 (업체 데이터)', JSON.stringify(top2) === JSON.stringify(['dash','so','os']), top2.join(','));
    DB.set('settings','menu', { primary: NAV_PRIMARY_DEFAULT.slice() }); renderNav();
    // 권한: 작업자 역할이면 MANAGER_ONLY 는 여전히 안 보임
    SESSION.role = 'worker'; renderNav();
    P('N7 권한 로직 무변 — worker 에게 MANAGER_ONLY 메뉴 미노출', MANAGER_ONLY.every(id => !nav.querySelector('button[data-page="' + id + '"]')), '');
    SESSION.role = 'manager'; renderNav();
    P('버전 ≥ v0.35.7', /^v0\.35\.([7-9]|[1-9][0-9])$|^v0\.3[6-9]/.test(APP_VERSION) && APP_DATE >= '2026-09-07', APP_VERSION);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nDASH357 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 22) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 22 (중도 종료·누락)`);
process.exit((fail || out.length !== 22) ? 1 : 0);
