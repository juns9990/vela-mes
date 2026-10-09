// ★ v0.38.0 — 개인화 회귀 (SPEC 「개인화 v0.1」 v0.1.1 §6 · GPT 🟢): 대시보드 카드 12종 · 보고서 8 섹션 · Picker 1벌 · [보고서 만들기] 원클릭 · prefs 필드 · 라벨
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window;
// XLSX/PPT 스텁 — 시트·슬라이드 호출을 기록해 섹션 필터를 검증
win.__xl = []; win.__pp = [];
win.XLSX = { utils:{ book_new:()=>({}), aoa_to_sheet:(aoa)=>{ win.__xl.push(aoa); return { '!merges':[] }; }, encode_cell:()=>'A1', book_append_sheet:()=>{} }, writeFile:()=>{} };
win.PptxGenJS = function(){ const self = this; this.ShapeType = { rect:'rect' }; this.addSlide = () => { const sl = { titles:[], addShape(){}, addText(t){ if(typeof t === 'string') sl.titles.push(t); }, addTable(){} }; win.__pp.push(sl); return sl; }; this.writeFile = () => Promise.resolve(); };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.matchMedia = win.matchMedia || (() => ({ matches:false, addEventListener(){}, removeEventListener(){}, addListener(){}, removeListener(){} }));
win.open = () => ({ document:{ open(){}, write(){}, close(){} }, focus(){}, print(){}, close(){} });
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const body = `
;(async function(){
  const R = window.__T = [];
  const P = (n, ok, note) => R.push({ n, ok, note: String(note||'') });
  const T = s => (s||'').replace(/\\s+/g,' ');
  const view = () => document.getElementById('view');
  const modal = () => document.getElementById('modal');
  const shown = () => document.getElementById('modal-bg').classList.contains('show');
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa'; SESSION.userId = 'U-DEMO';
    renderNav();
    // ── A. 대시보드 카탈로그 · 기본 ──
    P('A1 DASH_CARDS 12종 ID 고정 · k_alert fixed · DASH_DEFAULT = v3 A안 6종', JSON.stringify(DASH_CARDS.map(c=>c.id)) === JSON.stringify(['k_backlog','k_sales','k_today','k_alert','c_fulfil','c_trend','c_capa','c_quality','c_due','c_stock','c_cust','c_prog']) && DASH_CARDS.find(c=>c.id==='k_alert').fixed === true && JSON.stringify(DASH_DEFAULT) === JSON.stringify(['k_backlog','k_sales','k_today','k_alert','c_fulfil','c_trend']), '');
    location.hash = '#dash'; router();
    P('A2 prefs 없음 → 회사 기본 = 스트립 4 + 카드 2 (dash357 무변) · 「카드 고르기」 버튼', view().querySelectorAll('.dkpi').length === 4 && view().querySelectorAll('.card').length === 2 && !!document.getElementById('dash-pick'), view().querySelectorAll('.card').length);
    // 카드 선택·순서
    setPrefs('U-DEMO', { dash_cards: ['k_today','k_alert','c_trend','c_stock','c_cust'] }); router();
    const kp = [...view().querySelectorAll('.dkpi .lab')].map(e => T(e.textContent).trim());
    P('A3 prefs.dash_cards 반영 — KPI 2(오늘 생산·오늘 이상 순서) · 카드 3(추세·재고·고객사) · 계산 무변(총 수주잔 값 동일)', kp.length === 2 && kp[0].includes('오늘 생산') && kp[1].includes('오늘 이상') && view().querySelectorAll('.card').length === 3 && /완제품 재고/.test(view().textContent) && /고객사별 매출/.test(view().textContent) && !/수주 대비 생산실적/.test(view().textContent), kp.join('|'));
    setPrefs('U-DEMO', { dash_cards: ['k_backlog','c_fulfil'] }); router();
    P('A4 k_alert 를 빼도 항상 표시(고정 · 안전 신호 은폐 금지) — KPI 뒤에 삽입', JSON.stringify(dashCardsFor()) === JSON.stringify(['k_backlog','k_alert','c_fulfil']) && [...view().querySelectorAll('.dkpi .lab')].some(e => /오늘 이상/.test(e.textContent)), dashCardsFor().join(','));
    // 권한·플래그
    DB.set('users','U-PV',{ id:'U-PV', name:'권한', pin:'x', role:'operator', active:true, perms:{ inv:'none', ship:'none' } });
    SESSION.userId = 'U-PV'; SESSION.role = 'operator';
    setPrefs('U-PV', { dash_cards: ['k_backlog','k_sales','k_alert','c_stock','c_cust','c_trend'] }); router();
    P('A5 권한 none(inv·ship) 카드 = 표시 0(k_sales·c_stock·c_cust 제외) · 카탈로그(Picker)에서도 0', JSON.stringify(dashCardsFor()) === JSON.stringify(['k_backlog','k_alert','c_trend']) && (()=>{ openDashPicker(); const ids = [...modal().querySelectorAll('[data-pk]')].map(r=>r.dataset.pk); closeModal(); return !ids.includes('c_stock') && !ids.includes('c_cust') && !ids.includes('k_sales') && ids.includes('c_due'); })(), dashCardsFor().join(','));
    SESSION.userId = 'U-DEMO'; SESSION.role = 'manager';
    const fl = DB.get('settings','features')?.flags; setFlag('capa', false); router();
    setPrefs('U-DEMO', { dash_cards: ['k_alert','c_capa','c_trend'] }); router();
    P('A6 플래그 OFF(capa) 카드 제외', JSON.stringify(dashCardsFor()) === JSON.stringify(['k_alert','c_trend']), dashCardsFor().join(','));
    setFlag('capa', true); router();
    P('A6 플래그 ON → 설비 부하 카드 표시 (computeCapa 읽기만)', dashCardsFor().includes('c_capa') && /설비 부하/.test(view().textContent), '');
    // 회사 기본 문서
    setPrefs('U-DEMO', { dash_cards: [] });
    DB.set('settings','dash', { default: ['k_backlog','k_alert','c_fulfil','c_prog'] }); router();
    P('A7 prefs [] + settings/dash.default → 회사 기본 문서 적용 (공정 진행 4단 표시)', JSON.stringify(dashCardsFor()) === JSON.stringify(['k_backlog','k_alert','c_fulfil','c_prog']) && /공정 진행/.test(view().textContent), dashCardsFor().join(','));
    DB.set('settings','dash', { default: [] }); router();
    P('A7 회사 기본 문서 비움 → 코드 기본 6종', JSON.stringify(dashCardsFor()) === JSON.stringify(DASH_DEFAULT), '');
    // ── B. Picker 1벌 ──
    openDashPicker();
    let m = modal();
    P('B1 Picker — 내 구성 6행(★) + 추가 가능 6행(☆) · k_alert 자물쇠(해제 0 · ▲▼ 0 · 🔒) · 나머지 5행 ▲▼', m.querySelectorAll('.pk-row.on').length === 6 && m.querySelectorAll('[data-pk-on]').length === 6 && !!m.querySelector('[data-pk="k_alert"] .pk-star.fixed') && !m.querySelector('[data-pk-off="k_alert"]') && !m.querySelector('[data-pk-up="k_alert"]') && !!m.querySelector('[data-pk="k_alert"] .pk-lock') && m.querySelectorAll('[data-pk-up]').length === 5, '');
    m.querySelector('[data-pk-dn="k_today"]').onclick();
    P('B1b 다른 KPI 를 k_alert 아래로 내려도 k_alert 는 항상 마지막 KPI 뒤 (위치 규칙 · 결정론)', JSON.stringify([...m.querySelectorAll('.pk-row.on')].map(r=>r.dataset.pk).slice(0,4)) === JSON.stringify(['k_backlog','k_sales','k_today','k_alert']), [...m.querySelectorAll('.pk-row.on')].map(r=>r.dataset.pk).join(','));
    P('B1c dashPinAlert — 저장 순서가 k_alert 앞이어도 렌더는 마지막 KPI 뒤 · KPI 없으면 맨 앞', JSON.stringify(dashPinAlert(['k_alert','k_sales','c_trend'])) === JSON.stringify(['k_sales','k_alert','c_trend']) && JSON.stringify(dashPinAlert(['c_trend','c_due'])) === JSON.stringify(['k_alert','c_trend','c_due']), '');
    m.querySelector('[data-pk-on="c_due"]').onclick();
    m.querySelector('[data-pk-off="c_fulfil"]').onclick();
    m.querySelector('[data-pk-up="c_due"]').onclick();
    const order = [...m.querySelectorAll('.pk-row.on')].map(r => r.dataset.pk);
    P('B2 ☆ 넣기 · ★ 빼기 · ▲ 순서 (창 안 상태만 · 저장 전 prefs 무변)', JSON.stringify(order) === JSON.stringify(['k_backlog','k_sales','k_today','k_alert','c_due','c_trend']) && JSON.stringify(getPrefs('U-DEMO').dash_cards) === '[]', order.join(','));
    document.getElementById('pk-save').onclick();
    P('B3 저장 → prefs.dash_cards = 창 순서 · 창 닫힘 · 대시보드 재렌더(납기 카드)', JSON.stringify(getPrefs('U-DEMO').dash_cards) === JSON.stringify(order) && !shown() && /납기/.test(view().textContent), '');
    openDashPicker(); document.getElementById('pk-reset').onclick();
    P('B4 「회사 기본으로」 → prefs.dash_cards = [] · 기본 6종', JSON.stringify(getPrefs('U-DEMO').dash_cards) === '[]' && JSON.stringify(dashCardsFor()) === JSON.stringify(DASH_DEFAULT), '');
    SESSION.userId = null; SESSION.role = 'vendor'; openDashPicker(); m = modal();
    P('B5 userId 없는 개발자 — 저장·회사 기본 버튼 비활성 + 안내', document.getElementById('pk-save').disabled && document.getElementById('pk-reset').disabled && m.innerHTML.includes('개발자 세션은 개인 구성을 저장할 수 없습니다'), '');
    closeModal(); SESSION.userId = 'U-DEMO'; SESSION.role = 'manager';
    // ── C. 보고서 섹션 ──
    P('C1 REPORT_SECTIONS 8 ID 고정 · kpi fixed · REPORT_FORMATS 4', JSON.stringify(REPORT_SECTIONS.map(x=>x.id)) === JSON.stringify(['kpi','prod','sales','pur','qc','fin','issues','ai']) && REPORT_SECTIONS[0].fixed === true && JSON.stringify(Object.keys(REPORT_FORMATS)) === JSON.stringify(['screen','xlsx','pptx','print']), '');
    P('C2 라벨 — 메뉴 「보고서」 · 제목 「주간보고서」(AI 접두 0) · 모듈 정의', PAGES.find(p=>p.id==='report').title === '보고서' && _REPORT_TITLES.week === '주간보고서' && !/AI 보고서/.test(JSON.stringify(_REPORT_TITLES)), PAGES.find(p=>p.id==='report').title);
    location.hash = '#report'; router();
    const rv = view();
    P('C3 prefs 없음 → 전부 8 섹션 · 「내 보고서 구성」 행 + [☆ 구성 바꾸기] + 기본 출력 select(screen) · [보고서 만들기] 버튼 · [보고서 생성] 0', rv.querySelectorAll('[data-rep-sec]').length === 7 && !!document.getElementById('btn-rep-pick') && document.getElementById('rep-fmt').value === 'screen' && !!document.getElementById('btn-rep-make') && !document.getElementById('btn-rep-gen') && JSON.stringify(myReportSections()) === JSON.stringify(REPORT_SECTIONS.map(x=>x.id)), rv.querySelectorAll('[data-rep-sec]').length);
    P('C4 키 없음 → 자동 브리핑 섹션은 숨기지 않고 「자동 브리핑 사용 안 함」 안내', !!rv.querySelector('[data-rep-sec="ai"]') && /자동 브리핑 사용 안 함/.test(rv.textContent) && !/AI 종합 브리핑/.test(rv.textContent), '');
    setPrefs('U-DEMO', { report_sections: ['kpi','qc','prod','issues'] }); router();
    const secsShown = [...view().querySelectorAll('[data-rep-sec]')].map(e => e.dataset.repSec);
    P('C5 prefs.report_sections 반영 — 화면 = 품질 → 생산부 → 이슈 (순서 그대로) · 영업·구매·경영·브리핑 0 · 전사 KPI 항상', JSON.stringify(secsShown) === JSON.stringify(['qc','prod','issues']) && /생산 실적/.test(view().textContent), secsShown.join(','));
    // 엑셀·PPT·인쇄 = 같은 배열
    window.__xl.length = 0; _reportExportXlsx(buildReportData('week'));
    const aoaTxt = JSON.stringify(window.__xl[0] || []);
    P('C6 엑셀 = 같은 섹션 (■ 품질·■ 생산부·■ 이슈 있음 · ■ 영업부·■ 경영·■ 자동 브리핑 없음)', /■ 품질/.test(aoaTxt) && /■ 생산부/.test(aoaTxt) && /■ 이슈/.test(aoaTxt) && !/■ 영업부/.test(aoaTxt) && !/■ 경영/.test(aoaTxt) && !/브리핑/.test(aoaTxt) && /■ 전사 요약/.test(aoaTxt), '');
    window.__pp.length = 0; buildReportPpt(buildReportData('week'));
    const ppTitles = window.__pp.map(s => s.titles[0] || '');
    P('C7 PPT = 같은 섹션·같은 순서 — 표지·전사·품질·생산부·이슈 5장 (영업·구매·경영·브리핑 0)', window.__pp.length === 5 && /품질/.test(ppTitles[2]) && /생산부/.test(ppTitles[3]) && /이슈/.test(ppTitles[4]) && !ppTitles.some(t=>/영업부|경영|브리핑/.test(t)), ppTitles.join('|'));
    P('C7 엑셀 순서 = prefs 순서 (■ 품질 이 ■ 생산부 보다 앞)', aoaTxt.indexOf('■ 품질') < aoaTxt.indexOf('■ 생산부'), '');
    const ai = _aiBriefingInput(buildReportData('week'), myReportSections());
    P('C8 자동 브리핑 입력 = 고른 섹션만([기간][전사] 항상 · [생산][품질][이슈] · [영업][구매][경영] 0) · 형식 줄도 고른 부서만 · 라벨 「선택한 섹션 기준」', /\\[기간\\]/.test(ai.text) && /\\[전사\\]/.test(ai.text) && /\\[생산\\]/.test(ai.text) && /\\[품질\\]/.test(ai.text) && !/\\[영업\\]/.test(ai.text) && !/\\[구매/.test(ai.text) && !/\\[경영/.test(ai.text) && /\\[생산부\\]/.test(ai.format) && !/\\[영업부\\]/.test(ai.format) && /선택한 섹션 기준/.test(ai.label), ai.label);
    P('C8 전부 선택이면 라벨 「전체 섹션 기준」', reportSectionLabel(REPORT_SECTIONS.map(x=>x.id)) === '전체 섹션 기준', '');
    P('C9 buildReportData 무변 — 섹션 선택과 무관하게 8키 전부 계산', (()=>{ const d = buildReportData('week'); return ['companyKPI','production','sales','procurement','quality','finance','issues'].every(k => d[k] !== undefined); })(), '');
    // 권한 none → 섹션 제외
    DB.set('users','U-RV',{ id:'U-RV', name:'보고권한', pin:'x', role:'operator', active:true, perms:{ closing:'none', acct:'none', qc:'edit' } });
    SESSION.userId = 'U-RV'; SESSION.role = 'operator';
    P('C10 매핑 메뉴 전부 none(closing·acct) → 경영·정산 섹션 제외 · 하나라도 view 면 유지', !myReportSections().includes('fin') && myReportSections().includes('qc') && (()=>{ openReportPicker(); const ids = [...modal().querySelectorAll('[data-pk]')].map(r=>r.dataset.pk); closeModal(); return !ids.includes('fin') && ids.includes('prod'); })(), myReportSections().join(','));
    SESSION.userId = 'U-DEMO'; SESSION.role = 'manager';
    // ── D. 원클릭 · 스냅샷 sections · 하위 호환 ──
    location.hash = '#report'; router();
    const n0 = DB.all('reports').length;
    location.hash = '#dash'; router(); location.hash = '#report'; router(); document.querySelector('[data-rep-tab="month"]').onclick(); document.querySelector('[data-rep-tab="week"]').onclick();
    P('D1 🔴 페이지 열기·탭 전환 시 reports 자동 생성 0', DB.all('reports').length === n0, DB.all('reports').length - n0);
    document.getElementById('rep-fmt').value = 'pptx'; document.getElementById('rep-fmt').onchange();
    P('D2 기본 출력 select → prefs.report_format = pptx', getPrefs('U-DEMO').report_format === 'pptx' && myReportFormat() === 'pptx', '');
    window.__pp.length = 0;
    const st = window._reportState; st.viewSnapshotId = null; st.scope = 'week';
    const rid = await makeReportOneClick(st);
    const snap = DB.get('reports', rid);
    P('D3 [보고서 만들기] 1회 = reports +1 · sections = 내 구성 저장 · data 는 전부(영업·경영 포함) · 열람 모드로 전환', DB.all('reports').length === n0 + 1 && JSON.stringify(snap.sections) === JSON.stringify(['kpi','qc','prod','issues']) && snap.data.sales && snap.data.finance && st.viewSnapshotId === rid, JSON.stringify(snap.sections));
    P('D4 키 없음 → 브리핑 건너뜀(ai_summary 없음 · 오류 0) · 기본 출력 PPT 호출 1회(스냅샷 섹션 5장)', snap.data.ai_summary == null && window.__pp.length === 5, window.__pp.length);
    router();
    P('D5 스냅샷 열람 = 저장 당시 sections (prefs 바꿔도 무변)', (()=>{ setPrefs('U-DEMO', { report_sections: ['kpi','sales'] }); router(); const ids = [...view().querySelectorAll('[data-rep-sec]')].map(e=>e.dataset.repSec); return JSON.stringify(ids) === JSON.stringify(['qc','prod','issues']) && /이 스냅샷의 구성/.test(view().textContent); })(), '');
    DB.set('reports','RP-OLD-TEST', { id:'RP-OLD-TEST', scope:'week', period: snap.period, data: snap.data, ts: Date.now()-1000, by:'qa' });
    st.viewSnapshotId = 'RP-OLD-TEST'; router();
    P('D6 sections 없는 구 스냅샷 → 전부 표시 (하위 호환)', view().querySelectorAll('[data-rep-sec]').length === 7, view().querySelectorAll('[data-rep-sec]').length);
    st.viewSnapshotId = null; setPrefs('U-DEMO', { report_sections: [] }); router();
    P('D7 prefs [] → 전부 (회사 기본 문서 없음)', view().querySelectorAll('[data-rep-sec]').length === 7, '');
    DB.set('settings','report', { default: ['kpi','prod','ai'] }); router();
    P('D7 settings/report.default 적용', JSON.stringify(myReportSections()) === JSON.stringify(['kpi','prod','ai']), myReportSections().join(','));
    DB.set('settings','report', { default: [] }); router();
    // 스냅샷 원장 무변 · 라이브 수치 = 스냅샷 수치
    P('D8 원클릭 스냅샷 data = 라이브 buildReportData 와 동일 수치(companyKPI)', JSON.stringify(snap.data.companyKPI) === JSON.stringify(buildReportData('week', st.anchor).companyKPI), '');
    // ── E. prefs 계약 (v0.37.7 승계) ──
    P('E1 PREFS_FIELDS 5 (v0.38.17 sec_off 추가) · report_format 화이트리스트 · 남의 prefs 거부', JSON.stringify(PREFS_FIELDS) === JSON.stringify(['menu_primary','dash_cards','report_sections','report_format','sec_off']) && setPrefs('U-PV', { report_format:'xlsx' }) === null && !!setPrefs('U-DEMO', { report_format:'xlsx' }) && getPrefs('U-DEMO').report_format === 'xlsx', '');
    P('E2 폰 현장 화면 무변 — .m-mode 에서 dash-tools 숨김 CSS · #m/ 라우트 무변', /#app\\.m-mode \\.dash-tools\\{display:none\\}/.test(document.documentElement.outerHTML) && typeof mobileRouter === 'function', '');
    P('C11 kpi 는 항상 첫째 (_normReportSections)', JSON.stringify(_normReportSections(['qc','kpi','prod'])) === JSON.stringify(['kpi','qc','prod']), '');
    P('버전 v0.38.21', APP_VERSION === 'v0.38.21', APP_VERSION);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
  window.__DONE = true;
})();`;
win.eval(scripts.join('\n;\n') + body);
(function wait(){
  if(!win.__DONE){ setTimeout(wait, 50); return; }
  const out = win.__T || [];
  let fail = 0;
  out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
  console.log(`\nPERSONAL TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
  // ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
  if(out.length !== 41) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 41 (중도 종료·누락)`);
  process.exit((fail || out.length !== 41) ? 1 : 0);
})();
