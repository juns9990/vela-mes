// ★ v0.35.8 — 입력 창 심플 규칙 회귀 (파일럿 담당자 피드백: 새 수주 · 계획 수립 · 새 견적 원가 미리보기) + 담당자별 내 메뉴
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
  const M = () => document.getElementById('modal');
  const visibleInputs = () => [...M().querySelectorAll('input,select')].filter(e => { let p = e; while(p && p !== M()){ if(p.tagName === 'DETAILS' && !p.open) return false; if(p.style && p.style.display === 'none') return false; p = p.parentElement; } return e.type !== 'hidden'; });
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa';
    // ── 새 수주 ──
    location.hash = '#so'; router();
    document.getElementById('btn-so-new').onclick();
    const vis = visibleInputs().map(e => e.id);
    P('S1 새 수주 첫 화면 입력칸 = 고객사·품번·수량·납기 4개 (수주번호·단가는 더 보기)', JSON.stringify(vis) === JSON.stringify(['so-cust','so-item','so-qty','so-due']), vis.join(','));
    const custSel = document.getElementById('so-cust'), itemSel = document.getElementById('so-item');
    const custItems = DB.all('items').filter(i => i.cust === custSel.value && i.active !== false);
    P('S2 품번 목록 = 선택 고객사의 품번만 (짧은 라벨 · 품명/단가 미포함)', custItems.length > 0 && itemSel.options.length === custItems.length && ![...itemSel.options].some(o => /₩/.test(o.textContent)), itemSel.options.length + ' vs ' + custItems.length);
    const it = DB.get('items', itemSel.value);
    const auto = T(document.getElementById('so-auto').textContent);
    P('S3 품번 고르면 품명·단가·최근 수주 자동 카드', auto.includes(it.name) && auto.includes('₩' + (+it.price).toLocaleString()) && auto.includes('최근 수주'), auto.slice(0,100));
    P('S3 단가 칸(더 보기) = 마스터 단가 자동', +document.getElementById('so-price').value === +it.price, '');
    // 고객사 바꾸면 품번 목록 바뀜
    const other = [...custSel.options].map(o=>o.value).find(c => c !== custSel.value && DB.all('items').some(i => i.cust === c));
    if(other){ custSel.value = other; custSel.onchange(); }
    P('S4 고객사 변경 → 품번 목록 갱신 + 자동 카드 갱신', !other || ([...itemSel.options].every(o => DB.get('items', o.value).cust === other) && T(document.getElementById('so-auto').textContent).includes(DB.get('items', itemSel.value).name)), other || '고객사 1개');
    // 저장: 수량·납기 필수
    window.__toasts = []; const _t = window.toast; window.toast = (m,k) => { window.__toasts.push({m,k}); };
    const n0 = DB.all('orders').length;
    document.getElementById('so-qty').value = ''; document.getElementById('so-ok').onclick();
    P('S5 수량 없으면 저장 안 됨', DB.all('orders').length === n0 && window.__toasts.some(t => /수량/.test(t.m)), '');
    document.getElementById('so-qty').value = '120'; document.getElementById('so-due').value = ''; document.getElementById('so-ok').onclick();
    P('S5 납기 없으면 저장 안 됨 (담당자 요구: 납기만 고르면 됨 = 납기는 필수)', DB.all('orders').length === n0 && window.__toasts.some(t => /납기/.test(t.m)), '');
    document.getElementById('so-due').value = '2026-10-15'; document.getElementById('so-ok').onclick();
    const created = DB.all('orders').sort((a,b)=>(b.created||0)-(a.created||0))[0];
    P('S6 고객사·품번·수량·납기만으로 생성 → 단가는 마스터값 자동 저장', DB.all('orders').length === n0 + 1 && created.qty === 120 && created.due === '2026-10-15' && created.price === +DB.get('items', created.itemCode).price && created.cust === DB.get('items', created.itemCode).cust, JSON.stringify({q:created.qty, p:created.price, d:created.due}));
    window.toast = _t;
    // ── 계획 수립 ──
    location.hash = '#plan'; router();
    VIEWS._planNew();
    const tbl = document.getElementById('pln-so-body');
    const rows = [...tbl.querySelectorAll('tr[data-so]')];
    const hidden = document.getElementById('pln-so');
    P('P1 수주 연결 = 표 (재고생산 행 + 잔량 있는 수주 행) · 긴 드롭다운은 숨김', rows.length === hidden.options.length && hidden.style.display === 'none' && rows.length >= 2 && !document.getElementById('pln-so-cust'), rows.length + ' rows');
    const r1 = rows.find(r => r.dataset.so);
    r1.onclick();
    const opt = hidden.selectedOptions[0];
    P('P2 행 클릭 → 숨은 select 선택 + 품번·수량(계획 가능) 프리필 + 강조', hidden.value === r1.dataset.so && document.getElementById('pln-item').value === opt.dataset.item && +document.getElementById('pln-qty').value === +opt.dataset.planremain && /color-mix/.test(tbl.querySelector('tr[data-so="' + r1.dataset.so + '"]').getAttribute('style')||''), hidden.value);
    P('P2 표 컬럼 = 수주번호·고객사·품번·납기·계획 가능 (5개)', document.querySelectorAll('#pln-so-tbl thead th').length === 5, '');
    const q = document.getElementById('pln-so-q'); const target = hidden.options[1].value;
    q.value = target; q.oninput();
    const rowsQ = [...tbl.querySelectorAll('tr[data-so]')].filter(r => r.dataset.so);
    P('P3 검색 1줄로 수주 좁히기 (수주번호)', rowsQ.length === 1 && rowsQ[0].dataset.so === target, rowsQ.length);
    q.value = ''; q.oninput();
    // 프리필 경로 (수주 화면 [계획] 버튼) — 선택 행 강조
    closeModal(); VIEWS._planNew({ so_id: target });
    P('P4 프리필(so_id) → 표에서 그 행 강조 + 수량 프리필', document.getElementById('pln-so').value === target && /color-mix/.test(document.querySelector('#pln-so-body tr[data-so="' + target + '"]').getAttribute('style')||''), '');
    closeModal();
    // ── 새 견적 원가 미리보기 ──
    location.hash = '#quote'; router();
    document.getElementById('btn-q-new').onclick();
    const pv = document.getElementById('q-preview');
    const pvT = T(pv.textContent);
    const ratesReal = !!(DB.get('settings','cost_rates') && Object.keys(DB.get('settings','cost_rates').material_rates||{}).length);
    P('Q1 근거 줄 — 단가표 등록 여부를 문장으로 (실데이터 / 미등록)', ratesReal ? pvT.includes('● 실데이터') && pvT.includes('원가 › 단가표') : pvT.includes('단가표 미등록'), pvT.slice(0,90));
    const boxes = [...pv.querySelectorAll('div[style*="grid-template-columns:repeat(4,1fr)"] > div')];
    P('Q2 원가 4항목(재료·외주·가공·간접) 칸 + 비중 %', boxes.length === 4 && ['재료','외주','가공','간접'].every((l,i)=>boxes[i].textContent.includes(l)) && /%/.test(boxes[0].textContent), boxes.length);
    const c = computeCost(document.getElementById('q-item').value);
    const mp = +document.getElementById('q-margin').value, qty = +document.getElementById('q-qty').value;
    const unit = Math.round(c.total * (1 + mp/100)), amt = Math.round(c.total * (1 + mp/100) * qty);
    P('Q3 원가 → 마진 → 단가 → 금액 한 줄 · 숫자 = computeCost 1벌', pvT.includes('₩' + Math.round(c.total).toLocaleString()) && pvT.includes('마진 ' + mp + '%') && pvT.includes('₩' + unit.toLocaleString()) && pvT.includes('₩' + amt.toLocaleString()), '');
    // 단가표 없는 상태 → 미등록 문구
    const saved = DB.get('settings','cost_rates'); DB.set('settings','cost_rates', { material_rates:{}, process_rates:{}, outsource_rates:{}, overhead_pct:0.1 });
    document.getElementById('q-item').oninput();
    P('Q4 단가표 비면 「단가표 미등록 (예시 아님 · 0으로 계산)」 + 단가표 링크', T(document.getElementById('q-preview').textContent).includes('단가표 미등록') && !!document.querySelector('#q-preview a[href="#cost"]'), '');
    DB.set('settings','cost_rates', saved); document.getElementById('q-item').oninput();
    P('Q4 복원 → 실데이터 문구', T(document.getElementById('q-preview').textContent).includes('● 실데이터'), '');
    closeModal();
    // ── 담당자별 내 메뉴 ──
    DB.set('users','U-QA', { id:'U-QA', name:'검사', pin:'x', role:'worker', active:true });
    SESSION.userId = 'U-QA'; SESSION.role = 'worker';
    renderNav();
    let top = [...document.getElementById('nav').children].filter(el => el.matches('button[data-page]')).map(b => b.dataset.page);
    P('M1 내 메뉴 미설정 → 회사 기본 (dash + 7)', top.length === 8 && top[0] === 'dash', top.join(','));
    // ★ v0.37.7 승계 — [내 메뉴 편집] → 사이드바 「전체 메뉴」 버튼 · 체크박스 → ★ 토글 즉시 저장 · users.menu_primary → settings/prefs.U-QA
    P('M1 사이드바에 「전체 메뉴」 버튼', !!document.querySelector('#nav [data-allmenu]'), '');
    document.querySelector('#nav [data-allmenu]').onclick();
    P('M2 전체 메뉴 창 = 권한 있는 메뉴 ★ 토글 · MANAGER_ONLY 제외 · 대시보드는 고정(★ 토글 없음)', document.querySelectorAll('#modal [data-star]').length > 5 && !document.querySelector('#modal [data-star="dash"]') && !!document.querySelector('#modal [data-am="dash"] .am-star.fixed') && MANAGER_ONLY.every(id => !document.querySelector('#modal [data-am="' + id + '"]')), document.querySelectorAll('#modal [data-star]').length);
    ['so','plan','scan','perf','ship'].forEach(id => document.querySelector('#modal [data-star="' + id + '"]').onclick());   // 회사 기본 7 중 5 제외 → lot·qc 만
    top = [...document.getElementById('nav').children].filter(el => el.matches('button[data-page]')).map(b => b.dataset.page);
    P('M3 ★ 토글 즉시 저장 → 내 메뉴 = dash + 2개 · settings/prefs.U-QA.menu_primary 저장 · users 문서에는 쓰지 않음', JSON.stringify(top) === JSON.stringify(['dash','lot','qc']) && JSON.stringify(DB.get('settings','prefs.U-QA').menu_primary) === JSON.stringify(['lot','qc']) && DB.get('users','U-QA').menu_primary === undefined, top.join(',') + ' / ' + JSON.stringify(DB.get('settings','prefs.U-QA')?.menu_primary));
    P('M3 나머지는 전체 메뉴 창 안 (기능 숨김 0)', document.querySelectorAll('#modal [data-am]').length >= 8 && document.querySelector('#modal [data-star="so"]').textContent === '☆', document.querySelectorAll('#modal [data-am]').length);
    closeModal();
    // 다른 사용자는 영향 없음
    SESSION.userId = 'U-DEMO'; renderNav();
    top = [...document.getElementById('nav').children].filter(el => el.matches('button[data-page]')).map(b => b.dataset.page);
    P('M4 다른 계정은 회사 기본 그대로 (계정별 독립)', top.length === 8, top.join(','));
    SESSION.userId = 'U-QA'; renderNav(); document.querySelector('#nav [data-allmenu]').onclick(); document.getElementById('am-reset').onclick();
    top = [...document.getElementById('nav').children].filter(el => el.matches('button[data-page]')).map(b => b.dataset.page);
    P('M5 [회사 기본으로] → 복원', top.length === 8, top.join(','));
    SESSION.userId = null; SESSION.role = 'manager';
    P('버전 ≥ v0.35.8', /^v0\\.35\\.([8-9]|[1-9][0-9])$|^v0\\.3[6-9]/.test(APP_VERSION), APP_VERSION);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nFORMS358 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 26) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 26 (중도 종료·누락)`);
process.exit((fail || out.length !== 26) ? 1 : 0);
