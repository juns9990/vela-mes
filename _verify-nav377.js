// ★ v0.37.7 — 사이드바 = 내 메뉴 + 「전체 메뉴」 창 · prefs 저장 회귀 (SPEC 「개인화 v0.1」 §3 · GPT 🟢 2026-09-23)
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.join(__dirname, 'vela-mes-prototype.html'), 'utf8');
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window;
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.matchMedia = win.matchMedia || (() => ({ matches:false, addEventListener(){}, removeEventListener(){}, addListener(){}, removeListener(){} }));
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const body = `
;(function(){
  const R = window.__T = [];
  const P = (n, ok, note) => R.push({ n, ok, note: String(note||'') });
  const T = s => (s||'').replace(/\\s+/g,' ');
  const nav = () => document.getElementById('nav');
  const side = () => [...nav().querySelectorAll('button[data-page]')].map(b => b.dataset.page);
  const modal = () => document.getElementById('modal');
  const open = () => { nav().querySelector('[data-allmenu]').onclick(); return modal(); };
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'qa'; SESSION.userId = 'U-DEMO';
    location.hash = '#dash'; renderNav(); router();
    // ── 1. 사이드바 구조 ──
    P('S1 사이드바 = 회사 기본(dash+7) · 아이콘 0 · 관리 접이식 0 · 그룹 라벨 0 · 「전체 메뉴」 버튼 1개', JSON.stringify(side()) === JSON.stringify(NAV_PRIMARY_DEFAULT) && !nav().querySelector('svg') && !nav().querySelector('.nav-more,.nav-grp,[data-more],[data-mymenu-btn]') && nav().querySelectorAll('[data-allmenu]').length === 1, side().join(','));
    P('S2 구 API 제거 — openMyMenuModal/setNavMore/navMoreOpen 없음 · menuCanSee/getPrefs/setPrefs/openAllMenuModal 있음', typeof openMyMenuModal === 'undefined' && typeof setNavMore === 'undefined' && typeof navMoreOpen === 'undefined' && typeof menuCanSee === 'function' && typeof getPrefs === 'function' && typeof setPrefs === 'function' && typeof openAllMenuModal === 'function', '');
    // ── 2. 전체 메뉴 창 = canSee 1벌 ──
    let m = open();
    const flags = getFlags();
    const expect = getMenuOrder().map(id => PAGES.find(p => p.id === id)).filter(p => p && menuCanSee(p, flags)).map(p => p.id);
    const got = [...m.querySelectorAll('[data-am]')].map(r => r.dataset.am);
    P('A1 전체 메뉴 창 항목 = getMenuOrder ∩ menuCanSee (관리자 = 전부 · 같은 집합 · 그룹 안 순서 = 회사 메뉴 순서)', got.length === expect.length && expect.every(id => got.includes(id)) && NAV_GROUPS.every(g => { const ids = got.filter(id => g.ids.includes(id)); return JSON.stringify(ids) === JSON.stringify(expect.filter(id => g.ids.includes(id))); }), got.length + '/' + expect.length);
    P('A2 그룹 라벨 = NAV_GROUPS(영업·생산·품질·기술·구매·재고·경영·기준정보) · 아이콘 0', ['영업','생산','품질·기술','구매·재고','경영','기준정보'].every(l => [...m.querySelectorAll('.am-grp-lbl')].some(g => g.textContent === l)) && !m.querySelector('svg:not(.x svg)'), [...m.querySelectorAll('.am-grp-lbl')].map(g=>g.textContent).join('|'));
    P('A3 안내 문구 「관리책임자가 열어준 메뉴만 보입니다」 · 대시보드 ★ 고정(토글 없음)', m.innerHTML.includes('관리책임자가 열어준 메뉴만 보입니다') && !!m.querySelector('[data-am="dash"] .am-star.fixed') && !m.querySelector('[data-star="dash"]'), '');
    P('A4 내 메뉴(회사 기본 7) 는 ★ · 나머지는 ☆', NAV_PRIMARY_DEFAULT.filter(id => id !== 'dash').every(id => m.querySelector('[data-star="' + id + '"]').textContent === '★') && m.querySelector('[data-star="inv"]').textContent === '☆', '');
    // 검색 = DOM show/hide (재렌더 0)
    const q = m.querySelector('#am-q'); const rowsBefore = m.querySelectorAll('[data-am]').length;
    q.value = '재고'; q.oninput({ target:q });
    const vis = [...m.querySelectorAll('[data-am]')].filter(r => !r.hidden).map(r => r.dataset.am);
    P('A5 검색 "재고" → 재고 행만 표시 · 나머지 hidden · 행 수 불변(재렌더 0) · 빈 그룹 hidden', vis.includes('inv') && vis.every(id => PAGES.find(p=>p.id===id).title.includes('재고')) && m.querySelectorAll('[data-am]').length === rowsBefore && [...m.querySelectorAll('[data-am-grp]')].some(g => g.hidden), vis.join(','));
    q.value = ''; q.oninput({ target:q });
    P('A5 검색 지움 → 전부 복귀', [...m.querySelectorAll('[data-am]')].every(r => !r.hidden) && [...m.querySelectorAll('[data-am-grp]')].every(g => !g.hidden), '');
    // 행 클릭 = 이동 + 닫기
    m.querySelector('[data-am-go="inv"]').onclick();
    P('A6 행 클릭 → 이동(#inv) + 창 닫힘', location.hash === '#inv' && !document.getElementById('modal-bg').classList.contains('show'), location.hash);
    router();
    P('A7 내 메뉴 밖 화면 → 「전체 메뉴」 버튼 on + 「· 재고」 힌트', nav().querySelector('[data-allmenu]').classList.contains('on') && T(nav().querySelector('[data-allmenu]').textContent).includes('재고'), T(nav().querySelector('[data-allmenu]').textContent));
    location.hash = '#so'; router();
    P('A7 내 메뉴 화면 → 힌트 해제 · 해당 버튼 on', !nav().querySelector('[data-allmenu]').classList.contains('on') && nav().querySelector('button[data-page="so"]').classList.contains('on'), '');
    // ── 3. ★ 저장 = settings/prefs.{userId} ──
    m = open();
    m.querySelector('[data-star="inv"]').onclick();
    P('B1 ☆ 클릭 → ★ · 사이드바에 재고 추가(회사 메뉴 순서 위치) · settings/prefs.U-DEMO.menu_primary 저장 · users 문서 무변', m.querySelector('[data-star="inv"]').textContent === '★' && side().includes('inv') && JSON.stringify(DB.get('settings','prefs.U-DEMO').menu_primary) === JSON.stringify([...NAV_PRIMARY_DEFAULT.filter(id=>id!=='dash'), 'inv']) && DB.get('users','U-DEMO').menu_primary === undefined && DB.get('settings','prefs.U-DEMO').userId === 'U-DEMO', side().join(','));
    P('B1 사이드바 순서 = 회사 메뉴 순서(getMenuOrder) 유지', JSON.stringify(side()) === JSON.stringify(getMenuOrder().filter(id => ['dash',...DB.get('settings','prefs.U-DEMO').menu_primary].includes(id))), side().join(','));
    m.querySelector('[data-star="so"]').onclick();
    P('B2 ★ 클릭 → ☆ · 사이드바에서 수주 제외 · prefs 갱신 (창은 열린 채)', m.querySelector('[data-star="so"]').textContent === '☆' && !side().includes('so') && !DB.get('settings','prefs.U-DEMO').menu_primary.includes('so') && document.getElementById('modal-bg').classList.contains('show'), side().join(','));
    P('B3 회사 기본 settings/menu.primary 무변 · 다른 계정 무영향', !DB.get('settings','menu')?.primary || JSON.stringify(DB.get('settings','menu').primary) === JSON.stringify(NAV_PRIMARY_DEFAULT), '');
    SESSION.userId = 'U-OTHER'; renderNav();
    P('B3 다른 계정(U-OTHER) = 회사 기본', JSON.stringify(side()) === JSON.stringify(NAV_PRIMARY_DEFAULT), side().join(','));
    SESSION.userId = 'U-DEMO'; renderNav();
    // 회사 기본으로
    closeModal(); m = open(); m.querySelector('#am-reset').onclick();
    P('B4 「회사 기본으로 되돌리기」 → prefs.menu_primary = [] (명시적 회사 기본) · 사이드바 회사 기본 · 창 닫힘', JSON.stringify(DB.get('settings','prefs.U-DEMO').menu_primary) === '[]' && JSON.stringify(side()) === JSON.stringify(NAV_PRIMARY_DEFAULT) && !document.getElementById('modal-bg').classList.contains('show'), side().join(','));
    // ── 4. 레거시 users.menu_primary 읽기 (쓰지 않음) ──
    DB.set('users','U-LEG',{ id:'U-LEG', name:'레거시', pin:'x', role:'operator', active:true, menu_primary:['qc','lot'] });
    SESSION.userId = 'U-LEG'; SESSION.role = 'operator'; renderNav();
    P('C1 prefs 없음 + users.menu_primary 있음 → 레거시 읽기 (dash + lot·qc 순서 = 회사 메뉴 순서)', JSON.stringify(side()) === JSON.stringify(['dash','lot','qc']), side().join(','));
    m = open(); m.querySelector('[data-star="inv"]').onclick();
    P('C2 레거시 사용자 ★ → prefs 신설(레거시 + 추가) · users.menu_primary 는 그대로(삭제·갱신 0)', JSON.stringify(DB.get('settings','prefs.U-LEG').menu_primary) === JSON.stringify(['qc','lot','inv']) && JSON.stringify(DB.get('users','U-LEG').menu_primary) === JSON.stringify(['qc','lot']) && side().includes('inv'), side().join(','));
    m.querySelector('#am-reset').onclick();
    P('C3 레거시 사용자 「회사 기본으로」 → prefs [] 가 레거시보다 우선 → 회사 기본', JSON.stringify(DB.get('settings','prefs.U-LEG').menu_primary) === '[]' && JSON.stringify(side()) === JSON.stringify(NAV_PRIMARY_DEFAULT), side().join(','));
    // ── 5. 권한 · MANAGER_ONLY · 플래그 = 전체 메뉴 창에도 동일 ──
    DB.set('users','U-PERM',{ id:'U-PERM', name:'권한', pin:'x', role:'operator', active:true, perms:{ acct:'none', inv:'view' } });
    SESSION.userId = 'U-PERM'; SESSION.role = 'operator'; renderNav(); m = open();
    P('D1 operator: MANAGER_ONLY(set·masters·audit) 전체 메뉴 창에 0', MANAGER_ONLY.every(id => !m.querySelector('[data-am="' + id + '"]')), '');
    P('D2 권한 none(acct) 창에 0 · view(inv) 창에 있음', !m.querySelector('[data-am="acct"]') && !!m.querySelector('[data-am="inv"]'), '');
    // 보기 전용 화면에서도 prefs 저장 가능 (원장 아님) · 다른 컬렉션은 여전히 차단
    closeModal(); location.hash = '#inv'; router(); m = open();
    m.querySelector('[data-star="inv"]').onclick();
    P('D3 보기 전용(#inv) 화면에서 ★ 저장 허용 — prefs 는 개인 설정(_permBlockWrite 예외) · shipments 쓰기는 여전히 차단', JSON.stringify(DB.get('settings','prefs.U-PERM')?.menu_primary||[]).includes('inv') && _permBlockWrite('shipments') === true && _permBlockWrite('settings','prefs.U-PERM') === false && _permBlockWrite('settings','company') === true, JSON.stringify(DB.get('settings','prefs.U-PERM')?.menu_primary));
    closeModal();
    // ── 5b. prefs 클라이언트 계약 (GPT 조건부 승인 2026-09-23): 내 prefs 가능 · 남의 prefs 차단 · 회사 settings 차단 · 필드 화이트리스트
    P('P1 내 prefs → 저장 가능 (보기 전용 화면 #inv 에서도)', !!setPrefs('U-PERM', { menu_primary:['inv','qc'] }) && JSON.stringify(DB.get('settings','prefs.U-PERM').menu_primary) === '["inv","qc"]', location.hash);
    const beforeOther = JSON.stringify(DB.get('settings','prefs.U-DEMO') || null);
    P('P2 남의 prefs(U-DEMO) → setPrefs 거부(null) · 문서 무변 · _permBlockWrite 도 예외 아님', setPrefs('U-DEMO', { menu_primary:['so'] }) === null && JSON.stringify(DB.get('settings','prefs.U-DEMO') || null) === beforeOther && _permBlockWrite('settings','prefs.U-DEMO') === true, '');
    P('P3 회사 settings(menu·features·company) → 보기 전용 화면에서 차단 유지', _permBlockWrite('settings','menu') === true && _permBlockWrite('settings','features') === true && _permBlockWrite('settings','company') === true && DB.set('settings','menu',{ primary:['dash','so'] }) === null, '');
    P('P4 필드 화이트리스트 — role/perms/hash 같은 키는 버림 · 허용 키 없으면 쓰기 0', setPrefs('U-PERM', { role:'manager', perms:{}, x:1 }) === null && DB.get('settings','prefs.U-PERM').role === undefined && !!setPrefs('U-PERM', { report_format:'pptx', role:'manager' }) && DB.get('settings','prefs.U-PERM').report_format === 'pptx' && DB.get('settings','prefs.U-PERM').role === undefined, JSON.stringify(DB.get('settings','prefs.U-PERM')));
    P('P5 관리자도 남의 prefs 편집 불가 (SPEC §7 범위 밖)', (()=>{ const bak = { role:SESSION.role, userId:SESSION.userId }; SESSION.role = 'manager'; SESSION.userId = 'U-DEMO'; const r = setPrefs('U-PERM', { menu_primary:['so'] }); SESSION.role = bak.role; SESSION.userId = bak.userId; return r === null && JSON.stringify(DB.get('settings','prefs.U-PERM').menu_primary) === '["inv","qc"]'; })(), '');
    SESSION.role = 'manager'; SESSION.userId = 'U-DEMO';
    const f = getFlags(); const offId = Object.keys(f).find(k => f[k] && PAGES.some(p => p.flag === k && !NAV_PRIMARY_DEFAULT.includes(p.id) && !MANAGER_ONLY.includes(p.id)));
    if(offId){ const before = DB.get('settings','features')?.flags; setFlag(offId, false); renderNav(); m = open();
      P('D4 기능 플래그 OFF(' + offId + ') → 전체 메뉴 창에서도 0 (플래그 1벌)', PAGES.filter(p => p.flag === offId).every(p => !m.querySelector('[data-am="' + p.id + '"]')), '');
      closeModal(); DB.set('settings','features', { flags: before || defaultFlags() }); }
    // ── 6. 개발자(VENDOR · userId 없음) ──
    SESSION.userId = null; SESSION.role = 'vendor'; renderNav(); m = open();
    P('E1 userId 없는 개발자 세션 — ☆ 비활성 + 안내 · 회사 기본 표시', [...m.querySelectorAll('[data-star]')].every(b => b.disabled) && m.querySelector('#am-reset').disabled && m.innerHTML.includes('개발자 세션은 내 메뉴를 저장할 수 없습니다') && JSON.stringify(side()) === JSON.stringify(NAV_PRIMARY_DEFAULT), '');
    closeModal();
    // ── 7. 폰 현장 화면 무변 ──
    SESSION.userId = 'U-DEMO'; SESSION.role = 'manager';
    P('F1 CSS — #app.m-mode #nav 숨김 유지 · 구 .nav-more CSS 0 · .nav-grp{display:none}(nav321 정적) 유지', /#app\\.m-mode #nav, #app\\.m-mode \\.foot\\{display:none\\}/.test(document.documentElement.outerHTML) && !/\\.tabnav \\.nav-more\\{/.test(document.documentElement.outerHTML) && /\\.nav-grp\\{display:none\\}/.test(document.documentElement.outerHTML), '');
    P('버전 v0.38.20 · 2026-10-09', APP_VERSION === 'v0.38.20' && APP_DATE === '2026-10-09', APP_VERSION + ' ' + APP_DATE);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nNAV377 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 32) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 32 (중도 종료·누락)`);
process.exit((fail || out.length !== 32) ? 1 : 0);
