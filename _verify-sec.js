// ★ v0.38.17 — 화면 단락 켜기/끄기 + 외주 정산 → 마감 일원화 회귀
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.resolve(__dirname, process.env.VELA_FILE || 'vela-mes-prototype.html'), 'utf8')   // GitHub 저장소에서는 VELA_FILE=index-dev.html;
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const vc = new VirtualConsole(); const errs = []; vc.on('jsdomError', e => errs.push(String(e && e.message || e)));
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:vc, url:'http://localhost/' });
const win = dom.window; win.VELA_BACKEND = 'local';
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.matchMedia = q => ({ matches:false, media:q, addListener(){}, removeListener(){}, addEventListener(){}, removeEventListener(){} });
win.eval(scripts.join('\n;\n') + `;window.__X = s => eval(s);`);
const X = s => win.__X(s);
const R = []; const P = (n, ok, note) => { R.push({ n, ok }); console.log(`[${ok?'PASS':'FAIL'}] ${n}${note?'  — '+note:''}`); };
X(`seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='검증'; SESSION.userId='U-SEC';`);
const ledger = () => X(`JSON.stringify(['orders','shipments','production_lots','records','os_dispatches','os_receipts','closings'].map(c => DB.allRaw(c).length))`);
const L0 = ledger();
const go = h => X(`location.hash='${h}'; router(); 1`);

// 1 외주 정산 카드: 마감 ON → 없음 + [외주 정산 → 마감] · OFF → 있음
const o1 = X(`(()=>{ location.hash='#os'; router(); const v = document.getElementById('view');
  const on = { settle: /월별 외주처 정산/.test(v.textContent), goto: !!document.getElementById('os-goto-closing'), flag: !!getFlags().closing };
  const f0 = DB.get('settings','features') || {}; return on; })()`);
P('1 마감 모듈 ON → 외주 화면에 「월별 외주처 정산」 없음 · [외주 정산 → 마감] 바로 가기', o1.flag && !o1.settle && o1.goto, JSON.stringify(o1));
const o2 = X(`(()=>{ const gf = getFlags; getFlags = () => ({ ...gf(), closing:false }); try { location.hash='#os'; router(); const v = document.getElementById('view');
  return { settle: /월별 외주처 정산/.test(v.textContent), goto: !!document.getElementById('os-goto-closing'), sel: !!document.getElementById('os-settle-month') }; } finally { getFlags = gf; } })()`);
P('2 마감 모듈 OFF 업체 → 정산 카드 그대로 (정산 경로 유지)', o2.settle && !o2.goto && o2.sel, JSON.stringify(o2));
const o3 = X(`(()=>{ location.hash='#os'; router(); document.getElementById('os-goto-closing').click(); const h = location.hash; router();
  return { h, type: window._closingState && window._closingState.type, txt: /외주 \\(줄 돈\\)/.test(document.getElementById('view').textContent) }; })()`);
P('3 [외주 정산 → 마감] → 마감 화면 외주 탭으로 이동', o3.h === '#closing' && o3.type === 'outsource' && o3.txt, JSON.stringify(o3));

const o4 = X(`(()=>{ const role = SESSION.role, gp = getUserPerm; SESSION.role = 'operator'; SESSION.userId = 'U-NOCL';
  getUserPerm = (u, id) => id === 'closing' ? 'none' : 'edit';
  try { location.hash='#os'; router(); const v = document.getElementById('view');
    return { settle: /월별 외주처 정산/.test(v.textContent), ref: /참고용/.test(v.textContent), goto: !!document.getElementById('os-goto-closing') }; }
  finally { getUserPerm = gp; SESSION.role = role; SESSION.userId = 'U-SEC'; } })()`);
P('3b (GPT 소스 검토 P1) 마감 ON 이어도 마감 화면 권한이 없는 사용자 → 외주 화면 정산 카드 유지(「참고용 · 확정 = 마감」 표기) · 바로 가기 없음', o4.settle && o4.ref && !o4.goto, JSON.stringify(o4));
// 4~9 외주 화면 단락 끄기/켜기
go('#os');
const s1 = X(`(()=>{ const l = secList(); return { titles: l.map(s => s.title), sw: document.querySelectorAll('#view .sec-x').length, btn: document.getElementById('sec-pick-btn')?.textContent,
  order: (() => { const h = l[0].h; const kids = [...h.children]; return kids[0] && kids[0].classList.contains('sec-x'); })() }; })()`);
P('4 외주 화면 단락 = 외주 현황 · 최근 외주 입고 · 제목 옆 스위치 · 머리 [단락 2/2]', JSON.stringify(s1.titles) === JSON.stringify(['외주 현황','최근 외주 입고']) && s1.sw === 2 && s1.btn === '단락 2/2' && s1.order, JSON.stringify(s1));
const s2 = X(`(()=>{ const l = secList(); l[1].h.querySelector('.sec-x').click(); const c = l[1].card;
  return { off: c.classList.contains('sec-off'), pref: JSON.stringify(getPrefs('U-SEC').sec_off), btn: document.getElementById('sec-pick-btn').textContent, inDom: document.body.contains(c) }; })()`);
P('5 스위치 끄기 → 화면에서만 숨김(DOM 유지) · prefs.U-SEC.sec_off 저장 · [단락 1/2]', s2.off && s2.inDom && s2.pref === '{"os:최근 외주 입고":true}' && s2.btn === '단락 1/2', JSON.stringify(s2));
const s3 = X(`(()=>{ router(); const c = secList().find(s => s.title === '최근 외주 입고').card; return c.classList.contains('sec-off'); })()`);
P('6 다시 그려도(router) 꺼진 상태 유지', s3 === true);
const s4 = X(`(()=>{ openSecPicker(); const rows = [...document.querySelectorAll('#sec-pick-list .sec-pick')].map(r => r.textContent.trim());
  const b = document.querySelector('#sec-pick-list [data-sec-key="os:최근 외주 입고"]'); const wasOff = b.classList.contains('off'); b.click();
  const c = secList().find(s => s.title === '최근 외주 입고').card; const vis = !c.classList.contains('sec-off'); closeModal();
  return { rows, wasOff, vis, pref: JSON.stringify(getPrefs('U-SEC').sec_off) }; })()`);
P('7 [단락] 목록에서 선택해 다시 켜기 → 보임 · prefs 에서 제거', s4.rows.length === 2 && s4.wasOff && s4.vis && s4.pref === '{}', JSON.stringify(s4));
const s5 = X(`(()=>{ secSetOff('os:외주 현황', true); secSetOff('os:최근 외주 입고', true); router(); const n0 = document.querySelectorAll('#view .card.sec-off').length;
  openSecPicker(); document.getElementById('sec-all-on').click(); closeModal(); const n1 = document.querySelectorAll('#view .card.sec-off').length; return [n0, n1]; })()`);
P('8 [모두 켜기] → 이 화면 단락 전부 표시', s5[0] === 2 && s5[1] === 0, JSON.stringify(s5));
const s6 = X(`(()=>{ secSetOff('os:외주 현황', true); SESSION.userId = 'U-OTHER'; router(); const other = document.querySelectorAll('#view .card.sec-off').length;
  SESSION.userId = 'U-SEC'; router(); const me = document.querySelectorAll('#view .card.sec-off').length; secSetOff('os:외주 현황', false); return [me, other]; })()`);
P('9 사람마다 저장 — 다른 사용자 화면은 그대로', s6[0] === 1 && s6[1] === 0, JSON.stringify(s6));

// 10 안전 신호 단락 = 끌 수 없음
const s7 = X(`(()=>{ const v = document.getElementById('view'); v.insertAdjacentHTML('beforeend', '<div class="card"><h3>오늘 이상 3건</h3><div>x</div></div>');
  _secToggleEnhance(); const s = secList().find(x => /오늘 이상/.test(x.title)); secSetOff(s.key, true); _secToggleEnhance();
  return { locked: s.locked, sw: !!s.h.querySelector('.sec-x'), off: s.card.classList.contains('sec-off'), key: s.key }; })()`);
P('10 안전 신호(이상·HOLD·격리·주의·경고·보류·판정 대기) 단락 = 스위치 없음 · prefs 에 있어도 숨기지 않음', s7.locked && !s7.sw && !s7.off, JSON.stringify(s7));
X(`secSetOff('os:오늘 이상 건', false); setPrefs('U-SEC', { sec_off:{} }); 1`);

// 10b~10f — v0.38.18 GPT 재판정 결정 1
const m1 = X(`(()=>{ location.hash='#qc'; router(); const s = secList().find(x => x.key === 'qc:qc-nc-list'); secSetOff('qc:qc-nc-list', true); _secToggleEnhance();
  const r = { found: !!s, lockBy: s && s.lockBy, sw: s && !!s.h.querySelector('.sec-x'), off: s && s.card.classList.contains('sec-off') };
  setPrefs('U-SEC', { sec_off:{} }); location.hash='#audit'; router(); const a = secList().find(x => x.key === 'audit:audit-incident'); r.audit = a && a.lockBy; return r; })()`);
P('10b 명시 표식 data-sec-lock: 부적합 목록(HOLD·판정)·동시성 사고 = 잠금 1순위(mark) · 스위치 없음 · 저장값 있어도 표시', m1.found && m1.lockBy === 'mark' && !m1.sw && !m1.off && m1.audit === 'mark', JSON.stringify(m1));
const m2 = X(`(()=>{ const v = document.getElementById('view');
  v.insertAdjacentHTML('beforeend', '<div class="card" data-sec-lock="safety" data-sec-id="t-plain"><h3>평범한 제목</h3><div>x</div></div><div class="card"><h3>원장 정합성 경고</h3><div>y</div></div>');
  _secToggleEnhance(); const l = secList(); const a = l.find(x => x.key.endsWith(':t-plain')), b = l.find(x => /정합성/.test(x.title));
  return { a: a && a.lockBy, aSw: a && !!a.h.querySelector('.sec-x'), b: b && b.lockBy }; })()`);
P('10c 키워드 없는 제목이라도 표식이 있으면 잠금 · 표식 없는 단락은 제목 키워드가 보조 방어(정합성 등)', m2.a === 'mark' && !m2.aSw && m2.b === 'keyword', JSON.stringify(m2));
const m3 = X(`(()=>{ location.hash='#plan'; router(); const k1 = (secList().find(x => x.card.dataset.secId === 'plan-day') || {}).key;
  const v = document.getElementById('view'); v.insertAdjacentHTML('beforeend', '<div class="card"><h3>중복 제목 <span class="tag">1</span></h3><div>a</div></div><div class="card"><h3>중복 제목 <span class="tag">2</span></h3><div>b</div></div>');
  _secToggleEnhance(); const d = secList().filter(x => x.title === '중복 제목').map(x => x.key);
  secSetOff(d[1], true); _secToggleEnhance(); const offs = secList().filter(x => x.title === '중복 제목').map(x => x.card.classList.contains('sec-off'));
  secSetOff(d[1], false); return { k1, d, offs }; })()`);
P('10d 단락 키 고유: 명시 ID(plan:plan-day — 날짜가 바뀌어도 동일) · 같은 제목 2개 = 다른 키 · 하나만 꺼짐', m3.k1 === 'plan:plan-day' && m3.d.length === 2 && m3.d[0] !== m3.d[1] && JSON.stringify(m3.offs) === '[false,true]', JSON.stringify(m3));
const m4 = X(`(()=>{ location.hash='#os'; router(); setPrefs('U-SEC', { sec_off:{} }); router(); const l = secList();
  l[0].h.querySelector('.sec-x').click(); const l2 = secList(); const sw2 = l2[1].h.querySelector('.sec-x'); sw2.click();
  const visible = secList().filter(x => !x.card.classList.contains('sec-off')).length; const can = secCanTurnOff(l2[1].key);
  setPrefs('U-SEC', { sec_off:{} }); router(); return { visible, can }; })()`);
P('10e 화면의 마지막 보이는 단락은 끌 수 없음 (빈 화면 방지)', m4.visible === 1 && m4.can === false, JSON.stringify(m4));
// 10f — v0.38.19 잠금 전수 조사: HOLD·사고 상태 배지를 보여주는 단락은 전부 명시 잠금
const audit = X(`(()=>{ const bad = [], lockedMarks = []; const pages = PAGES.filter(p => getFlags()[p.flag] && p.id !== 'dash').map(p => [p.id, null]).concat([['matin','buy'],['matin','os'],['ship','otd']]);
  pages.forEach(([id, t]) => { if(id === 'matin') window._matinHub = { tab: t || 'material' }; window._shTab = t === 'otd' ? 'otd' : 'list'; location.hash = '#' + id; router();
    secList().forEach(s => { if(s.lockBy === 'mark') lockedMarks.push(s.key);
      const badge = [...s.card.querySelectorAll('.pill,.badge,.tag')].some(e => e !== s.h && !s.h.contains(e) && /HOLD|사고/.test(e.textContent));
      if(badge && s.lockBy !== 'mark') bad.push(s.key); }); });
  window._matinHub = { tab:'material' }; window._shTab = 'list'; return { bad, lockedMarks }; })()`);
P('10f 전수 조사(30화면): HOLD·사고 배지를 보여주는 단락은 모두 명시 잠금 · 명시 잠금 = 생산 LOT·부적합 목록·동시성 사고', audit.bad.length === 0 && ['lot:lot-list','qc:qc-nc-list','audit:audit-incident'].every(k => audit.lockedMarks.includes(k)), JSON.stringify(audit));
// 11 대시보드 제외 · 12 전 메뉴 순회
const d = X(`(()=>{ location.hash='#dash'; router(); return { sw: document.querySelectorAll('#view .sec-x').length, btn: !!document.getElementById('sec-pick-btn') }; })()`);
P('11 대시보드 = 제외 (카드 Picker 별도)', d.sw === 0 && !d.btn, JSON.stringify(d));
const sweep = X(`(()=>{ const out = []; PAGES.filter(p => getFlags()[p.flag] && p.id !== 'dash').forEach(p => { try { location.hash = '#' + p.id; router();
    const l = secList(); const can = l.filter(s => !s.locked); out.push({ id:p.id, n:l.length, sw: document.querySelectorAll('#view .sec-x').length, can: can.length, btn: document.getElementById('sec-pick-btn')?.textContent || '', locked: l.filter(s=>s.locked).map(s=>s.title) });
  } catch(e){ out.push({ id:p.id, err: String(e.message) }); } }); return out; })()`);
const bad = sweep.filter(x => x.err || x.sw !== x.can || (x.can > 0 && x.btn !== `단락 ${x.n}/${x.n}`));
P('12 전 메뉴 순회: 예외 0 · 끌 수 있는 단락마다 스위치 1 · [단락 n/n]', bad.length === 0 && sweep.length >= 15, `${sweep.length}화면 · 단락 합계 ${sweep.reduce((a,x)=>a+(x.n||0),0)} · ` + JSON.stringify(bad));
console.log('   화면별 단락:', sweep.map(x => `${x.id}:${x.n}${x.locked && x.locked.length ? '(고정 '+x.locked.join('/')+')' : ''}`).join(' · '));
// 13 보기 전용 사용자도 내 단락 설정은 저장 · 남의 prefs 는 거부
const pv = X(`(()=>{ const role = SESSION.role; SESSION.role = 'operator'; SESSION.userId = 'U-VIEW';
  const gp = getUserPerm; getUserPerm = () => 'view';
  try { location.hash = '#os'; router(); const l = secList(); l[0].h.querySelector('.sec-x').click(); const mine = JSON.stringify((getPrefs('U-VIEW')||{}).sec_off||{});
    const other = setPrefs('U-SEC', { sec_off:{ 'os:외주 현황':true } }); return { mine, other: other === null }; }
  finally { getUserPerm = gp; SESSION.role = role; SESSION.userId = 'U-SEC'; } })()`);
P('13 보기 전용 권한이어도 내 화면 단락은 저장 · 남의 설정은 저장 거부(setPrefs 본인 계약)', pv.mine === '{"os:외주 현황":true}' && pv.other, JSON.stringify(pv));
// 14 원장 무변 · 15 OTD 화면에도 적용
const otd = X(`(()=>{ window._shTab='otd'; location.hash='#ship'; router(); const t = secList().map(s => s.title); window._shTab='list'; return t; })()`);
P('14 출하 › 납기 실적 화면에도 적용 (고객사별 · 품목별 · 지연·미출하 수주)', JSON.stringify(otd) === JSON.stringify(['고객사별','품목별','지연·미출하 수주']), JSON.stringify(otd));
const L1 = ledger();
P('15 원장 무변 — 수주·출하·LOT·실적·외주·마감 문서 수 변화 0 (쓰기 = settings/prefs 만)', L0 === L1, L0 + ' → ' + L1);
P('16 콘솔 jsdom 오류 0', errs.length === 0, errs.slice(0,3).join(' | '));
(async () => {
  X(`location.hash='#os'; router(); secSetOff('os:최근 외주 입고', true); _secToggleEnhance(); const v = document.getElementById('view'); v.innerHTML = v.innerHTML.replace(/sec-off/g,''); 1`);
  await new Promise(r => setTimeout(r, 0));
  const m = X(`(()=>{ const r = { btn: document.getElementById('sec-pick-btn')?.textContent, nBtn: document.querySelectorAll('#sec-pick-btn').length, off: secList().find(s => s.title === '최근 외주 입고').card.classList.contains('sec-off') };
    secSetOff('os:최근 외주 입고', false); _secToggleEnhance(); const sw = secList().find(s => s.title === '외주 현황').h.querySelector('.sec-x'); r.bound = !!(sw && sw._secBound); sw.click(); r.clickWorks = secList().find(s => s.title === '외주 현황').card.classList.contains('sec-off'); secSetOff('os:외주 현황', false); return r; })()`);
  X(`secSetOff('os:최근 외주 입고', false); 1`);
  P('6b 화면이 router 밖에서 다시 그려져도(innerHTML 교체) [단락]·꺼짐 재부착 (MutationObserver)', m.btn === '단락 1/2' && m.nBtn === 1 && m.off && m.bound && m.clickWorks, JSON.stringify(m));
  // 6c 반복 재렌더 — 스위치 중복 0 · 관찰자 반복 실행 없음
  X(`_secEnhanceRuns = 0; location.hash='#os'; for(let i=0;i<5;i++) router(); 1`);
  await new Promise(r => setTimeout(r, 20));
  const loop = X(`(()=>{ const runs0 = _secEnhanceRuns; return { runs0, sw: document.querySelectorAll('#view .sec-x').length, perH3: [...document.querySelectorAll('#view .card h3')].map(h => h.querySelectorAll('.sec-x').length), btn: document.querySelectorAll('#sec-pick-btn').length }; })()`);
  await new Promise(r => setTimeout(r, 50));
  const runs1 = X(`_secEnhanceRuns`);
  P('6c 5회 연속 다시 그림 → 스위치 단락당 1개 · [단락] 1개 · 정지 후 추가 실행 0 (무한 반복 없음)', loop.sw === 2 && loop.perH3.every(n => n <= 1) && loop.btn === 1 && runs1 === loop.runs0 && loop.runs0 <= 15, JSON.stringify({ ...loop, runs1 }));
const fail = R.filter(x => !x.ok).length;
console.log(`\nSEC TOTAL ${R.length} · PASS ${R.length - fail} · FAIL ${fail}`);
process.exit(fail ? 1 : 0);
})();
