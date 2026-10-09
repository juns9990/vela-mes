// R4 별도 회귀 — 자동계획 병목 판정 (FORGE 하드코딩 → 부하율 최고 사내 공정) · GPT 마감 조건 ②
// 케이스: 기존 동등성(단조 데모에서 병목=FORGE 유지) · 공정 1개 · 동률(결정성) · 사내/외주 혼재(out 제외) · 부하 데이터 없음
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
  try {
    seedDemoData();
    SESSION.loggedIn = true; SESSION.role = 'manager'; SESSION.user = 'r4';
    // 공통: 병목 기대값 재계산 함수 (신 로직 정의와 독립 대조)
    const expectBn = () => {
      const cap = {}; computeCapa().forEach(c => cap[c.proc.code] = c);
      let best = null, bl = -1;
      Object.values(cap).forEach(c => {
        if(!c.dailyCapa || c.proc?.type === 'out') return;
        const l = c.assigned / c.dailyCapa;
        if(l > bl){ bl = l; best = c.proc.code; }
      });
      return best;
    };
    // T1 기존 동등성: 단조 데모 + FORGE 과부하 강제 → 경고가 FORGE 를 지목 (구 하드코딩과 동일 입력·동일 결과)
    DB.set('orders','SO-R4-BIG',{ id:'SO-R4-BIG', cust:'C001', itemCode:'HF-3301', qty: 50000, price:1000, due: fmtDate(Date.now()+3*86400000), created: Date.now() });
    const g1 = generateAutoPlan();
    const bn1 = expectBn();
    const warned = g1.items.filter(p => p.capaWarn);
    P('T1 병목 자동 판정 = 재계산 기대값 일치', warned.length > 0 && warned.every(p => p.capaWarn.startsWith(bn1 + ' 부하')), '기대 ' + bn1 + ' · 경고 ' + (warned[0]?.capaWarn||'없음'));
    P('T1 단조 데모 동등성 — 병목이 FORGE (구 로직과 동일 결과)', bn1 === 'FORGE', '병목=' + bn1);
    // T2 사내/외주 혼재: 외주(HEAT·type out)는 절대 병목으로 선정 안 됨
    P('T2 외주 공정 제외', bn1 !== 'HEAT' && warned.every(p => !p.capaWarn.startsWith('HEAT')), '');
    // T3 결정성 (동률 포함): 같은 입력 2회 → 같은 판정
    const g2 = generateAutoPlan();
    P('T3 결정성 — 2회 실행 동일 경고', JSON.stringify(g1.items.map(p=>p.capaWarn)) === JSON.stringify(g2.items.map(p=>p.capaWarn)), '');
    // T4 부하 데이터 없음: 설비 능력 전부 제거 → 예외 0 · capaWarn 전부 null
    const snap = DB.snapshot();
    try {
      DB.allRaw('equipment').forEach(e => DB.set('equipment', e.id, { hourly_capa: 0 }));  // computeCapa 원천 = 'equipment' 컬렉션
      const g3 = generateAutoPlan();
      P('T4 부하 데이터 없음 — 예외 0 · 경고 없음', Array.isArray(g3.items) && g3.items.every(p => !p.capaWarn), 'items ' + g3.items.length);
      // T5 공정 1개만 능력 보유: CUT 만 능력 복구 → 병목 = CUT (유일 후보)
      const cutEq = DB.allRaw('equipment').find(e => e.proc === 'CUT');
      if(cutEq){ DB.set('equipment', cutEq.id, { hourly_capa: 1, working_hours: 8 }); }
      const g4 = generateAutoPlan();
      const w4 = g4.items.filter(p => p.capaWarn);
      P('T5 공정 1개 — 유일 능력 공정(CUT)이 병목', w4.length === 0 || w4.every(p => p.capaWarn.startsWith('CUT 부하')), w4[0]?.capaWarn || '경고 없음(용량 충분)');
    } finally { DB.restore(snap); }
    P('T6 스냅샷 복원', DB.all('orders').length >= 40, '');
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nR4 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 7) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 7 (중도 종료·누락)`);
process.exit((fail || out.length !== 7) ? 1 : 0);
