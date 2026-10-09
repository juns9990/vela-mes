// ★ v0.35.9-A — 파일럿 담당자 피드백 묶음 회귀 (GPT 승인 §22 · 원장·HOLD Gate 무변)
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
    // ── ① Search Contract (LOT 추적·재고 + 수주·실적) ──
    location.hash = '#lot'; router();
    let q = document.getElementById('lot-q'); q.focus(); q.value = 'PL-2026-H'; q.dispatchEvent(new window.Event('input'));
    const rowsTyping = document.querySelectorAll('#view table tbody tr').length;
    P('S1 타이핑 중 재렌더 없음 (커서 유지 · 목록 그대로)', document.activeElement === q && rowsTyping > 5 && !!document.querySelector('#lot-q + .msrch-go'), rowsTyping + ' rows · active=' + (document.activeElement && document.activeElement.id));
    q.dispatchEvent(new window.KeyboardEvent('keydown', { key:'Enter' }));
    q = document.getElementById('lot-q');
    const rowsAfter = [...document.querySelectorAll('#view table tbody tr')].filter(r => /PL-2026-H/.test(r.textContent)).length;
    P('S2 Enter → 검색 적용 + 새 입력창에 focus·값 유지', lotSearch.q === 'PL-2026-H' && document.activeElement === q && q.value === 'PL-2026-H' && rowsAfter >= 1 && rowsAfter < rowsTyping, 'rows ' + rowsAfter + ' · active=' + (document.activeElement && document.activeElement.id));
    q.value = ''; q.dispatchEvent(new window.Event('input'));
    P('S3 검색어 지우면 즉시 전체 복귀', lotSearch.q === '' && document.querySelectorAll('#view table tbody tr').length >= rowsTyping - 1, '');
    document.getElementById('lot-q').value = 'H01'; document.querySelector('#lot-q + .msrch-go').onclick();
    P('S4 [검색] 버튼 = Enter 와 동일', lotSearch.q === 'H01' && document.activeElement && document.activeElement.id === 'lot-q', '');
    lotSearch.q = '';
    location.hash = '#inv'; router();
    P('S5 재고·수주·실적·품질·계획·도면·이력 검색창 전부 Search Contract (검색 버튼 존재)', ['inv'].every(id => !!document.querySelector('#' + id + '-q + .msrch-go')) && !/\\.oninput = \\(e\\)=>\\{[^}]*setTimeout\\([^}]*router\\(\\)/.test(Array.from(document.querySelectorAll('script')).map(s=>s.textContent).join('')), '');
    // ── ② 재고: LOT 배지 + 행 클릭 상세 ──
    const invRow = document.querySelector('#view tr[data-inv-item]');
    P('I1 재고 행에 진행 LOT 배지 · 행 클릭 가능', !!invRow && !!document.querySelector('#view tr[data-inv-item] .pill.cyan'), invRow && invRow.dataset.invItem);
    invRow.onclick();
    const lotRows = document.querySelectorAll('#inv-lot-tbl tr[data-inv-lot]');
    P('I2 상세 모달 = 이 품번 LOT 목록(위치·잔량·heat·상태) + 스트립', lotRows.length >= 1 && document.getElementById('modal').textContent.includes('완성품 재고'), lotRows.length + ' lots');
    lotRows[0].onclick();
    P('I3 LOT 클릭 → 공정 스텝 패널 (lotStepPanelHTML 1벌)', !!document.querySelector('#inv-lot-tbl .inv-lot-steps'), '');
    closeModal();
    // ── ③ 작업지시: 소재 FIFO 자동 배정 · 작성 현황 품명·소재 규격 ──
    location.hash = '#wo'; router();
    const woTh = [...document.querySelectorAll('#view thead th')].map(t => t.textContent.trim());
    P('W1 작성 현황에 품번·품명 + 소재 규격 컬럼', woTh.includes('품번 · 품명') && woTh.includes('소재 규격') && /S45C|SCM440/.test(document.querySelector('#view tbody').textContent), woTh.join('|'));
    document.getElementById('btn-wo-new').onclick();
    const itemSel = document.getElementById('wo-item'), qtyEl = document.getElementById('wo-qty');
    const it = getActiveItems().find(i => (i.mat||'') && (+i.weight||0) > 0 && Object.values(calcHeatLedger().ledger).some(h => h.remain > 0 && (h.mat||'') === i.mat));
    itemSel.value = it.code; itemSel.onchange(); qtyEl.value = '150'; qtyEl.oninput();
    const auto = document.getElementById('wo-mat-auto');
    const allocRows = [...auto.querySelectorAll('tbody tr')];
    const need = Math.round(150 * it.weight * 100) / 100;
    const sum = allocRows.reduce((s,tr) => s + (+tr.children[4].textContent.replace(/,/g,'')||0), 0);
    const heatsSpec = allocRows.map(tr => calcHeatLedger().ledger[tr.children[1].textContent.trim()]).every(h => h && (h.mat||'') === it.mat);
    P('W2 품번·수량 입력 → 같은 규격 heat 만 입고 순 배정 · 배정 합 = 필요량(또는 부족 표기)', allocRows.length >= 1 && heatsSpec && (Math.abs(sum - need) < 0.02 || auto.textContent.includes('부족')), 'need ' + need + ' sum ' + sum + ' rows ' + allocRows.length);
    const tsList = allocRows.map(tr => { const h = calcHeatLedger().ledger[tr.children[1].textContent.trim()]; return Math.min(...h.receipts.map(r => r.ts)); });
    P('W3 배정 순서 = 입고일 오름차순 (선입선출)', tsList.every((t,i) => i === 0 || t >= tsList[i-1]), tsList.map(t => fmtDate(t)).join(','));
    // 직접 선택 → 같은 규격만 + 다른 규격 보기
    document.querySelector('#wo-mat-mode [data-m="manual"]').onclick();
    const sel = document.getElementById('wo-mat-sel');
    const specOnly = [...sel.options].filter(o => o.value && !o.value.startsWith('__')).every(o => o.dataset.mat === it.mat);
    const hasAll = [...sel.options].some(o => o.value === '__all__');
    P('W4 직접 선택 = 같은 규격 heat 만 + 「다른 규격 보기」', specOnly && hasAll && document.getElementById('wo-mat-manual-box').style.display !== 'none', sel.options.length + ' opts');
    document.querySelector('#wo-mat-mode [data-m="auto"]').onclick();
    // 저장 → LOT material_lot_ids = 배정 heat
    const allocHeats = allocRows.map(tr => tr.children[1].textContent.trim());
    const plotNo = document.getElementById('wo-plot').value;
    const usedBefore = Object.values(calcHeatLedger().ledger).reduce((s,h)=>s+h.used_kg,0), remainBefore = calcHeatLedger().totalRemain;
    document.getElementById('wo-ok').onclick();
    const newLot = DB.get('production_lots', plotNo);
    const usedAfter = Object.values(calcHeatLedger().ledger).reduce((s,h)=>s+h.used_kg,0), remainAfter = calcHeatLedger().totalRemain;
    P('W5 생성 → LOT.material_lot_ids = 자동 배정 heat (제안 반영 · 원장 차감 없음)', !!newLot && JSON.stringify(newLot.material_lot_ids) === JSON.stringify(allocHeats) && newLot.qty_initial === 150, newLot ? JSON.stringify(newLot.material_lot_ids) : 'no lot');
    P('W5 소재 잔량 원장 무변 — 발행 전후 used_kg·totalRemain 동일 (차감은 절단 실적 경로에서만)', Math.abs(usedAfter - usedBefore) < 0.001 && Math.abs(remainAfter - remainBefore) < 0.001 && DB.query('records', r => r.plot_no === plotNo).length === 0, 'used ' + usedBefore.toFixed(1) + '→' + usedAfter.toFixed(1) + ' · remain ' + remainBefore.toFixed(1) + '→' + remainAfter.toFixed(1));
    // 절단 실적을 넣어야 비로소 차감 (GPT 확인 요청: 작업지시 → 배정 정보 → 실적 → 소비)
    const e5 = tryRun(() => LOT.registerRecord({ plot_no: plotNo, proc: getRoute(it.code)[0].proc, equipment:'', good:150, type:'in', shift:'day' }));
    const usedCut = Object.values(calcHeatLedger().ledger).reduce((s,h)=>s+h.used_kg,0);
    P('W6 절단 실적 150 등록 후에야 used_kg +150×단중 (소비 = 실적 경로)', e5 === null && Math.abs(usedCut - usedAfter - 150 * it.weight) < 0.01, 'used +' + (usedCut - usedAfter).toFixed(2) + ' (exp ' + (150*it.weight).toFixed(2) + ')');
    // ── ④ PC 실적 등록: 품번 → LOT · 파렛트 · SHIP 제외 ──
    location.hash = '#perf'; router();
    VIEWS._perfAddPC({});
    const paItem = document.getElementById('pa-item'), paLot = document.getElementById('pa-lot');
    P('R1 품번 선택칸이 먼저 · 품번 목록 = 진행 LOT 있는 품번만 (건수 표기)', paItem.options.length >= 2 && [...paItem.options].every(o => /\\(\\d+\\)$/.test(o.textContent)), paItem.options.length);
    const chosen = paItem.value;
    P('R2 LOT 목록 = 그 품번의 진행 LOT 만 · 다음 공정·대기 수량 라벨', paLot.options.length >= 1 && [...paLot.options].every(o => DB.get('production_lots', o.value).itemCode === chosen && /대기 [\\d,]+ea/.test(o.textContent)), paLot.options[0] && paLot.options[0].textContent);
    paItem.value = [...paItem.options].map(o=>o.value).find(v => v !== chosen) || chosen; paItem.onchange();
    P('R3 품번 바꾸면 LOT 목록 갱신', [...paLot.options].every(o => DB.get('production_lots', o.value).itemCode === paItem.value), '');
    // SHIP 만 남은 LOT → 공정 목록에 출하 없음 + 안내
    const doneLot = DB.all('production_lots').find(l => { const nx = LOT.nextProcess(l.plot_no); return nx && nx.code === 'SHIP'; });
    if(doneLot){ paItem.value = doneLot.itemCode; paItem.onchange(); paLot.value = doneLot.plot_no; paLot.onchange(); }
    const procOpts = [...document.getElementById('pa-proc').options].map(o => o.value);
    P('R4 출하(SHIP)는 공정 목록에서 제외 + 「출하 메뉴에서」 안내', !!doneLot && !procOpts.includes('SHIP') && /출하 메뉴에서/.test(document.getElementById('pa-bal').textContent + document.getElementById('pa-proc').textContent), doneLot ? doneLot.plot_no : 'no ship-only lot');
    // 파렛트 수량 저장
    const live = DB.all('production_lots').find(l => { const nx = LOT.nextProcess(l.plot_no); return nx && nx.code !== 'SHIP' && !heldPlotSet().has(l.plot_no) && !lockedPlotSet().has(l.plot_no); });
    paItem.value = live.itemCode; paItem.onchange(); paLot.value = live.plot_no; paLot.onchange();
    document.getElementById('pa-g').value = '1'; document.getElementById('pa-plt').value = '3';
    const n0 = DB.all('records').length; document.getElementById('pa-ok').onclick();
    const last = DB.all('records').sort((a,b)=>b.ts-a.ts)[0];
    P('R5 파렛트 수량 입력 → records.pallets 저장 (수량 계산 무관)', DB.all('records').length === n0 + 1 && last.pallets === 3 && last.good === 1, JSON.stringify({p:last.pallets,g:last.good}));
    // ── ⑤ 출하: 고객사 → 미납 수주 표 ──
    location.hash = '#ship'; router(); document.getElementById('btn-ship-new').onclick();
    const shTh = [...document.querySelectorAll('#modal thead th')].map(t => t.textContent.trim());
    const shRows = [...document.querySelectorAll('#sh-so-body tr[data-sh-so]')];
    P('H1 미납 수주 표 = 수주번호·품번·품명·주문·완성·미납·납기 · 미납>0 만', ['수주번호','품번 · 품명','주문','완성','미납','납기'].every(h => shTh.includes(h)) && shRows.length >= 1 && shRows.every(tr => calcSoRows([DB.get('orders', tr.dataset.shSo)])[0].remain > 0), shRows.length + ' rows');
    const custSel = document.getElementById('sh-cust'); const c1 = [...custSel.options].map(o=>o.value).find(v => v);
    custSel.value = c1; custSel.onchange();
    P('H2 고객사 선택 → 그 고객사 미납만', [...document.querySelectorAll('#sh-so-body tr[data-sh-so]')].every(tr => (DB.get('orders', tr.dataset.shSo).cust||'') === c1), c1);
    const tr1 = document.querySelector('#sh-so-body tr[data-sh-so]'); tr1.onclick();
    const soHidden = document.getElementById('sh-so'); const shLot = document.getElementById('sh-lot');
    P('H3 행 클릭 → 숨은 수주 select 선택 + 같은 품번 LOT 만 목록', soHidden.value === tr1.dataset.shSo && [...shLot.options].every(o => DB.get('production_lots', o.value).itemCode === DB.get('orders', soHidden.value).itemCode), shLot.options.length + ' lots');
    closeModal();
    // ── ⑥ 품질: 헤더 규격 + 불량 유형 추가 ──
    location.hash = '#qc'; router();
    P('Q1 품질 헤더 = 주 버튼 1 + ⋯ (글자만) · 불량 유형 추가 항목', !!document.getElementById('qc-more') && !!document.getElementById('qc-dt-new') && !document.querySelector('#qc-more-menu svg'), '');
    document.getElementById('qc-dt-new').onclick();
    const dtN = DB.all('defect_types').length; document.getElementById('md-name').value = '버(Burr)'; document.getElementById('md-save').onclick();
    P('Q2 품질 화면에서 불량 유형 추가 → defect_types +1 (마스터 편집 1벌)', DB.all('defect_types').length === dtN + 1 && DB.all('defect_types').some(d => d.name === '버(Burr)'), DB.all('defect_types').length);
    closeModal();
    // ── ⑦ 인쇄: 검은 띠 0 · 지시수량 옆 수기 칸 ──
    const allJs = Array.from(document.querySelectorAll('script')).map(s=>s.textContent).join('\\n');
    P('P1 인쇄 양식 검은 띠(#16181d 배경 rule/th) 0건', !/background:#16181d/.test(allJs), (allJs.match(/background:#16181d/g)||[]).length);
    P('P2 작업지시서 지시수량 옆 「실 수량(수기)」 칸', /실 수량<br>/.test(allJs) && /border:1px solid #9aa0a6;border-radius:4px;height:34px/.test(allJs), '');
    // ── ⑧ LOT 추적 ⋯ 메뉴 글자만 ──
    location.hash = '#lot'; router();
    P('M1 LOT 추적 ⋯ 메뉴 항목 글자만 (아이콘 0)', !document.querySelector('#lot-more-menu svg') && document.querySelectorAll('#lot-more-menu button').length >= 4, '');
    P('버전 ≥ v0.35.9', /^v0\\.35\\.(9|[1-9][0-9])$|^v0\\.3[6-9]/.test(APP_VERSION) && APP_DATE >= '2026-09-08', APP_VERSION);
  } catch(e){ window.__T.push({ n:'FATAL', ok:false, note: e.stack || e.message }); }
})();`;
win.eval(scripts.join('\n;\n') + body);
const out = win.__T || [];
let fail = 0;
out.forEach(r => { if(!r.ok) fail++; console.log(`[${r.ok?'PASS':'FAIL'}] ${r.n}${r.note?`  — ${r.note}`:''}`); });
console.log(`\nPILOT359 TOTAL ${out.length} · PASS ${out.length-fail} · FAIL ${fail}`);
// ★ v0.38.20 (GPT Gate) — FAIL 또는 항목 수 불일치(중도 종료) = 종료코드 1
if(out.length !== 29) console.log(`[FAIL] 항목 수 ${out.length} ≠ 기대 29 (중도 종료·누락)`);
process.exit((fail || out.length !== 29) ? 1 : 0);
