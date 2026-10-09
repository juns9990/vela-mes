// ★ v0.38.16(-r1 GPT Gate A~D) — 납기준수율 회귀 + 독립 산식 대조 (PART 1 · GPT 승인 · U1 취소 제외 · U2 현재 납기)
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));
const html = fs.readFileSync(path.resolve(__dirname, process.env.VELA_FILE || 'vela-mes-prototype.html'), 'utf8')   // GitHub 저장소에서는 VELA_FILE=index-dev.html;
const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
const win = dom.window; win.VELA_BACKEND = 'local';
win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
win.matchMedia = q => ({ matches:false, media:q, addListener(){}, removeListener(){}, addEventListener(){}, removeEventListener(){} });
win.eval(scripts.join('\n;\n') + `;window.__X = s => eval(s);`);
const X = s => win.__X(s);
const R = []; const P = (n, ok, note) => { R.push({ n, ok }); console.log(`[${ok?'PASS':'FAIL'}] ${n}${note?'  — '+note:''}`); };
X(`seedDemoData(); SESSION.loggedIn=true; SESSION.role='manager'; SESSION.user='검증'; SESSION.userId='U-T';`);
const L0 = X(`JSON.stringify(['orders','shipments','production_lots','records'].map(c => DB.allRaw(c).length))`);

// 손 계산 시나리오 (2020-01 — 데모 데이터와 분리)
X(`(()=>{
  const T = (d, h) => new Date(d + 'T' + (h||'10') + ':00:00').getTime();
  const O = (id, qty, due, cust, item) => DB.set('orders', id, { id, cust, itemCode:item, qty, price:100, due, created:T('2019-12-01') });
  const S = (id, so, qty, d) => DB.set('shipments', id, { id, so_id:so, itemCode:'HF-T', plot_no:'', qty, ts:T(d), by:'검증' });
  O('T1',100,'2020-01-15','CT1','HF-A'); S('S1','T1',100,'2020-01-14');
  O('T2',100,'2020-01-15','CT1','HF-A'); S('S2a','T2',60,'2020-01-15'); S('S2b','T2',40,'2020-01-17');
  O('T3',100,'2020-01-15','CT2','HF-B'); S('S3','T3',70,'2020-01-10'); DB.correct('orders','T3',{ close:{ kind:'short', at:T('2020-01-20'), by:'검증', reason:'협의', shipped_at:70, unshipped_at:30 } }, '부분종결: 협의');
  O('T4',50,'2020-01-15','CT2','HF-B'); DB.correct('orders','T4',{ close:{ kind:'cancel', at:T('2020-01-05'), by:'검증', reason:'취소', shipped_at:0, unshipped_at:50 } }, '수주취소');
  O('T5',50,'2020-01-15','CT2','HF-B'); DB.correct('orders','T5',{ close:{ kind:'cancel', at:T('2020-01-25'), by:'검증', reason:'취소', shipped_at:0, unshipped_at:50 } }, '수주취소');
  O('T6',80,'2020-01-15','CT1','HF-B'); S('S6','T6',80,'2020-01-12'); voidDoc('shipments','S6','검증');
  O('T7',100,'2020-01-10','CT1','HF-A'); S('S7','T7',100,'2020-01-13'); DB.correct('orders','T7',{ due:'2020-01-15' }, '납기 변경: 고객 협의');
  O('T8',10,'','CT1','HF-A');
  O('T9',10,'2099-01-15','CT1','HF-A');
  O('T10',100,'2020-01-20','CT2','HF-A'); S('S10a','T10',30,'2020-01-05'); S('S10b','T10',30,'2020-01-12'); S('S10c','T10',40,'2020-01-20');
})()`);
const r = X(`(()=>{ const r = calcOnTime({ from:'2020-01-01', to:'2020-01-31' }); return { t:r.total, ex:r.excluded, rows:r.rows.map(x => [x.id, x.ok, x.onTime, x.status, x.delayDays, x.dueChanged.length]), cust:r.byCust.map(g => [g.key, g.lines, g.okLines]), item:r.byItem.map(g => [g.key, g.lines, g.okLines]) }; })()`);
const row = id => r.rows.find(x => x[0] === id);
P('1 평가 대상 6건 (T1·T2·T3·T6·T7·T10) · 준수 3건 → 건수 기준 50.0%', r.t.lines === 6 && r.t.okLines === 3 && Math.abs(r.t.rateLines - 50) < 1e-9, JSON.stringify(r.t));
P('2 수량 기준 = 430 / 580 = 74.14%', r.t.qty === 580 && r.t.onTime === 430 && Math.abs(r.t.rateQty - 430/580*100) < 1e-9, JSON.stringify(r.t));
P('3 분할 출하 30+30+40 모두 납기일 이내(당일 포함) → 준수 (T10)', row('T10')[1] === true && row('T10')[2] === 100);
P('4 납기 당일 60 + 이틀 뒤 40 → 지연 출하 · 납기 내 60 · 지연 2일 (T2)', row('T2')[1] === false && row('T2')[2] === 60 && row('T2')[3] === '지연 출하' && row('T2')[4] === 2, JSON.stringify(row('T2')));
P('5 부분종결 70/100 → 미준수 · 수량 70 · 분모 100 유지 (T3)', row('T3')[1] === false && row('T3')[2] === 70 && row('T3')[3] === '부분종결', JSON.stringify(row('T3')));
P('6 수주취소(출하 0) 2건 집계 제외 · 그중 납기 후 취소 1건 별도 (U1)', r.ex.cancel === 2 && r.ex.lateCancel === 1 && !row('T4') && !row('T5'), JSON.stringify(r.ex));
P('7 취소(void)된 출하는 제외 → 미출하 (T6)', row('T6')[1] === false && row('T6')[2] === 0 && row('T6')[3] === '미출하', JSON.stringify(row('T6')));
P('8 납기 변경 → 현재 납기(1/15) 기준 준수 · 변경 이력 1건 보존 (T7 · U2)', row('T7')[1] === true && row('T7')[5] === 1, JSON.stringify(row('T7')));
P('9 납기 미등록·미도래 제외 + 별도 카운트 (T8·T9)', !row('T8') && !row('T9') && r.ex.noDue >= 1 && r.ex.notDue === 0, JSON.stringify(r.ex));
const r9 = X(`calcOnTime({ from:'2099-01-01', to:'2099-12-31' }).excluded.notDue`);
P('9b 미래 납기 = 미도래 카운트', r9 === 1);
P('10 고객사별·품목별 집계 일치 (CT1 4건 준수 2 · CT2 2건 준수 1 · HF-A 4건 준수 3 · HF-B 2건 준수 0)',
  JSON.stringify(r.cust.sort()) === JSON.stringify([['CT1',4,2],['CT2',2,1]]) && JSON.stringify(r.item.sort()) === JSON.stringify([['HF-A',4,3],['HF-B',2,0]]), JSON.stringify({cust:r.cust, item:r.item}));

// 독립 산식 대조 — 데모 데이터 · 최근 6개월 각 월 + 전체
const indep = X(`(()=>{
  const out = [];
  const today = new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  const months = []; for(let i=0;i<6;i++){ const d = new Date(); d.setDate(1); d.setMonth(d.getMonth()-i); months.push(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')); }
  const windows = months.map(m => [m+'-01', m+'-31']).concat([['2000-01-01', today]]);
  windows.forEach(([f,t]) => {
    // 독립 구현: 원시 컬렉션을 직접 읽어 계산 (calcOnTime·soCloseInfo·_dueEndTs 미사용)
    let lines=0, ok=0, q=0, on=0;
    Object.values(DB.allRaw('orders')).forEach(o => {
      if(o.void || !o.due || o.due < f || o.due > t || o.due > today) return;
      const [y,m,d] = o.due.split('-').map(Number); const end = Date.UTC(y,m-1,d,23,59,59,999) - 9*3600000;   // KST 23:59:59.999
      let s = 0, all = 0; Object.values(DB.allRaw('shipments')).forEach(sh => { if(!sh.void && sh.so_id === o.id){ all += (+sh.qty||0); if(sh.ts <= end) s += (+sh.qty||0); } });
      if(o.close && o.close.kind === 'cancel' && all === 0) return;
      if(o.due === today) return;   // v0.38.18 — 오늘 납기 = 출하 여부 무관 공식 집계 제외
      lines++; q += (+o.qty||0); on += Math.min(+o.qty||0, s); if(s >= (+o.qty||0) && (+o.qty||0) > 0) ok++;
    });
    const r = calcOnTime({ from:f, to:t });
    out.push({ w:f+'~'+t, same: r.total.lines===lines && r.total.okLines===ok && r.total.qty===q && r.total.onTime===on, a:[r.total.lines,r.total.okLines,r.total.qty,r.total.onTime], b:[lines,ok,q,on] });
  });
  return out;
})()`);
P('11 독립 산식 대조 — 최근 6개월 월별 + 전체 기간: 평가 건수·준수 건수·수주량·납기 내 출하량 모두 일치', indep.every(x => x.same) && indep.some(x => x.a[0] > 0), JSON.stringify(indep.filter(x => !x.same).concat(indep.slice(-1))));

// ── GPT Gate A — 납기 당일 경계 (KST 고정 · today 주입)
X(`(()=>{
  const K = s => Date.parse(s + '+09:00');
  const O = (id, qty, due, extra) => DB.set('orders', id, { id, cust:'CG', itemCode:'HF-G', qty, price:100, due, created:K('2021-03-01T09:00:00'), ...(extra||{}) });
  const S = (id, so, qty, at) => DB.set('shipments', id, { id, so_id:so, itemCode:'HF-G', plot_no:'', qty, ts:K(at), by:'검증' });
  O('G1',100,'2021-03-10');                                        // 오늘 납기 · 아직 미출하
  O('G2',100,'2021-03-10'); S('SG2','G2',100,'2021-03-10T15:00:00');  // 오늘 오후 전량 출하
  O('G3',100,'2021-03-10'); S('SG3','G3',100,'2021-03-10T23:30:00');  // 납기일 23:30 KST
  O('G4',100,'2021-03-09'); S('SG4','G4',100,'2021-03-10T00:10:00');  // 납기 다음날 00:10 KST
})()`);
const g = X(`(()=>{ const a = calcOnTime({ from:'2021-03-01', to:'2021-03-31', cust:'CG', today:'2021-03-10' }); const b = calcOnTime({ from:'2021-03-01', to:'2021-03-31', cust:'CG', today:'2021-03-11' });
  const st = (r, id) => (r.rows.find(x => x.id === id) || {}).status || null;
  return { aEx:a.excluded.dueToday, aDone:a.excluded.dueTodayDone, aLines:a.total.lines, a1:st(a,'G1'), a2:st(a,'G2'), a3:st(a,'G3'), a4:st(a,'G4'), b1:st(b,'G1'), b2:st(b,'G2'), b3:st(b,'G3'), bEx:b.excluded.dueToday, bLines:b.total.lines }; })()`);
P('A1 오늘 납기 · 오전 미출하 → 공식 집계 제외 (오늘 납기 진행 중)', g.a1 === null && g.aEx === 3, JSON.stringify(g));
P('A2 오늘 납기 · 당일 전량 출하(15:00·23:30 KST)도 당일엔 공식 집계 제외 · 「오늘 정시 출하 완료」 2건으로 별도 표시 (결정 2)', g.a2 === null && g.a3 === null && g.aDone === 2 && g.aLines === 1, JSON.stringify(g));
P('A3 다음 날 자동 전환: 오늘 납기 3건 전부 공식 평가 (G1 미출하 · G2·G3 준수) · 다음날 00:10 KST 출하 = 지연 출하', g.b1 === '미출하' && g.b2 === '준수' && g.b3 === '준수' && g.bEx === 0 && g.bLines === 4 && g.a4 === '지연 출하', JSON.stringify(g));
// GPT 예시 — 과거 10건 중 8 준수 + 오늘 납기 10건 중 2건 아침 출하 → 공식 = 80.0% (83.3% 아님)
X(`(()=>{ const K = s => Date.parse(s + '+09:00');
  for(let i=0;i<10;i++){ DB.set('orders','E'+i,{ id:'E'+i, cust:'CE', itemCode:'HF-G', qty:10, price:1, due:'2021-05-09' }); if(i<8) DB.set('shipments','SE'+i,{ id:'SE'+i, so_id:'E'+i, itemCode:'HF-G', qty:10, ts:K('2021-05-08T10:00:00'), by:'검증' }); }
  for(let i=0;i<10;i++){ DB.set('orders','F'+i,{ id:'F'+i, cust:'CE', itemCode:'HF-G', qty:10, price:1, due:'2021-05-10' }); if(i<2) DB.set('shipments','SF'+i,{ id:'SF'+i, so_id:'F'+i, itemCode:'HF-G', qty:10, ts:K('2021-05-10T09:00:00'), by:'검증' }); }
})()`);
const e = X(`(()=>{ const r = calcOnTime({ from:'2021-05-01', to:'2021-05-31', cust:'CE', today:'2021-05-10' }); return { rate:r.total.rateLines, lines:r.total.lines, today:r.excluded.dueToday, done:r.excluded.dueTodayDone }; })()`);
P('A4 선택 편향 방지 (GPT 예시): 과거 8/10 + 오늘 2/10 완료 → 공식 80.0% (83.3% 아님) · 오늘 10건 중 완료 2 별도', Math.abs(e.rate - 80) < 1e-9 && e.lines === 10 && e.today === 10 && e.done === 2, JSON.stringify(e));
// ── GPT Gate B — 취소 유형 = close.kind==='cancel' ∧ 출하 0 만 제외
X(`(()=>{
  const K = s => Date.parse(s + '+09:00');
  DB.set('orders','B1',{ id:'B1', cust:'CB', itemCode:'HF-G', qty:50, price:1, due:'2021-04-10' });                     // 출하 0 · 열린 수주
  DB.set('orders','B2',{ id:'B2', cust:'CB', itemCode:'HF-G', qty:50, price:1, due:'2021-04-10' });
  DB.correct('orders','B2',{ close:{ kind:'short', at:K('2021-04-12T09:00:00'), by:'검증', reason:'협의', shipped_at:0, unshipped_at:50 } }, '부분종결');   // 부분종결인데 출하 0
  DB.set('orders','B3',{ id:'B3', cust:'CB', itemCode:'HF-G', qty:50, price:1, due:'2021-04-10' });
  DB.set('shipments','SB3',{ id:'SB3', so_id:'B3', itemCode:'HF-G', qty:20, ts:K('2021-04-09T10:00:00'), by:'검증' });
  DB.correct('orders','B3',{ close:{ kind:'cancel', at:K('2021-04-12T09:00:00'), by:'검증', reason:'이상 데이터', shipped_at:20, unshipped_at:30 } }, '취소(출하 있음 — 이상 데이터)');
  DB.set('orders','B4',{ id:'B4', cust:'CB', itemCode:'HF-G', qty:50, price:1, due:'2021-04-10' });
  DB.correct('orders','B4',{ close:{ kind:'cancel', at:K('2021-04-08T09:00:00'), by:'검증', reason:'취소', shipped_at:0, unshipped_at:50 } }, '수주취소');
})()`);
const b = X(`(()=>{ const r = calcOnTime({ from:'2021-04-01', to:'2021-04-30', cust:'CB', today:'2021-05-01' }); const st = id => (r.rows.find(x => x.id === id) || {}).status || null;
  return { b1:st('B1'), b2:st('B2'), b3:st('B3'), b4:st('B4'), cancel:r.excluded.cancel, lines:r.total.lines }; })()`);
P('B1 출하 0 인 열린 수주 = 미출하로 평가 (제외 아님)', b.b1 === '미출하', JSON.stringify(b));
P('B2 출하 0 인 부분종결 = 부분종결로 평가 (제외 아님)', b.b2 === '부분종결', JSON.stringify(b));
P('B3 취소 표시여도 출하가 있으면 제외하지 않음 (부분종결로 평가) · B4 출하 0 취소만 제외', b.b3 === '부분종결' && b.b4 === null && b.cancel === 1 && b.lines === 3, JSON.stringify(b));
// ── GPT Gate C — 출하일 수정 감사 이력 + 납기준수 영향 기록
const c = X(`(()=>{
  const K = s => Date.parse(s + '+09:00'); const now = new Date(); const ym = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
  const due = ym + '-01'; const dueTs = K(due + 'T10:00:00');
  DB.set('orders','C1',{ id:'C1', cust:'CC', itemCode:'HF-G', qty:10, price:1, due });
  DB.set('shipments','SC1',{ id:'SC1', so_id:'C1', itemCode:'HF-G', plot_no:'', qty:10, ts:dueTs, by:'검증' });
  const im = otdShipDateImpact('SC1', _tsWithDate(dueTs, ym + '-02'));
  let noReason = false; try { docEditSave('shipments','SC1',{ date: ym + '-02' }, ''); } catch(e){ noReason = /사유/.test(e.message); }
  const r = docEditSave('shipments','SC1',{ date: ym + '-02' }, '출하일 오입력 정정');
  const sh = DB.get('shipments','SC1'); const lg = (sh._logs||[]).filter(l => l.op === 'correct').pop() || {};
  const after = calcOnTime({ from: due, to: due, cust:'CC' }).rows.find(x => x.id === 'C1');
  return { im, noReason, ok: r.ok, prev: lg.prev && fmtDate(lg.prev.ts), next: lg.next && fmtDate(lg.next.ts), by: lg.by, at: !!lg.ts, reason: lg.reason, status: after && after.status };
})()`);
P('C1 출하일 수정 → 영향 미리 확인(준수 → 지연 출하) · 사유 없으면 저장 거부', c.im && c.im.changed && /^준수/.test(c.im.before) && /^지연 출하/.test(c.im.after) && c.noReason, JSON.stringify(c.im));
P('C2 감사 이력: 수정 전·후 출하일 · 수정자 · 시각 · 사유 + 「납기준수 영향」 문구 · 재계산 결과 = 지연 출하', c.ok && c.prev && c.next && c.prev !== c.next && c.by && c.at && /출하일 오입력 정정/.test(c.reason) && /납기준수 영향: C1 준수 \(납기 내 10\/10\) → 지연 출하 \(납기 내 0\/10\)/.test(c.reason) && c.status === '지연 출하', JSON.stringify(c));
const c4 = X(`(()=>{
  const K = s => Date.parse(s + '+09:00'); const now = new Date(); const ym = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
  const due = ym + '-03';
  DB.set('orders','C4',{ id:'C4', cust:'CC', itemCode:'HF-G', qty:100, price:1, due });
  DB.set('shipments','SC4a',{ id:'SC4a', so_id:'C4', itemCode:'HF-G', plot_no:'', qty:20, ts:K(ym + '-01T10:00:00'), by:'검증' });
  DB.set('shipments','SC4b',{ id:'SC4b', so_id:'C4', itemCode:'HF-G', plot_no:'', qty:20, ts:K(ym + '-02T10:00:00'), by:'검증' });
  const im = otdShipDateImpact('SC4b', K(ym + '-05T10:00:00'));
  const same = otdShipDateImpact('SC4b', K(ym + '-01T15:00:00'));
  return { im, same };
})()`);
P('C4 (GPT 소스 검토 P1) 일부 출하 → 일부 출하라도 납기 내 수량 40 → 20 이면 영향 있음 · 납기 안에서 날짜만 바뀌면 영향 없음', c4.im && c4.im.changed && /일부 출하 \(납기 내 40\/100\)/.test(c4.im.before) && /일부 출하 \(납기 내 20\/100\)/.test(c4.im.after) && c4.same && c4.same.changed === false, JSON.stringify(c4));
const snap = X(`(()=>{ const src = Object.keys(DB.get('settings','report_snapshot_test')||{}); const rp = DB.all('reports'); return { reports: rp.length, otdInSnapshot: rp.some(x => JSON.stringify(x.data||{}).includes('rateLines')) }; })()`);
P('C3 저장된 보고서 스냅샷에 납기준수 값 없음 (PART 2 보고서 블록에서 스냅샷 저장 예정 → 지금 수정이 과거 보고서를 바꾸지 않음)', snap.otdInSnapshot === false, JSON.stringify(snap));
// ── GPT Gate D — 수량 단위 혼합 방지
X(`(()=>{
  const K = s => Date.parse(s + '+09:00');
  DB.set('orders','D1',{ id:'D1', cust:'CD', itemCode:'HF-G', qty:100, price:1, due:'2021-06-10' });
  DB.set('shipments','SD1',{ id:'SD1', so_id:'D1', itemCode:'HF-G', qty:100, ts:K('2021-06-09T10:00:00'), by:'검증' });
  DB.set('orders','D2',{ id:'D2', cust:'CD', itemCode:'HF-K', qty:500, unit:'kg', price:1, due:'2021-06-10' });
  DB.set('shipments','SD2',{ id:'SD2', so_id:'D2', itemCode:'HF-K', qty:200, ts:K('2021-06-09T10:00:00'), by:'검증' });
})()`);
const d = X(`(()=>{ const r = calcOnTime({ from:'2021-06-01', to:'2021-06-30', cust:'CD' }); const one = calcOnTime({ from:'2021-06-01', to:'2021-06-30', cust:'CD', itemCode:'HF-K' });
  return { mixed:r.total.mixedUnit, rq:r.total.rateQty, q:r.total.qty, rl:r.total.rateLines, itemK: r.byItem.find(g => g.key==='HF-K'), one: [one.total.mixedUnit, one.total.rateQty, one.total.unit], dflt: soUnit({ itemCode:'HF-G' }) }; })()`);
P('D1 ea + kg 혼합 → 수량 기준 미산출(합산 안 함) · 건수 기준은 정상(50%)', d.mixed === true && d.rq === null && d.q === null && d.rl === 50, JSON.stringify(d));
P('D2 같은 단위 안(품목 HF-K · kg)에서는 수량 기준 산출 40% · 단위 기본값 ea', d.one[0] === false && Math.abs(d.one[1] - 40) < 1e-9 && d.one[2] === 'kg' && d.itemK && d.itemK.rateQty === 40 && d.dflt === 'ea', JSON.stringify(d));

// UI · 원장 무변
const ui = X(`(()=>{ window._shTab = 'list'; location.hash = '#ship'; VIEWS.ship(); document.getElementById('btn-ship-otd').onclick();
  const v = document.getElementById('view').textContent; const ok1 = /납기 실적/.test(v) && /납기준수율 \\(건수 · 주지표\\)/.test(v) && /출하일/.test(v) && /오늘 납기 · 진행 중/.test(v) && /정시 출하 완료 \\d+ · 미완료 \\d+/.test(v);
  window._otdFilter = { from:'2020-01-01', to:'2020-01-31', cust:'', item:'' }; router(); const v2 = document.getElementById('view').textContent;
  const ok2 = /50\\.0%/.test(v2) && /74\\.1%/.test(v2) && /T2/.test(v2) && !/>T1</.test(document.getElementById('view').innerHTML) && /수주취소 2 \\(납기 후 취소 1\\)/.test(v2);
  document.getElementById('otd-back').onclick(); const ok3 = !!document.getElementById('btn-ship-new');
  window._otdFilter = null; return { ok1, ok2, ok3 }; })()`);
P('12 UI: 출하 [납기 실적] → KPI(건수 50.0% · 수량 74.1%) · 지연 목록(T2 포함 · 준수 T1 제외) · 제외 표시 · [출하 목록] 복귀', ui.ok1 && ui.ok2 && ui.ok3, JSON.stringify(ui));
const L1 = X(`JSON.stringify(['orders','shipments','production_lots','records'].map(c => DB.allRaw(c).length))`);
const before = JSON.parse(L0), after = JSON.parse(L1);
P('13 조회 기능 — 원장 쓰기 0 (테스트 준비분 외 문서 수 변화 없음: 수주 +42 · 출하 +28)', after[0] - before[0] === 42 && after[1] - before[1] === 28 && after[2] === before[2] && after[3] === before[3], L0 + ' → ' + L1);
const fail = R.filter(x => !x.ok).length;
console.log(`\nOTD TOTAL ${R.length} · PASS ${R.length - fail} · FAIL ${fail}`);
process.exit(fail ? 1 : 0);
