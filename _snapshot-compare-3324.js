// P1-1 증거 — v0.33.2(승인 기준) vs v0.33.4(현재) 동일 시드 원장·집계 스냅샷 비교
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require(path.join(__dirname, 'node_modules/jsdom'));

function snapshot(file){
  const html = fs.readFileSync(path.join(__dirname, file), 'utf8');
  const stripped = html.replace(/<script\s+src=[^>]*><\/script>/g, '').replace(/<link\s+rel="manifest"[^>]*>/g, '');
  const dom = new JSDOM(stripped, { runScripts:'outside-only', pretendToBeVisual:true, virtualConsole:new VirtualConsole(), url:'http://localhost/' });
  const win = dom.window;
  win.XLSX = { utils:{book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{}}, writeFile:()=>{} };
  win.Html5Qrcode = function(){}; win.qrcode = function(){ return { addData(){}, make(){}, createSvgTag(){return ''}, createDataURL(){return ''} }; };
  const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  const body = `
;(function(){
  seedDemoData();
  const ships = DB.all('shipments');
  const s = calcSales(ships);
  const st = calcSoFulfillment();
  const recs = DB.all('records');
  window.__SNAP = {
    ver: APP_VERSION,
    counts: {
      orders: DB.all('orders').length, work_orders: DB.all('work_orders').length,
      production_lots: DB.all('production_lots').length, records: recs.length,
      shipments: ships.length, purchase_orders: DB.all('purchase_orders').length,
      nc_records: DB.all('nc_records').length, quotes: DB.all('quotes').length,
      acct_receipts: DB.all('acct_receipts').length, acct_expenses: DB.all('acct_expenses').length,
      items: DB.all('items').length, vendors: DB.all('vendors').length, workers: DB.all('workers').length,
    },
    good_total: recs.reduce((a,r)=>a+(+r.good||0),0),
    defect_total: recs.reduce((a,r)=>a+(+r.defect||0),0),
    ship_qty_total: ships.reduce((a,r)=>a+(+r.qty||0),0),
    monthSales: Math.round(s.monthSales), yearSales: Math.round(s.yearSales),
    so_total: Math.round(st.total), so_done: Math.round(st.done), so_prog: Math.round(st.prog), so_wait: Math.round(st.wait),
    so_identity_err: Math.round(st.total) - (Math.round(st.done)+Math.round(st.prog)+Math.round(st.wait)),
    sample: {
      h01_bal_ship: LOT.balance('PL-2026-H01','SHIP'),
      h01_qty: DB.get('production_lots','PL-2026-H01')?.qty_initial ?? null,
      first_order: DB.all('orders').map(o=>o.id).sort()[0] || null,
      first_ship: ships.map(x=>x.id).sort()[0] || null,
    },
  };
})();`;
  win.eval(scripts.join('\n;\n') + body);
  return win.__SNAP;
}

const a = snapshot('backup/vela-mes-v0.33.2-백업.html');
const b = snapshot('vela-mes-prototype.html');
const flat = o => { const out={}; (function w(p,v){ if(v && typeof v==='object'){ Object.entries(v).forEach(([k,x])=>w(p?p+'.'+k:k,x)); } else out[p]=v; })('',o); return out; };
const fa = flat(a), fb = flat(b);
let diff = 0;
console.log('key | v0.33.2 | v0.33.4 | same');
Object.keys(fa).forEach(k => {
  if(k==='ver') return;
  const same = String(fa[k]) === String(fb[k]);
  if(!same) diff++;
  console.log(`${k} | ${fa[k]} | ${fb[k]} | ${same?'✓':'✗ MISMATCH'}`);
});
console.log(`\nRESULT: ${diff===0?'IDENTICAL — 원장·집계 무변 증명':'DIFF '+diff+'건'}`);
process.exit(diff === 0 ? 0 : 1);   // ★ v0.38.20 — 불일치 = 종료코드 1
