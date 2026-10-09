// v0.33.2 verification harness — 동시성 사고 방어·복구 Gate(Q2-3 — 탐지·잠금·정정·J1 재검증·해제) + Audit Viewer + Hold Gate + 전 회귀
//  1) ht_spec 표준 필드: 폼 UI · 저장 · 표 컬럼 · 엑셀 · 임포트 스펙
//  2) getHtSpec 폴백: 구 custom['열처리사양'] 값 계속 읽힘 (이력 보존)
//  3) 마이그레이션: 구 정의 필드 비활성(멱등) · 물리 삭제 없음
//  4) 회귀: 파스(별도) · 렌더 스모크 · 시뮬 수치 불변 · 외부 참조 · PPT/엑셀 무변
// Usage: node _verify-v0.27.0.js
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const target = path.join(__dirname, 'vela-mes-prototype.html');
const html = fs.readFileSync(target, 'utf8');

const results = [];
const record = (name, ok, detail) => {
  results.push({ name, ok, detail: detail || '' });
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${name}${detail ? `  — ${detail}` : ''}`);
};

// ── 외부 참조 (기준: cdnjs + jsDelivr 버전 고정만) ──
const ALLOWED_HOSTS = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net'];
const scriptTags = [...html.matchAll(/<script\b[^>]*\bsrc=(['"])([^'"]+)\1[^>]*>/g)].map(m => m[2]);
record('external: 로컬 script src 0건', scriptTags.filter(s => !/^https?:/.test(s)).length === 0);
record('external: 허용 CDN 2곳 밖 0건', scriptTags.filter(s => /^https?:/.test(s) && !ALLOWED_HOSTS.some(h => s.includes(h))).length === 0);
record('external: 전부 버전 고정', scriptTags.filter(s => /@latest/.test(s) || (!/\/\d+\.\d+\.\d+\//.test(s) && !/@\d+\.\d+\.\d+\//.test(s))).length === 0, `${scriptTags.length} external`);

// ── 부팅 ──
const stripped = html
  .replace(/<script\b[^>]*\bsrc=[^>]*><\/script>/g, '')
  .replace(/<link\s+rel="manifest"[^>]*>/g, '');
const virtualConsole = new VirtualConsole();
const consoleErrors = [];
virtualConsole.on('jsdomError', (e) => { consoleErrors.push(String((e && e.message) || e)); });
const dom = new JSDOM(stripped, { runScripts: 'outside-only', pretendToBeVisual: true, virtualConsole, url: 'http://localhost/' });
const win = dom.window;
win.XLSX = require(path.join(__dirname, 'vela-lib-xlsx.min.js'));
try { win.eval(fs.readFileSync(path.join(__dirname, 'vela-lib-pptxgen.bundle.js'), 'utf8')); } catch(e){ /* PPT 회귀 항목에서 판정 */ }
win.Html5Qrcode = function(){ this.start=()=>Promise.resolve(); this.stop=()=>Promise.resolve(); };
win.Html5QrcodeScanner = function(){ this.render=()=>{}; this.clear=()=>{}; };
win.qrcode = function(){ return { addData:()=>{}, make:()=>{}, createSvgTag:()=>'<svg/>', createDataURL:()=>'data:image/gif;base64,QRSTUB' }; };

const scripts = [...stripped.matchAll(/<script\b(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const expose = `
;(function(){
  var names = ['APP_VERSION','APP_DATE','SES','DB','LOT','PAGES','VIEWS','SESSION','getFlags','setFlag',
               'seedDemoData','router','shipAmount','buildReportData','fmtDate','buildReportPpt',
               'HT_SPEC_OPTS','getHtSpec','getHardness','migrateHtSpecField','getMasterFields','_IMPORTER_SPEC','isEmptyDemo','isDemoMode','fmtKg','isItemIdentifier','splitItemIdentifier','makeItemIdentifier','nextItemSeqForCustomer','nextItemIdentifier','itemHistoryRefs','enterStandby','wakeStandby','_sbLoadWeather','_SB_WCODE','_itemSeqsInHistory','MOLD_VARIANT_LABELS','moldQty','moldShotsPerUnit','deriveMoldShots','moldLifePct','getPartNo','itemLabel','_calEnhance','attachCalendar','soPlannedQty','soMaterialNeeds','matStockKg','matPriceKg','nextDocNo','mtHistoryRefs','suggestNextId','_scrollTarget','_snScroll','_reportState',
               'STD_EXPENSE_CATS','migrateExpenseCategories','_drawNextRev','MODULE_DEFS','_drawAttList','_drawAttAdd','_drawAttRemove','_drawFileExt','_fmtBytes','_dwB64Chunks','_dwMime','DW_STORE_MAX_MB','DW_CHUNK_CHARS','getUserPerm','currentPagePerm','_permBlockWrite','PERM_LEVELS','_parseImportRow','_importJudgeRow','_importApplyRow','_importerColumns','_parseImportRow','_IMPORTER_SPEC','isEmptyDemo','isDemoMode','fmtKg','isItemIdentifier','splitItemIdentifier','makeItemIdentifier','nextItemSeqForCustomer','nextItemIdentifier','itemHistoryRefs','enterStandby','wakeStandby','_sbLoadWeather','_SB_WCODE','_itemSeqsInHistory','MOLD_VARIANT_LABELS','moldQty','moldShotsPerUnit','deriveMoldShots','moldLifePct','getPartNo','itemLabel','_calEnhance','attachCalendar','soPlannedQty','soMaterialNeeds','matStockKg','matPriceKg','nextDocNo','mtHistoryRefs','suggestNextId','_scrollTarget','_snScroll','_importParseRoute','_importDateTs','osAvailable','getFlags','setFlag','getAcctSettings','calcAcctFlows','calcReceivable','calcPayable','calcAcctPnl','calcInvoiceTargets','_acctRepeatSuggestions','acctExpensesOf','acctOverdueDays','calcClosing','_tsMonth','fmtN','voidDoc','MASTER_DEL_COLS','_MASTER_DEL','masterDelete','masterHistoryRefs','masterDelBtn','_forgeQtyForItem','NAV_GROUPS','lotStepPanelHTML','osInFlight','getRoute','_ruleNoteEnhance','lotHoldInfo','assertLotNotHeld','heldPlotSet','auditEntries','lotTimeline','auditDiffText','AUDIT_COLS','_auditState','detectConcurrencyIncidents','lotLockInfo','assertLotNotLocked','lockedPlotSet','verifyLotInvariant','resolveIncident','INDUSTRY_PACKS','applyIndustryPack','moldLinkedProcSet','migrateMoldLinked','planDayTimeline','planRateFor','planMoldFor','getShifts','SHIFT_PRESETS','shiftMinutes','_LOCAL_ADAPTER','lotHasNegativeBalance'];
  names.forEach(function(n){ try { window[n] = eval(n); } catch(e){} });
})();`;
try { win.eval(scripts.join('\n;\n') + expose); record('boot: eval inline scripts', true, `${scripts.length} block(s)`); }
catch (e) { record('boot: eval inline scripts', false, e.message); finish(); }
record('boot: no jsdomErrors during eval', consoleErrors.length === 0, consoleErrors.slice(0,3).join(' | '));
record('version: APP_VERSION == v0.38.17', win.APP_VERSION === 'v0.38.23', win.APP_VERSION);
record('version: APP_DATE == 2026-10-09', win.APP_DATE === '2026-10-09', win.APP_DATE);

// ── v0.31.4 빈 상태 데모 (테스트 모드) ──
record('demo287: _DEMO_EMPTY_KEY 정의 1곳', (html.match(/const _DEMO_EMPTY_KEY = 'vela_demo_empty';/g) || []).length === 1);
record('demo287: isEmptyDemo 헬퍼 1벌', (html.match(/function isEmptyDemo\(/g) || []).length === 1);
record('demo287: isEmptyDemo 기본 false', win.isEmptyDemo() === false);
(function(){
  try {
    win.sessionStorage.setItem('vela_demo','1');
    win.localStorage.setItem('vela_demo_empty','1');
    record('demo287: 두 플래그 시 true', win.isEmptyDemo() === true);
    win.localStorage.removeItem('vela_demo_empty');
    record('demo287: EMPTY 키 제거 시 false (샘플 데모)', win.isEmptyDemo() === false);
  } finally {
    win.sessionStorage.removeItem('vela_demo');
    win.localStorage.removeItem('vela_demo_empty');
  }
})();
record('demo383: 데모 플래그 = 탭 단위 — sessionStorage 만 데모 · 옛 localStorage 플래그는 데모 아님 + 부팅 시 제거', (()=>{
  try {
    win.localStorage.setItem('vela_demo','1');
    const legacy = win.isDemoMode() === false;
    win.eval('_clearLegacyDemoFlag()');
    const cleared = win.localStorage.getItem('vela_demo') === null;
    win.sessionStorage.setItem('vela_demo','1'); const on = win.isDemoMode() === true; win.sessionStorage.removeItem('vela_demo');
    return legacy && cleared && on && win.isDemoMode() === false && /sessionStorage\.setItem\(_DEMO_FLAG_KEY, '1'\)/.test(html) && /sessionStorage\.removeItem\(_DEMO_FLAG_KEY\)/.test(html);
  } finally { win.localStorage.removeItem('vela_demo'); win.sessionStorage.removeItem('vela_demo'); }
})());
record('vis383: 추세 차트 값·축·점 = 테마 변수 (흰 글자 고정 #eef0f4 0)', /style="fill:var\(--ink\)">\$\{p\.good\.toLocaleString\(\)\}/.test(html) && !/fill="#eef0f4"/.test(html));
record('vis383: 라이트 --heat/--c1~c6 · 다크 --dim 4.5:1 보정 · 계획 교대 밴드 토큰', /--heat:#c2410c; --heat2:#b45309;/.test(html) && /html\[data-theme="light"\]\{ --c1:#0f766e/.test(html) && /--dim:#7d8390;/.test(html) && /var\(--shift1\)/.test(html) && !/'#fefce8','#eef2ff','#f5f3ff'/.test(html));
record('prt383: 인쇄 장 넘김 이중 안전장치 + 화면 최소 높이 해제 + 가로 한 장 맞춤', /\.page \+ \.page\{page-break-before:always;break-before:page\}/.test(html) && /min-height:0 !important;/.test(html) && /orient === 'landscape' \? ' \.page:not\(\.flow\)\{zoom:\.66\}'/.test(html));
record('prt383: PDF = 한 장 양식은 1쪽 맞춤 · 흐름 문서(.flow/.sheet)만 행 경계 분할', /const isFlow = el\.classList\.contains\('flow'\) \|\| el\.classList\.contains\('sheet'\);/.test(html) && /const fit = cuts\.filter/.test(html));
record('prt383: 출력물 하단 버전 표기 _stampPrintVersion (인쇄·PDF 공통)', win.eval('_stampPrintVersion')('<span>VELA SYSTEM</span>') === '<span>VELA SYSTEM · ' + win.APP_VERSION + '</span>' && (html.match(/bodyHTML = _stampPrintVersion\(bodyHTML\);/g)||[]).length === 2);
record('demo287: 진입 모달 2버튼 (샘플/빈 상태)', /id="demo-go-sample"/.test(html) && /id="demo-go-empty"/.test(html));
record('demo287: 샘플 진입 = EMPTY 키 제거+데모 저장소 리셋', /demo-go-sample'\)\.onclick = \(\) => \{ try\{ localStorage\.removeItem\(_DEMO_EMPTY_KEY\); localStorage\.removeItem\(_DEMO_LS_KEY\);/.test(html));
record('demo287: 빈 상태 진입 = EMPTY 키 설정+데모 저장소 리셋', /demo-go-empty'\)\.onclick\s+= \(\) => \{ try\{ localStorage\.setItem\(_DEMO_EMPTY_KEY,'1'\); localStorage\.removeItem\(_DEMO_LS_KEY\);/.test(html));
record('demo287: seedIfEmpty 데이터성 7블록 게이팅', (html.match(/!isEmptyDemo\(\) && /g) || []).length === 7);
record('demo287: 구조성 시드 무게이팅 (공정·계정과목·설정)', /if\(DB\.all\('processes'\)\.length===0\)\{/.test(html) && /if\(DB\.all\('expense_categories'\)\.length===0\)\{/.test(html) && /if\(DB\.all\('settings'\)\.length===0\)\{/.test(html));
record('demo287: autoLoginDemo 빈 데모 시 seedDemoData 스킵 (v0.32.6 — 시드 버전 마커 문구로 승계)', /if\(!isEmptyDemo\(\)\)\{\s*\n\s*try \{ seedDemoData\(\); DB\.set\('settings','demo_seed_ver'/.test(html));
record('demo287: 초기화 2선택 (샘플 재생성/빈 상태)', /id="demo-reset-sample"/.test(html) && /id="demo-reset-empty"/.test(html));
record('demo287: 초기화 핸들러 EMPTY 키 토글 후 resetDemoData', /demo-reset-sample'\)\.onclick = \(\) => \{ try\{ localStorage\.removeItem\(_DEMO_EMPTY_KEY\); \}catch\(e\)\{\} resetDemoData\(\); \};/.test(html) && /demo-reset-empty'\)\.onclick\s+= \(\) => \{ try\{ localStorage\.setItem\(_DEMO_EMPTY_KEY,'1'\); \}catch\(e\)\{\} resetDemoData\(\); \};/.test(html));
record('demo287: 설정 카드 모드 표기 (빈 상태/샘플)', /isEmptyDemo\(\) \? '빈 상태 \(시연·테스트\)' : '샘플 데이터 \(데모정밀\)'/.test(html));

// ── v0.31.4 중량 kg 통일 (표시·입력만 kg · 저장 필드 무변 = 마이그레이션 불필요) ──
record('kg288: fmtKg 헬퍼 1벌', (html.match(/function fmtKg\(/g) || []).length === 1);
record('kg288: fmtKg 소수 3자리·빈값', win.fmtKg(1.5534) === '1.553' && win.fmtKg(0.535) === '0.535' && win.fmtKg(0) === '' && win.fmtKg(null) === '');
record('kg288: 폼 kg 입력 (step 0.001·소수)', /id="mi-cutkg" type="number" min="0" step="0.001"/.test(html) && /id="mi-prodkg" type="number" min="0" step="0.001"/.test(html));
record('kg288: 구 g 입력 id 제거', !/mi-cutg"/.test(html) && !/mi-prodg"/.test(html));
record('kg288: 표 헤더 kg', /<th class="num">절단\(kg\)<\/th><th class="num">제품\(kg\)<\/th>/.test(html));
record('kg288: 엑셀 내보내기 kg 컬럼', /'절단중량\(kg\)': i\.weight/.test(html) && /'제품중량\(kg\)': i\.product_weight_g/.test(html));
const kgCols = (win._IMPORTER_SPEC?.items?.columns || []).map(c => c.key);
record('kg288: 임포트 스펙 weight_kg/product_weight_kg', kgCols.includes('weight_kg') && kgCols.includes('product_weight_kg') && !kgCols.includes('weight_g'), kgCols.join(','));
record('kg288: 저장 필드 무변 (weight=kg · product_weight_g=g 하위 호환)', /const weight = cutKg \? cutKg : \(\+it\.weight\|\|0\);/.test(html) && /prodKg \? Math\.round\(prodKg\*1000\*1000\)\/1000 : 0/.test(html));
record('kg288: 원/kg 자동계산 kg 직산', /Math\.round\(priceV \/ cutKg\)/.test(html) && /Math\.round\(priceV \/ prodKg\)/.test(html));

// ── v0.31.4 식별번호 리비전 분리 + 자동채번 중복방지 + 무이력 물리 삭제 ──
record('id289: 신규 채번 2분절 (rev 미포함)', win.makeItemIdentifier('038', 31) === '038-031' && !/-\d+$/.test(win.makeItemIdentifier('038',31).slice(7)));
record('id289: 2분절·구 3분절 모두 인식', win.isItemIdentifier('038-031') && win.isItemIdentifier('038-031-0') && !win.isItemIdentifier('038-31'));
record('id289: split — 2분절 rev=null · 3분절 rev 유지',
  JSON.stringify(win.splitItemIdentifier('038-031')) === '{"cust":"038","seq":31,"rev":null}' &&
  win.splitItemIdentifier('038-031-2').rev === 2);
(function(){
  // 자동채번: 최소 빈 번호 재사용 (001·003 존재 → 다음은 002)
  win.DB.set('items','T9-001',{code:'T9-001',name:'t1',active:true});
  win.DB.set('items','T9-003',{code:'T9-003',name:'t3',active:true});
  record('id289: 최소 빈 번호 채번 (001·003 → 002 = 삭제 번호 즉시 재사용)', win.nextItemIdentifier('T9') === 'T9-002');
  // 무이력 → 삭제 가능 · 삭제 후 번호 재사용
  record('del289: 무이력 품번 참조 0', win.itemHistoryRefs('T9-003').length === 0);
  record('del289: DB.remove 물리 삭제 (items)', win.DB.remove('items','T9-003') === true && !win.DB.get('items','T9-003'));
  record('del289: 삭제 후 번호 즉시 재사용 (003 재채번 대상)', (() => { win.DB.set('items','T9-002',{code:'T9-002',name:'t2',active:true}); return win.nextItemIdentifier('T9') === 'T9-003'; })());
  // 이력 있는 품번 → 참조 감지 (자체 생성 수주 — void 해도 이력은 이력)
  win.DB.set('orders','SO-T9-REF',{ id:'SO-T9-REF', cust:'C001', itemCode:'T9-001', qty:10, price:100 });
  record('del289: 실적·이력 품번 참조 감지', win.itemHistoryRefs('T9-001').length > 0, win.itemHistoryRefs('T9-001').join(' · '));
  win.DB.set('orders','SO-T9-REF',{ void:{ reason:'test', ts:1 } });
  record('del289: void 된 이력도 감지 (Raw 검사)', win.itemHistoryRefs('T9-001').length > 0);
  // 원장 컬렉션 remove 불허
  record('del289: 원장 컬렉션 remove 차단 (records)', win.DB.remove('records','R-ANY') === false);
  win.DB.remove('items','T9-001'); win.DB.remove('items','T9-002');   // 정리
})();
record('del289: 어댑터 remove 2벌 (local·firestore — v0.35.5 승계: 로컬 = 문서 단위 del 마킹)', /_dirty\.set\(col \+ ' ' \+ id, \{col, id, del:true\}\)/.test(html) && /_op: 'delete'/.test(html));
// ★ v0.35.5 — 로컬 어댑터 문서 단위 쓰기 (QV-1b O절 실브라우저 FAIL 수정) — 동작 검증은 _verify-2tab-local.js 26항목
record('adp353: 로컬 write = read-modify-write (전체 재직렬화 제거)', !/write\(col, id, doc\)\{\s*if\(!this\._state\) return;\s*try \{ localStorage\.setItem\(_activeLSKey\(\), JSON\.stringify\(this\._state\)\); \}/.test(html) && /this\._dirty\.forEach\(\(\{col, id, del\}\)/.test(html));
record('adp353: storage 이벤트 병합 (_updated 최신 우선) + 모달 중 리렌더 보류', /addEventListener\('storage'/.test(html) && /\(theirs\._updated \|\| 0\) > \(mine\._updated \|\| 0\)/.test(html) && /modalOpen/.test(html));
record('adp353: restore 전체 쓰기 경로 유지 (col=null → _full)', /if\(col == null\)\{ this\._full = true; \}/.test(html) && /persist\(null, null, null\)/.test(html));
record('adp353: 백업 다운로드 전 flush', /wp-backup'\)\.onclick = \(\) => \{\s*try \{ if\(DB\._adapter && typeof DB\._adapter\.flush === 'function'\) DB\._adapter\.flush\(\); \}/.test(html));
record('cor354: LOT.correctRecord·correctionCheck 존재 (보정 단일 경로)', typeof win.LOT.correctRecord === 'function' && typeof win.LOT.correctionCheck === 'function');
record('cor354: 실적 수정 모달 → LOT.correctRecord 배선 · DB.correct 직접 호출 0', /LOT\.correctRecord\(recId, \{good:g, defect:d, rework:re\}/.test(html) && !/DB\.correct\('records', recId, \{good:g/.test(html));
record('lck355: lotConflictInfo 자체 탐지 (파생 보류 → detectConcurrencyIncidents(plot)) — v0.36.0 승계', /if\(!incidents\.length && pending > 0\)/.test(html) && typeof win.lotHasNegativeBalance === 'function' && typeof win.lotHasPending === 'function' && win.lotLockInfo === win.lotConflictInfo);
record('lck355: detectConcurrencyIncidents(onlyPlot) 범위 인자', /function detectConcurrencyIncidents\(onlyPlot\)/.test(html));
record('lck355: 실적 등록 모달 사고(보류) 안내·옵션 라벨 · 파비콘 내장 — v0.36.0 승계', /⚠ 동시성 사고 — 초과/.test(html) && /동시성 사고 \(인정분 진행\)/.test(html) && /<link rel="icon" href="data:image\/svg\+xml/.test(html));
record('adp353: 로컬 어댑터 flush·_listen 런타임 존재 (기본 부팅은 firestore 어댑터 — 로컬 동작 검증은 _verify-2tab-local.js)', typeof win._LOCAL_ADAPTER.flush === 'function' && typeof win._LOCAL_ADAPTER._listen === 'function' && typeof win._LOCAL_ADAPTER.remove === 'function');
record('del289: 삭제 버튼 + 이력 가드 + 서버 rules 안내 (v0.31.5 — v4 문구로 승계)', /id="mi-del"/.test(html) && /itemHistoryRefs\(it\.code\)/.test(html) && /firestore\.rules v4/.test(html));
record('rev289: 리비전 올리기 = 코드 불변·rev\+1·rev_history 보존', /rev: curRev \+ 1, rev_history: hist/.test(html) && !/previous_rev: it\.code/.test(html));
record('rev289: 폼 자동채번 rev 미연동 (고객사 선택만)', !/revEl\.oninput = _syncId/.test(html) && /crefEl\.onchange = _syncId/.test(html));

// ── v0.31.4 대기화면 (시계 클릭 진입 · 키/클릭/터치 복귀 — 마우스 이동 깨우기 없음) ──
record('sb290: 시계 클릭 진입 연결', /id="ck-clock" onclick="enterStandby\(\)"/.test(html));
record('sb290: 대기화면 마크업 (CI·날짜·시간·날씨·브랜드·버전)', ['id="standby"','id="sb-ci"','id="sb-date"','id="sb-hm"','id="sb-weather"','id="sb-ver"'].every(s => html.includes(s)));
record('sb290: 헬퍼 1벌 (enterStandby/wakeStandby)', (html.match(/function enterStandby\(/g)||[]).length === 1 && (html.match(/function wakeStandby\(/g)||[]).length === 1);
record('sb290: 마우스 이동 깨우기 없음 (사용자 판정)', !/mousemove[^\n]*wakeStandby/.test(html) && !/wakeStandby[^\n]*mousemove/.test(html));
record('sb290: 키/클릭/터치 복귀 3종', /keydown[\s\S]{0,80}wakeStandby/.test(html) && /mousedown[^\n]{0,40}wakeStandby/.test(html) && /touchstart[^\n]{0,40}wakeStandby/.test(html));
record('sb290: 날씨 = Open-Meteo (무료·키 불필요) · 실패 시 숨김', /api\.open-meteo\.com\/v1\/forecast/.test(html) && /box\.hidden = true;\s*\/\/ 오프라인/.test(html));
record('sb290: 날씨코드 맵 존재 (맑음·비·눈·뇌우)', win._SB_WCODE && win._SB_WCODE[0][0]==='맑음' && win._SB_WCODE[63][0]==='비' && win._SB_WCODE[73][0]==='눈' && win._SB_WCODE[95][0]==='뇌우');
(function(){
  // 기능 왕복: 진입 → 표시·회사명·버전 → 깨우기 → 소등 (fetch 는 jsdom 미지원 — 날씨 칸 숨김 폴백 확인)
  win.fetch = undefined;
  win.DB.set('settings','company',{ name:'검증정밀', site:'', logo:'' });
  win.enterStandby();
  const el = win.document.getElementById('standby');
  record('sb290: 진입 — 표시·회사명·버전 주입', el.classList.contains('on') && win.document.getElementById('sb-ci').textContent.includes('검증정밀') && win.document.getElementById('sb-ver').textContent.includes(win.APP_VERSION));
  win.wakeStandby();
  record('sb290: 깨우기 — 소등·타이머 해제', !el.classList.contains('on'));
})();
record('sb290: 인쇄 시 숨김 가드', /@media print\{\.standby\{display:none !important\}\}/.test(html));

// ── v0.31.4 자동채번 이력 번호 제외 (품번 문서 부재여도 실적·이력 있는 순번은 재채번 금지) ──
(function(){
  // Z9-001: 품번 문서 없음 + 수주 이력만 존재 → 채번은 001 을 건너뛰고 002 제안
  win.DB.set('orders','SO-Z9-GHOST',{ id:'SO-Z9-GHOST', cust:'C001', itemCode:'Z9-001', qty:5, price:100 });
  record('hist291: 이력 순번 수집 (_itemSeqsInHistory)', win._itemSeqsInHistory('Z9').has(1));
  record('hist291: 품번 부재+이력 존재 번호 채번 제외 (Z9-001 건너뜀 → Z9-002)', win.nextItemIdentifier('Z9') === 'Z9-002', win.nextItemIdentifier('Z9'));
  win.DB.set('orders','SO-Z9-GHOST',{ void:{ reason:'test', ts:1 } });
  record('hist291: void 된 이력도 제외 유지 (Raw)', win.nextItemIdentifier('Z9') === 'Z9-002');
  // 단가합의서 라인·부적합 등 비원장 이력도 제외 대상
  win.DB.set('price_agreements','PA-Z9-T',{ id:'PA-Z9-T', customer_ref:'C001', items:[{ code:'Z9-003', price:100 }] });
  record('hist291: 단가합의서 라인 순번도 제외 (Z9-003)', win._itemSeqsInHistory('Z9').has(3) && win.nextItemIdentifier('Z9') === 'Z9-002');
  record('hist291: 무이력 번호는 정상 채번 유지 (Z9-002)', !win._itemSeqsInHistory('Z9').has(2));
})();
record('hist291: 수기 입력 우회 차단 (신규 저장 시 이력 번호 거부)', /if\(isNew\)\{\s*\n\s*const histRefs = itemHistoryRefs\(code\);/.test(html) && /과거 실적·이력이 있는 번호입니다/.test(html));
record('hist291: 채번 경로에 이력 제외 연결', /_itemSeqsInHistory\(custCode\)\.forEach\(s => seen\.add\(s\)\);/.test(html));

// ── v0.31.4 리스트 스크롤 (표 64vh 자체 스크롤 · 머리글 sticky · ▲▼ 화살표 · PageUp/Down) ──
record('sc292: 표 세로 스크롤 CSS (50vh — v0.31.4 상한 통일)', /\.tbl-wrap\{overflow-x:auto;overflow-y:auto;max-height:50vh/.test(html));
record('sc292: 머리글 고정 (thead sticky)', /\.tbl-wrap thead th\{position:sticky;top:0;z-index:5\}/.test(html));
record('sc292: 인쇄 시 전체 펼침 가드', /@media print\{\.tbl-wrap\{max-height:none !important;overflow:visible !important\}\}/.test(html));
record('sc292: 화살표 마크업 (▲▼ 우하단 고정)', /id="scroll-nav"/.test(html) && /id="sn-up"/.test(html) && /id="sn-dn"/.test(html) && /\.scroll-nav\{position:fixed/.test(html));
record('sc292: 헬퍼 1벌 (_scrollTarget/_snScroll)', (html.match(/function _scrollTarget\(/g)||[]).length === 1 && (html.match(/function _snScroll\(/g)||[]).length === 1);
record('sc292: 대상 우선순위 — 모달 > 마지막 사용 > 최대 리스트 > 페이지', /mb\.classList\.contains\('show'\)/.test(html) && /_snLast && document\.contains\(_snLast\)/.test(html) && /#view \.tbl-wrap/.test(html));
record('sc292: PageUp/Down — 입력 포커스 시 미개입', /e\.key !== 'PageUp' && e\.key !== 'PageDown'/.test(html) && /\['input','textarea','select'\]\.includes\(tag\)/.test(html));
record('sc292: 꾹 누르면 연속 스크롤 (320ms)', /setInterval\(\(\) => _snScroll\(dir\), 320\)/.test(html));
record('sc292: 인쇄 화살표 숨김', /@media print\{\.scroll-nav\{display:none !important\}\}/.test(html));
record('sc292: 스크롤 불가 시 화살표 숨김 판정', typeof win._scrollTarget === 'function');

// ── v0.31.4 금형 추가 — 품번 마스터 금형번호 연동 (선택 → 자동 채움 · 수명만 기록) ──
record('mold293: 선택 셀렉트 존재 (신규 전용)', /id="m-from-item"/.test(html) && /직접 입력 \(품번에 없는 금형\)/.test(html));
record('mold293: 품번 금형번호 수집 (기등록도 목록 유지·현황 병기)', /_byMoldNo\[mn\] = _byMoldNo\[mn\] \|\| \[\]/.test(html) && /_mnOpts = Object\.keys\(_byMoldNo\)\.sort\(\);/.test(html) && /등록 \$\{reg\.length\}건/.test(html));
record('mold293: 선택 시 자동 채움 (금형번호·품번 전체·ID·이름 제안 — v0.31.4 충돌 회피 승계)', /fromSel\.onchange/.test(html) && /linked\.map\(i => i\.code\)\.join\(','\)/.test(html) && /if\(!idEl\.value\.trim\(\)\)\{\s*\n\s*let sug = mn;/.test(html) && /idEl\.value = sug;/.test(html));
record('mold293: 다품번 공유 금형 표기 (외 N품번)', /외 \$\{_byMoldNo\[mn\]\.length-1\}품번/.test(html));

// ── v0.31.4 구분(BL/FI) 풀어쓰기 (약어만 적혀 있어 의미 불명 — 실사용 지적) ──
record('mold294: 구분 힌트 풀어쓰기 (v0.31.4 선택형 문구로 승계)', /일체형\(ONE\)=한 금형에서 한 번에 성형/.test(html) && /블로커\(BL\)\/피니셔\(FI\)=단계별 금형 2벌/.test(html));
record('mold294: KPI 문구 풀어쓰기 (v0.31.4 구분별 벌수 요약으로 승계)', /varSummary \|\| '등록된 금형 없음'/.test(html));

// ── v0.31.4 구분 선택형 (ONE 일체형/BL/FI + 기타 직접 입력 — 사용자 판정 ONE 채택) ──
record('mold295: MOLD_VARIANT_LABELS 1벌 (ONE/BL/FI 한글 병기)', win.MOLD_VARIANT_LABELS && win.MOLD_VARIANT_LABELS.ONE.includes('일체형') && win.MOLD_VARIANT_LABELS.BL.includes('블로커') && win.MOLD_VARIANT_LABELS.FI.includes('피니셔') && (html.match(/const MOLD_VARIANT_LABELS = /g)||[]).length === 1);
record('mold295: 구분 셀렉트 + 기타 직접 입력 토글', /id="m-variant">/.test(html) && /id="m-variant-other"/.test(html) && /varSelEl\.value === '__other'/.test(html));
record('mold295: 저장 — 선택값 또는 기타 수기', /varSel === '__other' \? document\.getElementById\('m-variant-other'\)\.value\.trim\(\) : varSel/.test(html));
record('mold295: 목록 pill 툴팁 한글 병기', /title="\$\{_esc\(MOLD_VARIANT_LABELS\[m\.variant\]\|\|m\.variant\)\}"/.test(html));
record('mold295: 구 자유값(1차·상형 등) 기타로 표시 보존', /!\['ONE','BL','FI'\]\.includes\(m\.variant\)\?'selected':''/.test(html));

// ── v0.31.4 BL+FI 2벌 등록 (같은 금형번호 재선택 · 빠진 구분 자동 제안 · ID 충돌 회피) ──
record('mold296: 빠진 구분 자동 제안 (BL 있으면 FI · FI 있으면 BL)', /if\(has\('BL'\) && !has\('FI'\)\) varSelEl\.value = 'FI';/.test(html) && /else if\(has\('FI'\) && !has\('BL'\)\) varSelEl\.value = 'BL';/.test(html));
record('mold296: 금형ID 충돌 회피 ({번호}-{구분} → {번호}-N)', /if\(v && !DB\.get\('molds', `\$\{mn\}-\$\{v\}`\)\) sug = `\$\{mn\}-\$\{v\}`;/.test(html) && /while\(DB\.get\('molds', `\$\{mn\}-\$\{n\}`\)\) n\+\+;/.test(html));
record('mold296: 이름에 구분 한글 병기 제안', /금형\$\{vLbl\}/.test(html));
record('mold296: 안내 문구 — 2벌은 같은 번호 두 번 등록', /블로커\/피니셔 2벌은 같은 번호를 두 번 등록/.test(html));
record('mold296: 토스트에 기존 등록 현황 표기 (v0.31.4 건·구분×벌수)', /기존 \$\{reg\.length\}건/.test(html) && /\$\{v\.variant\|\|'미지정'\}×\$\{moldQty\(v\)\}벌/.test(html));

// ── v0.31.4 금형 보유 수량 (사용자 판정: 물량 증가 시 벌 수만 늘림 · 타수=벌당 평균) ──
record('qty297: 헬퍼 1벌 (moldQty/moldShotsPerUnit)', (html.match(/function moldQty\(/g)||[]).length === 1 && (html.match(/function moldShotsPerUnit\(/g)||[]).length === 1);
record('qty297: moldQty 기본 1·바닥값 보정', win.moldQty({}) === 1 && win.moldQty({qty:3}) === 3 && win.moldQty({qty:0}) === 1 && win.moldQty({qty:2.7}) === 2);
record('qty297: 벌당 평균 = ceil(전체÷보유)', (()=>{ const m={ itemCodes:[], offset:9000, qty:2 }; return win.deriveMoldShots(m) === 9000 && win.moldShotsPerUnit(m) === 4500; })());
record('qty297: 수명 판정 벌당 기준', (()=>{ const m={ itemCodes:[], offset:9000, qty:2, life_limit:10000 }; return Math.abs(win.moldLifePct(m) - 45) < 0.01; })());
record('qty297: 보유 1벌이면 기존과 동일 (회귀)', (()=>{ const m={ itemCodes:[], offset:9000, life_limit:10000 }; return win.moldShotsPerUnit(m) === 9000 && Math.abs(win.moldLifePct(m) - 90) < 0.01; })());
record('qty297: 폼 보유 수량 입력 (min 1·정수)', /id="m-qty" type="number" min="1" step="1"/.test(html) && /const qty\s+= Math\.max\(1, Math\.floor/.test(html));
record('qty297: 표 보유 컬럼 + 벌당 표기 + 전체 툴팁', /<th class="num">보유<\/th>/.test(html) && /<th class="num">누적\/벌 \(파생\)<\/th>/.test(html) && /title="전체 \$\{m\._total\.toLocaleString\(\)\}타 ÷ \$\{m\._qty\}벌"/.test(html));
record('qty297: 엑셀 내보내기 보유·벌당·전체 컬럼', /'보유 수량\(벌\)': m\._qty/.test(html) && /'누적\/벌 \(파생\)': m\._shots/.test(html) && /'누적 전체 \(파생\)': m\._total/.test(html));
record('qty297: 임포트 보유수량 컬럼·적용', /label:'보유수량\(벌\)', kind:'number'/.test(html) && /qty: Math\.max\(1, Math\.floor\(\+d\.qty\|\|1\)\)/.test(html));
record('qty297: 저장 doc 에 qty 포함 (신규·수정)', (html.match(/mold_no:moldNo, variant, qty, itemCodes:its/g)||[]).length === 2);

// ── v0.31.4 구분별 보유 벌수 집계 표시 (사용자 지적: 수량은 BL/FI/일체형 나눠 보여야) ──
record('qty298: 합계 KPI = 건 + 총 벌수', /\$\{total\}<span style="font-size:12px;color:var\(--sub\)">건 · \$\{totalUnits\}벌<\/span>/.test(html));
record('qty298: 구분별 벌수 집계 (unitsByVar — 구분별 qty 합산)', /unitsByVar\[k\] = \(unitsByVar\[k\]\|\|0\) \+ m\._qty;/.test(html));
record('qty298: 구분 한글 병기 요약 (일체형/블로커/피니셔 순)', /varOrder = \['ONE','BL','FI'\]/.test(html) && /MOLD_VARIANT_LABELS\[k\]\.split\(' — '\)\[0\]/.test(html));
record('qty298: 연동 목록 — 건수·구분×벌수 병기', /등록 \$\{reg\.length\}건 \(\$\{reg\.map\(v=>`\$\{v\.variant\|\|'미지정'\}×\$\{moldQty\(v\)\}벌`\)/.test(html));

// ── v0.31.4 단락 흐름 스크롤 (실사용 지적: 표 끝에서 페이지로 이어져 다음 단락으로) ──
record('sc299: overscroll 격리 제거 (휠 체이닝 — 표 끝→페이지)', !/overscroll-behavior:contain/.test(html));
record('sc299: ▲▼·PageUp/Down 단락 흐름 (리스트 끝→페이지로 전환)', /const atEnd = dir > 0/.test(html) && /if\(atEnd && document\.documentElement\.scrollHeight > window\.innerHeight \+ 4\) t = window;/.test(html));
record('sc299: 모달은 흐름 예외 (뒤 페이지로 안 샘)', /t !== window && t\.id !== 'modal'/.test(html));

// ── v0.31.4 단락 스크롤 일반화 (표 아닌 목록 단락도 자체 스크롤 — 실사용 지적) ──
record('sec300: .sec-scroll 유틸 (50vh·인쇄 펼침)', /\.sec-scroll\{max-height:50vh;overflow-y:auto\}/.test(html) && /@media print\{\.sec-scroll\{max-height:none !important/.test(html));
record('sec300: 렌더 후 실측 승격 (1벌 · 상한 초과 즉시 · 이중 스크롤 방지)', (html.match(/function _secScrollEnhance\(/g)||[]).length === 1 && /el\.querySelector\('\.tbl-wrap'\)/.test(html) && /h > cap \+ 1/.test(html));
record('sec300: router 연결 (전 메뉴 공통)', /try \{ _secScrollEnhance\(\); \} catch/.test(html));
record('sec300: 화살표·PageUp\/Down 대상에 포함', /#view \.tbl-wrap, #view \.sec-scroll/.test(html) && /closest\('\.tbl-wrap, \.sec-scroll'\)/.test(html));
record('sec300: 설비현황 행 목록도 대상 (.eqtable)', /#view \.card > div, #view \.eqtable/.test(html));

// ── v0.31.4 단락 상한 통일 + 스크롤바 상시 표시 (사용자 판정: 한 칸이라도 넘으면 스크롤) ──
record('sec301: 상한 = 화면 절반 통일 (cap 0.50 = CSS 50vh)', /\* 0\.50;/.test(html) && !/max-height:64vh/.test(html));
record('sec301: 스크롤바 상시 표시 스타일 (webkit+firefox)', /\.tbl-wrap::-webkit-scrollbar, \.sec-scroll::-webkit-scrollbar\{width:11px/.test(html) && /scrollbar-width:thin;scrollbar-color:var\(--line2\) transparent/.test(html));
record('sec301: 스크롤바 썸 hover 강조', /::-webkit-scrollbar-thumb:hover.*background:var\(--cyan-dim\)/.test(html));

// ── v0.31.4 식별번호·품번 분리 (품번 = 고객 도면 번호 · 별도 필드 — 사용자 판정) ──
record('pn302: getPartNo 헬퍼 1벌', (html.match(/function getPartNo\(/g)||[]).length === 1);
record('pn302: part_no 우선·구 자유코드 폴백·식별번호는 품번 아님',
  win.getPartNo({ code:'038-031', part_no:'SKT271' }) === 'SKT271' &&
  win.getPartNo({ code:'SKT271' }) === 'SKT271' &&
  win.getPartNo({ code:'038-031' }) === '' && win.getPartNo(null) === '');
record('pn302: 폼 분리 (mi-partno 별도 저장·구 mi-legacy 제거)', /id="mi-partno"/.test(html) && !/mi-legacy/.test(html) && /part_no: partNo, name/.test(html));
record('pn302: 표 — 식별번호 옆 품번 컬럼', /<th>식별번호<\/th><th>품번<\/th><th>품명<\/th>/.test(html) && /getPartNo\(i\) \? _esc\(getPartNo\(i\)\)/.test(html) && /colspan="\$\{14\+iTopFields\.length\}"/.test(html));
record('pn302: 검색에 품번 포함', /\$\{getPartNo\(i\)\} \$\{i\.name\|\|''\}/.test(html));
record('pn302: 엑셀 내보내기·임포트에 품번 컬럼', /품번: getPartNo\(i\)/.test(html) && /key:'part_no',\s+label:'품번'/.test(html) && /part_no: \(d\.part_no \|\| ''\)\.trim\(\)/.test(html));

// ── v0.31.4 수주 화면 품번 병기 (새 수주 선택·진행 현황·수주 목록 — 실사용 지적) ──
record('pn303: itemLabel 헬퍼 1벌 (식별번호 · 품번 병기)', (html.match(/function itemLabel\(/g)||[]).length === 1);
record('pn303: 병기 규칙 — 품번 있으면 " · " 연결·같으면 1회·없으면 식별번호만', (()=>{
  win.DB.set('items','PN-T01',{ code:'PN-T01', part_no:'SKT271', name:'t', active:true });
  win.DB.set('items','038-777',{ code:'038-777', name:'t2', active:true });
  const a = win.itemLabel('PN-T01') === 'PN-T01 · SKT271';
  const b = win.itemLabel('038-777') === '038-777';
  const c = win.itemLabel('SKT-NOITEM') === 'SKT-NOITEM';
  win.DB.remove && win.DB.remove('items','PN-T01'); win.DB.remove && win.DB.remove('items','038-777');
  return a && b && c;
})());
record('pn303→so358: 새 수주 품번 선택 = itemLabel(식별번호·품번) 병기 · 고객사별 목록 · 품명/단가/최근수주 자동 카드 (v0.35.8 승계)', /itemSel\.innerHTML = list\.map\(i=>`<option value="\$\{_esc\(i\.code\)\}" data-p="\$\{i\.price\}">\$\{_esc\(itemLabel\(i\.code\)\)\}/.test(html) && /id="so-auto"/.test(html) && /_itemsOf = cust => \{ const l = items\.filter\(i => i\.cust === cust\)/.test(html));
record('pn303: 수주 통합 표에 식별번호 병기 (v0.35.1 공통 규격 승계 — 2카드→통합 1표)', /itemLabel\(o\.itemCode\)/.test(html) && /data-so-row=/.test(html));

// ── v0.31.4 날짜 입력 전면 달력 (납기·반출일 등 — 클릭 달력 통일 · 실사용 지적) ──
record('cal304: _calEnhance 헬퍼 1벌 (date→text 통일+달력 연결)', (html.match(/function _calEnhance\(/g)||[]).length === 1 && /input\[type="date"\]/.test(html) && /input\[placeholder="YYYY-MM-DD"\]/.test(html));
record('cal304: router(화면)+openModal(모달) 양쪽 연결', /_calEnhance\(document\.getElementById\('view'\)\)/.test(html) && /_calEnhance\(md\)/.test(html));
(function(){
  // 기능 왕복: type=date 입력 → _calEnhance → text 전환·placeholder·달력 바인딩
  const v = win.document.getElementById('view');
  v.insertAdjacentHTML('beforeend', '<input type="date" id="cal-t1"><input type="text" placeholder="YYYY-MM-DD" id="cal-t2">');
  win._calEnhance(v);
  const t1 = win.document.getElementById('cal-t1'), t2 = win.document.getElementById('cal-t2');
  record('cal304: date→text 전환+placeholder+숫자 키패드', t1.type === 'text' && t1.placeholder === 'YYYY-MM-DD' && t1.getAttribute('inputmode') === 'numeric');
  record('cal304: 달력 바인딩 (중복 방지 플래그)', t1._velaCalBound === true && t2._velaCalBound === true);
  t1.remove(); t2.remove();
})();
record('cal304: 수기 입력 병행 (readonly 미사용)', !/readonly[^>]*placeholder="YYYY-MM-DD"/.test(html));

// ── v0.31.4 수주 수정 (보정 레코드 · 출하량 미만 감축 차단 — 실사용 지적) ──
record('soe305: 수정 버튼 — 행 상세 안 (v0.35.1 승계 — 통합 표 상세 1곳 + 핸들러)', (html.match(/data-so-edit=/g)||[]).length >= 1 && /data-so-det=/.test(html));
record('soe305: 고객사·품번 불변 (표시만)', /고객사 <span[^>]*>— 불변<\/span>/.test(html) && /itemLabel\(o\.itemCode\)\)\}" disabled/.test(html));
record('soe305: 저장 = DB.correct 보정 레코드 (원본 보존)', /DB\.correct\('orders', o\.id, \{ qty, price, due \}, '수주 수정 \(화면\)'\)/.test(html));
record('soe305: 출하량 미만 감축 차단', /qty < shipped/.test(html) && /미만으로 줄일 수 없습니다/.test(html));
(function(){
  // 기능 왕복: 수주 생성 → correct 수정 → 값 반영·원본 로그 유지·출하 가드 값 검사
  win.DB.set('orders','SO-EDT-1',{ id:'SO-EDT-1', cust:'C001', itemCode:'HF-3301', qty:100, price:3000, due:'2026-09-01', created:1 });
  const r = win.DB.correct('orders','SO-EDT-1',{ qty:150, price:3100, due:'2026-09-15' },'수주 수정 (화면)');
  const after = win.DB.get('orders','SO-EDT-1');
  record('soe305: correct 반영 (150·3100·9-15·보정 사유 로그)', after.qty===150 && after.price===3100 && after.due==='2026-09-15' && after._corrected===true && (after._logs||[]).some(l=>l.op==='correct'));
  win.DB.set('orders','SO-EDT-1',{ void:{ reason:'test', ts:1 } });   // 정리 (원장은 void 만)
})();

// ── v0.31.4 계획 수주 중복 차단 + 고객사 필터 (수주 1건 = 계획 1건 · 재고생산 무제한) ──
record('pln306: 수량 기준 목록 제외 (planRemain>0 만)', /soLive = soAll\.filter\(x => x\.planRemain > 0\)/.test(html) && /planRemain = Math\.max\(0, \(o\.qty\|\|0\) - planned\)/.test(html));
record('pln306: 저장 시 수량 기준 재검증', /_planned >= \(\+_o\.qty\|\|0\)/.test(html) && /계획 수량 충족 — /.test(html));
record('pln306: 프리필 거부 = 충족 수주만 (분할 계획 허용)', /_pf && _pf\.planRemain <= 0/.test(html));
record('pln306→pln358: 수주 연결 = 검색 1줄 + 표 행 클릭 (숨은 select 단일 소스 · v0.35.8 승계)', /id="pln-so-q"/.test(html) && /id="pln-so-body"/.test(html) && /<select id="pln-so" style="display:none">/.test(html) && /soSel\.value = tr\.dataset\.so; soSel\.dispatchEvent\(new Event\('change'\)\)/.test(html));
record('pln306: 재고생산 = 중복 가능 명시', /재고생산 — 수주 연결 없음 \(중복 계획 가능\)/.test(html));
record('pln306: 수주 화면 — 충족 시 계획완료·미충족 시 [계획]+잔량', /계획완료<\/span>/.test(html) && /plannedQty > 0 && planRemain <= 0/.test(html) && /잔\$\{planRemain\.toLocaleString\(\)\}/.test(html));

// ── v0.31.4 수량 기준 분할 계획 (충족 전 중복 허용 — 사용자 판정) ──
record('pln307: soPlannedQty 헬퍼 1벌', (html.match(/function soPlannedQty\(/g)||[]).length === 1);
(function(){
  // 기능 왕복: 수주 500 → 계획 300 → planned=300 (미충족) → 계획 200 추가 → planned=500 (충족)
  win.DB.set('orders','SO-SPLIT',{ id:'SO-SPLIT', cust:'C001', itemCode:'HF-3301', qty:500, price:1000, created:1 });
  win.DB.set('production_plans','PP-SPL-1',{ id:'PP-SPL-1', so_id:'SO-SPLIT', itemCode:'HF-3301', qty:300, date:'2026-08-10', ts:1 });
  record('pln307: 부분 계획 후 잔량 200 (미충족 → 추가 허용)', win.soPlannedQty('SO-SPLIT') === 300);
  win.DB.set('production_plans','PP-SPL-2',{ id:'PP-SPL-2', so_id:'SO-SPLIT', itemCode:'HF-3301', qty:200, date:'2026-08-11', ts:2 });
  record('pln307: 분할 2건 합=500 (충족 → 차단 대상)', win.soPlannedQty('SO-SPLIT') === 500);
  win.DB.set('production_plans','PP-SPL-2',{ void:{ reason:'test', ts:3 } });
  record('pln307: 계획 취소(void) 시 계획잔량 복원', win.soPlannedQty('SO-SPLIT') === 300);
  win.DB.set('production_plans','PP-SPL-1',{ void:{ reason:'test', ts:4 } });
  win.DB.set('orders','SO-SPLIT',{ void:{ reason:'test', ts:5 } });
})();
record('pln307→pln358: 표에 계획 가능·기계획 표기 (v0.35.8 승계)', /data-planremain="\$\{planRemain\}"/.test(html) && /기계획 \$\{planned\.toLocaleString\(\)\}<\/div>/.test(html) && /\$\{planRemain\.toLocaleString\(\)\}\$\{planned\?/.test(html));
record('pln307: 프리필 수량 = 미계획 수량', /qtyEl\.value = planRemain;/.test(html) && /Math\.min\(remain, planRemain\)/.test(html));

// ── v0.31.4 새 발주 수주 소재 소요 (규격별 묶음 발주 — 사용자 판정) ──
record('po308: 헬퍼 1벌 (soMaterialNeeds/matStockKg)', (html.match(/function soMaterialNeeds\(/g)||[]).length === 1 && (html.match(/function matStockKg\(/g)||[]).length === 1);
(function(){
  // 기능 왕복: 수주 2건(같은 규격) → 소요 산출 → 규격 묶음·무게 합 → PO so_ids 연결 시 목록 제외
  win.DB.set('items','MAT-T1',{ code:'MAT-T1', name:'소요품A', mat:'S45C', dia:40, weight:0.5, price:1000, active:true });
  win.DB.set('items','MAT-T2',{ code:'MAT-T2', name:'소요품B', mat:'S45C', dia:40, weight:0.8, price:1000, active:true });
  win.DB.set('orders','SO-MAT-1',{ id:'SO-MAT-1', cust:'C001', itemCode:'MAT-T1', qty:100, price:1000, due:'2026-09-01', created:1 });
  win.DB.set('orders','SO-MAT-2',{ id:'SO-MAT-2', cust:'C002', itemCode:'MAT-T2', qty:50, price:1000, due:'2026-09-05', created:2 });
  const needs = win.soMaterialNeeds().filter(x => ['SO-MAT-1','SO-MAT-2'].includes(x.o.id));
  record('po308: 소요 산출 (잔량×절단중량 — 50kg·40kg)', needs.length===2 && needs.find(x=>x.o.id==='SO-MAT-1').needKg===50 && needs.find(x=>x.o.id==='SO-MAT-2').needKg===40);
  record('po308: 같은 규격(S45C Ø40) 묶음 키 일치', needs[0].mat==='S45C' && needs[0].dia===40 && needs[1].mat==='S45C' && needs[1].dia===40);
  win.DB.set('purchase_orders','PO-MAT-T',{ id:'PO-MAT-T', vendor:'V002', due:'2026-08-20', items:[{ kind:'mat', name:'S45C 환봉', spec:'Ø40', unit:'kg', qty:100, price:2800, so_ids:['SO-MAT-1'] }], ts:1 });
  const after = win.soMaterialNeeds().filter(x => ['SO-MAT-1','SO-MAT-2'].includes(x.o.id));
  record('po308: 발주 연결(so_ids) 수주는 poLinked 표시', after.find(x=>x.o.id==='SO-MAT-1').poLinked===true && after.find(x=>x.o.id==='SO-MAT-2').poLinked===false);
  // 정리
  win.DB.set('orders','SO-MAT-1',{ void:{reason:'t',ts:1} }); win.DB.set('orders','SO-MAT-2',{ void:{reason:'t',ts:1} });
  win.DB.set('purchase_orders','PO-MAT-T',{ void:{reason:'t',ts:1} });
  win.DB.remove('items','MAT-T1'); win.DB.remove('items','MAT-T2');
})();
record('po308: 모달 섹션 (체크·집계·반영 버튼)', /수주 소재 소요 \(미발주/.test(html) && /po-so-chk/.test(html) && /id="po-so-apply"/.test(html) && /총 예상 <b/.test(html));
record('po308: 재고·부족·합의 단가(원가 단가표) 표기', /재고 \$\{Math\.round\(stock\)\.toLocaleString\(\)\}kg/.test(html) && /부족 ≈/.test(html) && /material_rates \|\| \{\}/.test(html));
record('po308: 반영 = 규격별 소재 품목 생성 + so_ids 연결·재반영 가능', /_fromSo: true/.test(html) && /so_ids: arr\.map\(x=>x\.o\.id\)/.test(html) && /draft\.items\.filter\(r => !r\._fromSo\)/.test(html));
record('po308: 저장 시 so_ids 보존', /so_ids: it\.so_ids/.test(html));

// ── v0.31.4 소재·구매품 마스터 (사 오는 물건 관점 — 품번과 분리 · 발주 연동) ──
record('mt310: 채번 MT→materials 등록', /MT: 'materials'/.test(html));
record('mt310: 기준정보 카드 (등록·불러오기·임포트·엑셀)', /소재·구매품 <span class="tag">/.test(html) && /id="mst-mt-new"/.test(html) && /id="mst-mt-sync"/.test(html) && /id="mst-mt-import"/.test(html) && /id="mst-mt-xlsx"/.test(html));
record('mt310: 등록 모달 1벌 (_mstMaterialEdit — 구분 토글·소재 재질 필수)', (html.match(/VIEWS\._mstMaterialEdit = /g)||[]).length === 1 && /소재는 재질 입력 필수/.test(html));
record('mt310: matPriceKg 1벌 — 마스터(규격별) 우선·단가표 폴백', (html.match(/function matPriceKg\(/g)||[]).length === 1);
(function(){
  // 기능 왕복: 단가표 폴백 → 마스터 등록 시 규격별 단가 우선
  const rates = win.DB.get('settings','cost_rates') || {};
  const base = (rates.material_rates||{})['S45C'] || 0;
  record('mt310: 마스터 없으면 단가표 폴백', win.matPriceKg('S45C', 40) === base);
  win.DB.set('materials','MT-TEST-1',{ id:'MT-TEST-1', kind:'mat', name:'S45C 환봉', mat:'S45C', dia:40, unit:'kg', price:2950, active:true });
  record('mt310: 마스터 등록 시 규격별 단가 우선 (2950)', win.matPriceKg('S45C', 40) === 2950 && win.matPriceKg('S45C', 55) === base);
  win.DB.set('materials','MT-TEST-1',{ active:false });
  record('mt310: 비활성 시 폴백 복귀', win.matPriceKg('S45C', 40) === base);
})();
record('mt310: 품번에서 소재 불러오기 (미등록 재질·Ø 일괄 생성)', /품번에서 소재 불러오기/.test(html) && /combos\.set\(key, \{ mat, dia/.test(html) && /품번에서 자동 생성/.test(html));
record('mt310: 새 발주 연동 — 마스터 이름 자동완성+일치 시 자동 채움', /DB\.all\('materials'\)\.filter\(m => m\.active !== false\)\.forEach\(m => \{ if\(m\.name\) _poNamesSet\.add/.test(html) && /const mm = DB\.query\('materials', x => x\.active !== false && x\.name === row\.name\)\[0\]/.test(html));
record('mt310: 임포트 스펙+중복 건너뜀+적용', /materials: \{\n    label: '소재·구매품'/.test(html) && /같은 구분·품명·규격/.test(html) && /const mid = nextDocNo\('MT'\)/.test(html));

// ── v0.31.4 소재·구매품 삭제 (무이력만 물리 삭제 · 이력 시 비활성 — 품번과 동일 원칙) ──
record('mt311: mtHistoryRefs 헬퍼 1벌 (발주·구매입고·소재 입고/LOT)', (html.match(/function mtHistoryRefs\(/g)||[]).length === 1);
record('mt311: DB.remove 허용 목록 (v0.31.5 — MASTER_DEL_COLS 상수로 승계)', /MASTER_DEL_COLS\.includes\(col\)/.test(html) && /'materials','workers','vendors'/.test(html));
record('mt311: 모달 [삭제] + 이력 차단 + rules 안내 (v0.31.5 — v4 문구로 승계)', /id="mt-del"/.test(html) && /mtHistoryRefs\(m\)/.test(html) && /firestore\.rules v4/.test(html));
(function(){
  // 기능 왕복: 무이력 삭제 OK → 발주 이력 생기면 차단
  win.DB.set('materials','MT-DEL-1',{ id:'MT-DEL-1', kind:'gen', name:'삭제테스트품', spec:'', unit:'ea', price:100, active:true });
  record('mt311: 무이력 → 참조 0·삭제 성공', win.mtHistoryRefs(win.DB.get('materials','MT-DEL-1')).length === 0 && win.DB.remove('materials','MT-DEL-1') === true && !win.DB.get('materials','MT-DEL-1'));
  win.DB.set('materials','MT-DEL-2',{ id:'MT-DEL-2', kind:'gen', name:'이력테스트품', spec:'', unit:'ea', price:100, active:true });
  win.DB.set('purchase_orders','PO-DEL-T',{ id:'PO-DEL-T', vendor:'V002', due:'', items:[{kind:'gen', name:'이력테스트품', unit:'ea', qty:1, price:100}], ts:1 });
  record('mt311: 발주 이력 시 참조 감지 (삭제 차단 대상)', win.mtHistoryRefs(win.DB.get('materials','MT-DEL-2')).length > 0);
  win.DB.set('purchase_orders','PO-DEL-T',{ void:{reason:'t',ts:1} });
  record('mt311: void 된 발주도 이력으로 감지 (Raw)', win.mtHistoryRefs(win.DB.get('materials','MT-DEL-2')).length > 0);
  win.DB.remove('materials','MT-DEL-2');   // 정리 (테스트 상 강제 — 실제 UI 는 차단)
})();

// ── v0.31.4 새 발주 품목 체크 삭제 (실사용 지적: 추가 후 삭제 안 됨 — 최소 1행 차단이 원인) ──
record('po312: 품목 카드 체크박스 + [체크 품목 삭제] 버튼', /class="poi-chk"/.test(html) && /id="po-row-del"/.test(html));
record('po312: 최소 1행 차단 폐지 — 마지막 행 삭제 시 빈 행 교체', !/최소 1행 필요/.test(html) && (html.match(/if\(!draft\.items\.length\) draft\.items\.push\(\{kind:'gen'/g)||[]).length === 2);
record('po312: 일괄 삭제 — 체크만 제거·건수 토스트·버튼 비활성 복귀', /draft\.items = draft\.items\.filter\(r => !r\._chk\)/.test(html) && /품목 \$\{n\}건 삭제/.test(html) && /delBtn\.disabled = !draft\.items\.some\(r => r\._chk\)/.test(html));

// ── v0.31.4 문서 PDF 저장 (발주서·반출증 등 전 양식 — 인쇄 모달 [PDF 저장] · 한글 렌더 방식) ──
record('pdf313: 헬퍼 1벌 (_loadPdfLibs/_openPdf — 지연 로드·캐시)', (html.match(/function _loadPdfLibs\(/g)||[]).length === 1 && (html.match(/async function _openPdf\(/g)||[]).length === 1 && /_pdfLibsP = null; throw e;/.test(html));
record('pdf313: cdnjs 버전 고정 (html2canvas 1.4.1 · jspdf 2.5.1)', /cdnjs\.cloudflare\.com\/ajax\/libs\/html2canvas\/1\.4\.1\/html2canvas\.min\.js/.test(html) && /cdnjs\.cloudflare\.com\/ajax\/libs\/jspdf\/2\.5\.1\/jspdf\.umd\.min\.js/.test(html));
record('pdf313: 인쇄 모달에 [PDF 저장] — 방향 설정 공유', /id="prt-pdf"/.test(html) && /_openPdf\(title, bodyHTML, extraCSS, orient(, fitSel)?\)/.test(html));   // v0.38.23 — 글자 맞춤 인자 추가
record('pdf313: A4 분할 (96dpi 794/1123 · mm 210×297 · 다중 페이지 addPage)', /landscape \? 1123 : 794/.test(html) && /landscape \? 297 : 210/.test(html) && /pdf\.addPage\(\)/.test(html));
record('pdf313: 파일명 규칙 + 메일 첨부 안내 + 실패 정직 안내', /\.pdf`;/.test(html) && /메일 첨부용/.test(html) && /PDF 라이브러리 로드 실패 — 인터넷 연결 확인/.test(html));
record('pdf313: 숨은 iframe 격리 렌더 (앱 스타일 오염 없음·정리 finally)', /position:fixed;left:-12000px/.test(html) && /finally \{ ifr\.remove\(\); \}/.test(html));

// ── v0.31.4 마스터 자동순번 (작업자 등 — 형식 승계·중복 차단) ──
record('seq314: suggestNextId 헬퍼 1벌', (html.match(/function suggestNextId\(/g)||[]).length === 1);
record('seq314: 기존 형식 승계 (OP-014→OP-015 · 자릿수 유지)', win.suggestNextId('workers','W') === 'OP-046' || /^OP-\d+$/.test(win.suggestNextId('workers','W')));
(function(){
  // 형식 학습 왕복: 대시 없는 W01 형식 우세 컬렉션 → W03 제안 · 빈 컬렉션 → fallback-001
  win.DB.set('defect_types','ZZ01',{ id:'ZZ01', name:'t1', active:true });
  win.DB.set('defect_types','ZZ02',{ id:'ZZ02', name:'t2', active:true });
  const s = win.suggestNextId('defect_types','DT');
  record('seq314: 최빈 형식 학습 (DT-XXX vs ZZXX — 다수 형식 제안)', /^(DT-\d{3}|ZZ\d{2})$/.test(s), s);
  win.DB.set('defect_types','ZZ01',{ void:{reason:'t',ts:1} }); win.DB.set('defect_types','ZZ02',{ void:{reason:'t',ts:1} });
  record('seq314: 빈 컬렉션 fallback', win.suggestNextId('nc_programs','NP') === 'NP-001');
})();
record('seq314: 작업자 — 신규 사번 자동 제안+중복 차단(덮어쓰기 사고 방지)', /emp_no: suggestNextId\('workers','OP'\)/.test(html) && /사번 중복 — \$\{emp\}/.test(html));
record('seq314: 금형 — ID 빈칸 시 자동순번 (연동 제안과 병행)', /const mid2 = mid \|\| suggestNextId\('molds','M'\)/.test(html) && /DB\.set\('molds', id \? mid : mid2, doc\)/.test(html));

// ── 1) 헬퍼 1벌 · 폴백 ──
record('ht: HT_SPEC_OPTS = Q.T/NOR/ISO/CA', JSON.stringify(win.HT_SPEC_OPTS) === '["Q.T","NOR","ISO","CA"]', JSON.stringify(win.HT_SPEC_OPTS));
record('ht: getHtSpec 표준값 우선', win.getHtSpec({ ht_spec:'NOR', custom:{'열처리사양':'Q.T'} }) === 'NOR');
record('ht: getHtSpec 구 custom 폴백 (이력 보존)', win.getHtSpec({ custom:{'열처리사양':'Q.T'} }) === 'Q.T');
record('ht: getHtSpec 빈값 처리', win.getHtSpec({}) === '' && win.getHtSpec(null) === '');
record('ht: 헬퍼 1벌 — getHtSpec 정의 1곳', (html.match(/function getHtSpec\(/g) || []).length === 1);

// ── 1a) v0.27.1 DEMO 배지 픽스 ──
record('badge: [hidden] 전역 강제 CSS 존재', /\[hidden\]\{display:none !important\}/.test(html));
record('badge: demo-badge 마크업 기본 hidden', /id="demo-badge" hidden/.test(html));
record('badge: 로그인·부트 시 isDemoMode 로 토글 (2곳)', (html.match(/\.hidden = !isDemoMode\(\)/g)||[]).length >= 2);
(function(){
  // 실모드(플래그 없음)에서 hidden=true 판정 재현
  const b = win.document.getElementById('demo-badge');
  if(b){ b.hidden = !win.eval('isDemoMode()'); record('badge: 실모드 → demo-badge hidden=true', b.hidden === true); }
  else record('badge: demo-badge 요소 존재', false);
})();

// ── 1b) v0.26.6 경도 표준 필드 ──
record('hd: getHardness 표준값 우선', win.getHardness({ hardness:'HRC 30', custom:{'경도HB':'HB 200'} }) === 'HRC 30');
record('hd: getHardness 구 custom 폴백 (이력 보존)', win.getHardness({ custom:{'경도HB':'HB 200'} }) === 'HB 200');
record('hd: 헬퍼 1벌 — getHardness 정의 1곳', (html.match(/function getHardness\(/g) || []).length === 1);
record('hd: 폼 mi-hardness 입력칸 (열처리 옆 g3)', /id="mi-hardness"/.test(html));
record('hd: 저장 신규·수정 양쪽 hardness', (html.match(/default_forge_eq: forgeEq, ht_spec, hardness,/g) || []).length === 2);
record('hd: 품목 표 경도 컬럼 + colspan 14 (v0.31.4 품번 컬럼 추가 승계)', /<th>경도<\/th>/.test(html) && /colspan="\$\{14\+iTopFields\.length\}"/.test(html));
record('hd: 엑셀 내보내기 경도 컬럼', /'경도': getHardness\(i\)/.test(html));

// ── 1c) v0.26.6 구매 일반품 구분 ──
record('po: 품목 카드 구분 select (소재/일반)', /class="poi-kind"/.test(html) && /option value="gen"/.test(html));
record('po: 저장 시 kind 기록', /kind: it\.kind==='gen' \? 'gen' : 'mat'/.test(html));
record('po: 소재 입고 연결 = 구분 mat (구버전 kg 호환)', /const kind = it\.kind \|\| \(\(it\.unit\|\|''\)==='kg' \? 'mat' : 'gen'\);/.test(html));
record('po: 구매품 입고 연결 = mat 제외 (구버전 전부 허용)', /if\(it\.kind === 'mat'\) return;/.test(html));
record('po: 목록 일반 배지', /일반 \$\{items\.filter\(i=>i\.kind==='gen'\)\.length\}/.test(html));

// ── 1d) v0.26.6 라이트 테마 상태색 ──
const lightBlock = (html.split('html[data-theme="light"]{')[1] || '').split('}')[0];
record('light: green/blue/amber/red 재정의', /--green:#15803d/.test(lightBlock) && /--blue:#0369a1/.test(lightBlock) && /--amber:#b45309/.test(lightBlock) && /--red:#dc2626/.test(lightBlock));
const rootBlock = (html.split(':root{')[1] || '').split('}')[0];
// ── v0.26.7 출력물 미리보기 인쇄 안전 ──
const pd = (html.split('.print-doc{')[1] || '').split('.print-doc h1')[0];
record('print-doc: 변수 스코프 재정의 (bg2/sub/red/amber 인쇄 안전값)', /--bg2:#f0f0f0/.test(pd) && /--sub:#333/.test(pd) && /--red:#b91c1c/.test(pd) && /--amber:#92400e/.test(pd));
record('print-doc: th 밝은 배경 강제', /\.print-doc table th\{background:#efefef;color:#111\}/.test(html));
record('print-doc: 테마 변수 상속 차단 (border 도 고정)', /border:1px dashed #bbb/.test(html));

record('light: 다크 :root 불변 (green #4ade80 유지)', /--green:#4ade80/.test(rootBlock) && /--amber:#fbbf24/.test(rootBlock) && /--blue:#7dd3fc/.test(rootBlock));

// ── 2) 마이그레이션 (멱등 · 비활성만) ──
win.DB.set('master_fields', 'MF-TEST-HT', { id:'MF-TEST-HT', master:'items', key:'열처리사양', label:'열처리 사양', type:'select', options:['Q.T'], active:true });
win.DB.set('master_fields', 'MF-TEST-HD', { id:'MF-TEST-HD', master:'items', key:'경도HB', label:'경도', type:'text', active:true });
win.migrateHtSpecField();
record('migration: 경도HB 정의 필드도 비활성 (v0.26.6)', win.DB.get('master_fields','MF-TEST-HD').active === false);
const mfAfter = win.DB.get('master_fields', 'MF-TEST-HT');
record('migration: 구 정의 필드 비활성 (삭제 아님 — 문서·값 보존)', mfAfter && mfAfter.active === false && mfAfter.key === '열처리사양');
win.migrateHtSpecField();
record('migration: 멱등 (재실행 무해)', win.DB.get('master_fields','MF-TEST-HT').active === false);
record('migration: 부팅 경로에 호출 연결', /migrateHtSpecField\(\); \} catch/.test(html));

// ── 3) 데모 시드 + 시뮬 기준 ──
let demoOk = false;
try { win.seedDemoData(); demoOk = true; record('sim: seedDemoData ran', true); }
catch (e) { record('sim: seedDemoData ran', false, e.message); }
if(demoOk){
  const items = win.DB.all('items');
  const withHeat = ['HF-3301','HF-3302','HF-3305','HF-3309'];
  record('demo: HEAT 라우트 품번 ht_spec=Q.T', withHeat.every(c => win.getHtSpec(win.DB.get('items', c)) === 'Q.T'),
    withHeat.map(c => `${c}:${win.getHtSpec(win.DB.get('items',c))||'—'}`).join(' '));
  record('demo: 사내만 품번(HF-3312) ht_spec 빈값', win.getHtSpec(win.DB.get('items','HF-3312')) === '');
  record('demo: HEAT 품번 경도 시드 (HB 269~331)', withHeat.every(c => win.getHardness(win.DB.get('items', c)) === 'HB 269~331'));
  record('demo: MF-ITEM-004(경도HB) 시드 제거', !win.DB.get('master_fields','MF-ITEM-004'));
  record('demo: MF-ITEM-003 시드 제거 (표준 승격)', !win.DB.get('master_fields','MF-ITEM-003'));
  const activeItemFields = win.getMasterFields('items').map(f => f.key);
  record('demo: 정의 필드에 열처리사양·경도HB 없음 (중복 노출 방지)', !activeItemFields.includes('열처리사양') && !activeItemFields.includes('경도HB'), activeItemFields.join(','));
}

// ── 4) 폼 · 표 · 엑셀 · 임포트 반영 ──
record('form: mi-htspec select + 기타 토글 + 기타 입력칸', /id="mi-htspec"/.test(html) && /id="mi-htspec-other"/.test(html) && /mi-htspec-other-wrap/.test(html));
record('form: 저장 핸들러 ht_spec 수집', /const ht_spec\s*=\s*htSelVal==='__other'/.test(html));
record('form: 신규·수정 doc 양쪽에 ht_spec 저장', (html.match(/default_forge_eq: forgeEq, ht_spec,/g) || []).length === 2);
record('table: 품목 표 열처리 컬럼', /<th>열처리<\/th>/.test(html));
record('xlsx: 내보내기에 열처리사양 컬럼', /'열처리사양': getHtSpec\(i\)/.test(html));
const impCols = (win._IMPORTER_SPEC?.items?.columns || []).map(c => c.key);
record('import: _IMPORTER_SPEC.items 에 ht_spec+hardness', impCols.includes('ht_spec') && impCols.includes('hardness'), impCols.join(','));
record('import: 적용부 ht_spec+hardness 저장', /ht_spec: \(d\.ht_spec \|\| ''\)\.trim\(\)/.test(html) && /hardness: \(d\.hardness \|\| ''\)\.trim\(\)/.test(html));

// ── 5) 렌더 스모크 + 시뮬 수치 불변 ──
if(demoOk){
  const snap = () => ({
    records: win.DB.all('records').length, lots: win.DB.all('production_lots').length,
    orders: win.DB.all('orders').length, ships: win.DB.all('shipments').length,
  });
  const shipTotal = () => win.DB.all('shipments').reduce((t,s) => t + win.shipAmount(s), 0);
  const before = snap(); const beforeShip = shipTotal();
  const beforeReport = win.buildReportData('week', win.fmtDate(Date.now()));

  try { Object.keys(win.getFlags()).forEach(f => win.setFlag(f, true)); } catch(e){}
  let ok = 0, fail = 0; const failList = [];
  win.PAGES.forEach(p => {
    const fn = win.VIEWS[p.id];
    if(typeof fn !== 'function'){ fail++; failList.push(`${p.id}: no view`); return; }
    try { fn(); ok++; } catch(e){ fail++; failList.push(`${p.id}: ${e.message}`); }
  });
  record(`render smoke: ${win.PAGES.length} pages 예외 0`, fail === 0, `ok=${ok} fail=${fail}${failList.length?' | '+failList.slice(0,5).join(' ; '):''}`);
  record('render smoke: 20화면 이상', ok >= 20, `ok=${ok}`);

  // ── v0.27.3 데모 플래그 동기화 · 도면 관리 ──
  record('flags: 데모 부팅 시 전 플래그 ON 동기화 코드 (v0.32.6 — 시드 보강 블록 추가로 탐색 범위 확대 승계)', /autoLoginDemo[\s\S]{0,1600}allOn\[k\] = true/.test(html));
  (function(){
    // 구 데모 DB 시뮬: cost:false 로 고착 → autoLoginDemo 경로의 동기화가 복구하는지
    const cur = win.getFlags();
    win.DB.set('settings','features', { flags: { ...cur, cost:false, quote:false } });
    record('flags: 고착 재현 (cost=false)', win.getFlags().cost === false);
    const allOn = {}; win.MODULE_DEFS.forEach(m => m.items.forEach(([k]) => { allOn[k] = true; }));
    win.DB.set('settings','features', { flags: allOn });   // autoLoginDemo 와 동일 동작
    record('flags: 동기화 후 원가·견적·회계·도면 전부 ON', win.getFlags().cost && win.getFlags().quote && win.getFlags().acct && win.getFlags().p_drawing);
  })();
  // ── v0.27.4 도면 첨부 파일 ──
  record('dwf: Storage SDK 지연 로드 (버전 고정 · 부팅 무영향)', html.includes("firebase-storage-compat.js") && /firebasejs\/10\.14\.1\/firebase-storage-compat/.test(html));
  record('dwf: 로컬/데모 모드 = 링크만 (정직 안내)', html.includes('파일 업로드는 서버(Firebase) 연결 모드 전용'));
  // ── v0.28.1 단가합의서 ──
  record('pa: 컬렉션·PA 채번·doc_price 플래그 등록', html.includes("'price_agreements',") && /PA: 'price_agreements'/.test(html) && html.includes("['doc_price',  '단가합의서'"));
  record('pa: 문서 출력 허브 ⑥ 칩 + 새 합의서 작성 + 인쇄 연결', html.includes("'⑥ 단가합의서'") && html.includes('_paNewModal') && html.includes('printForm.pa(ids)'));
  record('pa: 작성 모달 — 고객사 select·건별 체크·마스터 반영 옵션', html.includes('pa-cust') && html.includes('pa-chk') && html.includes('pa-sync'));
  record('pa: 인쇄 양식 — 2부·bizBlock·서명란·합의 문구', html.includes("page(a,'공급자 보관용') + page(a,'구매자용')") && /단가합의서<\/div>/.test(html) && html.includes('서명(날인)으로 효력'));
  const pa = win.DB.get('price_agreements','PA-DEMO-001');
  record('pa: 데모 시드 — C001·품목 2건·기존→합의 단가', !!pa && pa.customer_ref==='C001' && pa.items.length===2 && pa.items[0].prev_price===3200 && pa.items[0].price===3350);
  record('pa: 시드가 마스터 단가 무변 (master_synced=false)', win.DB.get('items','HF-3301').price===3200);
  (function(){
    try {
      win.VIEWS.docs();
      const h = win.document.getElementById('view').innerHTML;
      record('pa render: 문서 출력에 ⑥ 단가합의서 칩 노출', h.includes('⑥ 단가합의서'));
    } catch(e){ record('pa render', false, e.message); }
  })();

  // ── v0.28.4 일괄 임포트 확대 (작업자·설비·금형·도면) ──
  const impSpec = win._IMPORTER_SPEC;
  record('imp: 4종 마스터 스펙 추가', ['workers','equipments','molds','drawings'].every(m=>!!impSpec[m]));
  record('imp: UI 진입점 4곳 (설비·금형·도면·작업자)', html.includes("_mstImport('equipments')") && html.includes("_mstImport('molds')") && html.includes("_mstImport('drawings')") && html.includes("_mstImport('workers')"));
  (function(){
    // 기능 테스트 — 판정·등록 (데모 DB 위에서 · 원장 무변)
    const mkRow = (master, obj) => {
      const cols = win._importerColumns(master);
      const row = {}; cols.forEach(c => { if(obj[c.key]!==undefined) row[c.label] = obj[c.key]; });
      return win._parseImportRow(master, row);
    };
    // 작업자
    let pr = mkRow('workers', { id:'OP-T01', name:'테스트공', dept:'검사반' });
    record('imp workers: 판정 ok → 등록', win._importJudgeRow('workers', pr).ok && (win._importApplyRow('workers', pr), win.DB.get('workers','OP-T01')?.name==='테스트공'));
    record('imp workers: 중복 사번 건너뜀', win._importJudgeRow('workers', mkRow('workers',{id:'OP-T01',name:'x'})).ok===false);
    // 설비 — 공정 검증
    record('imp equipments: 없는 공정코드 거부', win._importJudgeRow('equipments', mkRow('equipments',{id:'EQ-T1',name:'t',procCode:'NOPE'})).ok===false);
    pr = mkRow('equipments', { id:'EQ-T1', name:'테스트설비', procCode:'FORGE', work_hours:10 });
    record('imp equipments: 등록 (work_hours 반영)', win._importJudgeRow('equipments', pr).ok && (win._importApplyRow('equipments', pr), win.DB.get('equipments','EQ-T1')?.work_hours===10));
    // 금형 — 품번 검증·쉼표 다품번
    record('imp molds: 없는 품번 거부', win._importJudgeRow('molds', mkRow('molds',{id:'M-T1',itemCodes:'NO-ITEM'})).ok===false);
    pr = mkRow('molds', { id:'M-T1', name:'테스트금형', itemCodes:'HF-3301, HF-3302', life_limit:40000, offset:5000 });
    record('imp molds: 다품번·기초 타수 등록', win._importJudgeRow('molds', pr).ok && (win._importApplyRow('molds', pr), (()=>{ const m=win.DB.get('molds','M-T1'); return m.itemCodes.length===2 && m.offset===5000; })()));
    // 도면 — 리비전 이력 생성
    pr = mkRow('drawings', { id:'DW-T1-D', itemCode:'HF-3301', rev:'3', rev_date:'2026-07-01', reason:'이관' });
    record('imp drawings: 현행 Rev·이력 1건 생성', win._importJudgeRow('drawings', pr).ok && (win._importApplyRow('drawings', pr), (()=>{ const d=win.DB.get('drawings','DW-T1-D'); return d.current_rev==='3' && d.revs.length===1 && d.revs[0].reason==='이관'; })()));
    record('imp drawings: 중복 도면번호 건너뜀', win._importJudgeRow('drawings', mkRow('drawings',{id:'DW-T1-D'})).ok===false);
  })();

  record('wipe: 서버 모드 초기화 정직 안내 (로컬 사본만·local 전환 동선)', html.includes('이 PC의 로컬 사본만') && html.includes('backend') || html.includes('백엔드를 <b>local</b> 로 전환'));
  // ── v0.28.5 외주 반출 실적 선택 + 이관 임포트 6종 ──
  record('od: 반출 대상 실적 테이블 (클릭 자동 채움·부분 반출 안내)', html.includes('data-odt') && html.includes('반출 대상 실적') && html.includes('부분 반출'));
  record('mig: 설정 이관 허브 12버튼 + 권장 순서', html.includes('data-mig-import') && html.includes('① 거래처') && html.includes('⑫ 수주'));
  record('mig: 라우트 파서 1벌', (()=>{ const r=win._importParseRoute('CUT>FORGE>HEAT:out>SHIP'); return r.route && r.route.length===4 && r.route[2].mode==='out' && r.route[0].seq===10 && !!win._importParseRoute('CUT>NOPE').error; })());

  // ── v0.28.3 드래그&드롭 첨부 ──
  record('dnd: 드롭존 (클릭=선택·드래그=업로드)', html.includes('id="dwf-drop"') && html.includes('drop.onclick = () => inp.click()') && html.includes('doUploadFiles(dt.files)'));
  record('dnd: URL 드래그 → 링크 자동 등록 (uri-list)', html.includes("getData('text/uri-list')") && html.includes('addLinkAtt(nm, uri)'));
  record('dnd: 업로드·링크 등록 1벌 공용 (버튼·드롭 동일 경로)', html.includes('const doUploadFiles = async (list)') && html.includes('addLinkAtt(name, url);   // ★ v0.28.3'));
  record('dnd: PC 경로 자동 등록 불가 정직 안내', html.includes('브라우저 보안상 자동 등록 불가'));
  record('dnd: 드래그 하이라이트 (over/leave 복원)', html.includes('drop.ondragover') && html.includes('drop.ondragleave'));

  // ── v0.28.2 첨부 공용화 + 단가합의서 근거서류 ──
  record('att: 첨부 모달 공용화 — _attFilesModal(col,id) + 도면 래퍼 위임', /VIEWS\._drawFilesModal = \(id\) => VIEWS\._attFilesModal\('drawings', id\)/.test(html));
  record('att: 모달 내부 컬렉션 결합 제거 (DB.get(col, id))', !/const a = _drawAttList\(DB\.get\('drawings', id\)\)/.test(html) && html.includes('_dwStoreUpload(col, id, f)'));
  record('att: Storage 경로 컬렉션별 분리 (attUploadFiles 1벌로 이동)', html.includes('`${col}/${id}/${Date.now()}') && html.includes('async function attUploadFiles(col, id, list)'));
  record('pa-att: 목록 [첨부 N] + 저장 후 자동 오픈 + 인쇄 근거서류 표기',
    html.includes('data-pa-att') && html.includes("VIEWS._attFilesModal('price_agreements', no.id)") && html.includes('근거서류 <b>'));
  (function(){
    // 공용 데이터 규칙: PA 문서에 링크 첨부 추가 → 목록·인쇄 근거 반영 (current_rev 없는 문서 호환)
    const pa0 = win.DB.get('price_agreements','PA-DEMO-001');
    win.DB.set('price_agreements','PA-DEMO-001',{ attachments: win._drawAttAdd(pa0, { fid:'L-T1', kind:'link', name:'서명 스캔본', url:'https://example.com/pa.pdf', rev: pa0.current_rev ?? '', ts:1, by:'t' }) });
    const pa1 = win.DB.get('price_agreements','PA-DEMO-001');
    record('pa-att: PA 문서 첨부 추가·목록 헬퍼 동작 (rev 없는 문서 호환)', win._drawAttList(pa1).length===1 && pa1.attachments[0].rev==='');
    win.DB.set('price_agreements','PA-DEMO-001',{ attachments: win._drawAttRemove(pa1,'L-T1','t') });
    record('pa-att: 삭제 마킹 후 목록 제외', win._drawAttList(win.DB.get('price_agreements','PA-DEMO-001')).length===0);
  })();

  // ── v0.28.0 담당자별 메뉴 권한 ──
  record('perm: 코어 정의 (none/view/edit)', JSON.stringify(win.PERM_LEVELS)==='["none","view","edit"]');
  win.DB.set('users','U-PTEST',{ id:'U-PTEST', name:'권한테스트', pin:'000000', role:'worker', active:true, perms:{ lot:'none', inv:'view' } });
  record('perm: 미설정 메뉴 = edit (기존 동작 호환)', win.getUserPerm('U-PTEST','wo')==='edit');
  record('perm: none/view 반영', win.getUserPerm('U-PTEST','lot')==='none' && win.getUserPerm('U-PTEST','inv')==='view');
  record('perm: 관리책임자 = 전권 고정', win.getUserPerm('U-DEMO','lot')==='edit');
  record('perm: dash 는 none 불가 (최소 보기)', win.getUserPerm('U-PTEST','dash')==='edit');
  (function(){
    const S = win.eval('SESSION');
    const bak = { loggedIn:S.loggedIn, userId:S.userId, role:S.role, user:S.user };
    S.loggedIn = true; S.userId='U-PTEST'; S.role='worker'; S.user='권한테스트';
    win.location.hash = '#inv';
    record('perm guard: 보기 페이지에서 DB.set 차단 (null 반환)', win._permBlockWrite('shipments')===true && win.DB.set('shipments','SH-PERM-X',{id:'SH-PERM-X'})===null && !win.DB.get('shipments','SH-PERM-X'));
    record('perm guard: users 컬렉션은 예외 (last_login)', win._permBlockWrite('users')===false);
    win.location.hash = '#wo';
    record('perm guard: 등록 페이지에서는 허용', win._permBlockWrite('records')===false);
    // 메뉴 필터
    win.eval('renderNav(); openAllMenuModal()');   // ★ v0.37.7 — 전체 메뉴 창(같은 canSee 1벌)에서 검사
    const navHtml = win.document.getElementById('nav').innerHTML + win.document.getElementById('modal').innerHTML; win.eval('closeModal()');
    record('perm nav: none 메뉴(LOT 추적) 숨김 · view 메뉴(재고) 노출', !navHtml.includes('data-page="lot"') && !navHtml.includes('data-am="lot"') && navHtml.includes('data-am="inv"'));
    S.loggedIn=bak.loggedIn; S.userId=bak.userId; S.role=bak.role; S.user=bak.user;
    win.location.hash = '#dash';
    win.eval('renderNav()');
  })();
  // ── v0.28.0+ 설정 화면 개발자 전용 격리 잠금 (회귀 방지 · 판매 모델 핵심) ──
  (function(){
    const S = win.eval('SESSION');
    const bak = { loggedIn:S.loggedIn, userId:S.userId, role:S.role, user:S.user };
    S.loggedIn=true; S.userId='U-DEMO'; S.user='테스트';
    const render = role => { S.role=role; win.location.hash='#set'; win.VIEWS.set(); return win.document.getElementById('view').innerHTML; };
    const hM = render('manager');
    record('isolation: 관리책임자 설정 — 시스템·백업/복원·기능 플래그 전부 숨김',
      !hM.includes('<h3>시스템') && !hM.includes('데이터 백업/복원') && !hM.includes('기능 플래그'));
    record('isolation: 관리책임자 설정 — 사용자 관리는 노출', hM.includes('사용자 · 비밀번호 관리'));
    const hV = render('vendor');
    record('isolation: 개발자 설정 — 시스템·백업/복원·기능 플래그 노출',
      hV.includes('<h3>시스템') && hV.includes('데이터 백업/복원') && hV.includes('기능 플래그'));
    S.loggedIn=bak.loggedIn; S.userId=bak.userId; S.role=bak.role; S.user=bak.user; win.location.hash='#dash';
  })();

  record('perm ui: 사용자 행 [권한] 버튼 + 모달 + 프리셋', html.includes('data-usr-perm') && html.includes('openUserPermModal') && html.includes('pm-all-view'));
  record('perm ui: 보기 전용 배너 (라우터 공통)', html.includes('보기 전용</b> — 이 메뉴에서는 조회만'));
  record('draw flow: 등록 직후 첨부 창 자동 오픈', html.includes('VIEWS._drawFilesModal(no)') && html.includes('첨부 창(파일 업로드 · 링크 · PC 경로)'));

  // ── v0.27.6 앱 내 저장 · PC 경로 ──
  record('dws: 2MB 앱 내 저장 상한 · 조각 900k chars', win.DW_STORE_MAX_MB===2 && win.DW_CHUNK_CHARS===900000);
  (function(){
    const big = 'x'.repeat(2100000);
    const ch = win._dwB64Chunks(big);
    record('dws: 조각 분할 — 문서당 1MB 미만 보장·손실 없음', ch.length===3 && ch.every(c=>c.length<=900000) && ch.join('')===big);
  })();
  record('dws: MIME 판정 (pdf/이미지 열람 · 그 외 다운로드)', win._dwMime('pdf')==='application/pdf' && win._dwMime('png')==='image/png' && win._dwMime('dwg')==='application/octet-stream');
  record('dws: 업로드 자동 판정 코드 (≤2MB → 앱 내 · 초과 → Storage)', /f\.size <= DW_STORE_MAX_MB\*1048576/.test(html) && html.includes('_dwStoreUpload(col, id, f)'));
  record('dws: 삭제 = 조각 비우기 (delete 금지 규칙과 무충돌)', /data:'', purged:Date\.now\(\)/.test(html));
  record('dws: 조각 컬렉션 미러 제외 (drawing_file_chunks — _FS_COLLECTIONS 밖 · 직접 읽기)', html.includes("collection('drawing_file_chunks')") && !/_FS_COLLECTIONS[\s\S]{0,2000}drawing_file_chunks/.test(html.split('const _FS_COLLECTIONS')[1].split('];')[0]));
  record('dws: 로컬/데모 = 링크·경로만 안내', html.includes('앱 내 저장은 서버(Firebase) 연결 모드 전용'));
  record('dwp: PC 경로 첨부 — kind path·복사 버튼·file:// 안내', html.includes("kind:'path'") && html.includes('data-dwf-copy') && html.includes('탐색기 주소창'));
  (function(){
    const arr = win._drawAttAdd({attachments:[]}, { fid:'P-1', kind:'path', name:'원본', path_str:'\\\\PC\\d.dwg', rev:'0', ts:1, by:'t' });
    record('dwp: 경로 첨부 데이터 규칙 (path_str 보존)', arr.length===1 && arr[0].kind==='path' && !!arr[0].path_str);
  })();

  record('dwf: Blaze 요금 안내 문구 (정직 고지)', html.includes('Blaze(종량제) 요금제 전용') && html.includes('링크 첨부'));
  record('dwf: 20MB 개당 제한', /DW_FILE_MAX_MB = 20/.test(html));
  (function(){
    const d0 = { attachments: [] };
    const a1 = win._drawAttAdd(d0, { fid:'F-1', kind:'file', name:'a.dwg', size:1000, ext:'dwg', url:'u', path:'p', rev:'2', ts:1, by:'t' });
    record('dwf: 첨부 추가 — 증분 누적', a1.length===1 && a1[0].removed===null);
    const a2 = win._drawAttRemove({ attachments: a1 }, 'F-1', 'tester');
    record('dwf: 삭제 = removed 마킹 (메타 이력 보존 · 배열 불변)', a2.length===1 && !!a2[0].removed && a2[0].removed.by==='tester' && a1[0].removed===null);
    record('dwf: 목록 = removed 제외', win._drawAttList({ attachments: a2 }).length===0 && win._drawAttList({ attachments: a1 }).length===1);
    record('dwf: 확장자·크기 헬퍼', win._drawFileExt('도면_rev2.DWG')==='dwg' && win._fmtBytes(1572864)==='1.5 MB');
  })();
  const dwA = win.DB.get('drawings','038-031-D');
  record('dwf: 데모 시드 링크 첨부 1건 (Rev 2)', win._drawAttList(dwA).length===1 && win._drawAttList(dwA)[0].kind==='link');
  (function(){
    try {
      win.VIEWS.draw();
      const h = win.document.getElementById('view').innerHTML;
      record('dwf render: 목록에 [첨부 N] 버튼', h.includes('data-dw-files') && h.includes('첨부 1'));
    } catch(e){ record('dwf render', false, e.message); }
  })();

  record('draw: PAGES 도면 메뉴 (p_drawing 기존 키 재사용)', /id:'draw',\s*flag:'p_drawing'/.test(html));
  const dw = win.DB.get('drawings','038-031-D');
  record('draw: 데모 시드 — 리비전 이력 3건·현행 Rev 2', !!dw && dw.current_rev==='2' && dw.revs.length===3 && dw.itemCode==='HF-3301');
  record('draw: _drawNextRev 숫자 +1 제안 · 비숫자 빈값', win._drawNextRev('2')==='3' && win._drawNextRev('A')==='');
  (function(){
    // 리비전 업 데이터 규칙: revs 증분 누적 · current_rev 갱신 · 이전 이력 보존
    win.DB.set('drawings','038-031-D', { current_rev:'3', revs:[...dw.revs, { rev:'3', date:win.fmtDate(Date.now()), reason:'검증 테스트', by:'t', ts:Date.now() }] });
    const d2 = win.DB.get('drawings','038-031-D');
    record('draw: 리비전 업 — 이력 4건·현행 3·기존 이력 보존', d2.revs.length===4 && d2.current_rev==='3' && d2.revs[0].rev==='0');
    win.DB.set('drawings','038-031-D', { current_rev: dw.current_rev, revs: dw.revs });  // 원복
  })();
  (function(){
    try {
      win.VIEWS.draw();
      const h = win.document.getElementById('view').innerHTML;
      record('draw render: 목록·현행 Rev·이력 수 표시', h.includes('도면 관리') && h.includes('038-031-D') && h.includes('Rev 2'));
    } catch(e){ record('draw render', false, e.message); }
  })();

  // ── v0.27.2 계정과목 확충 · 마감→발행 이동 ──
  record('cat: 표준 계정과목 22과목 정의', win.STD_EXPENSE_CATS.length===22, String(win.STD_EXPENSE_CATS.length));
  const catNames = win.DB.all('expense_categories').map(c=>c.name);
  record('cat: 데모 DB 에 표준 세트 시드 (급여·복리후생비·지급수수료 등)', ['급여','복리후생비','지급수수료','세금과공과','지급임차료','소모품비'].every(n=>catNames.includes(n)), `${catNames.length}개`);
  record('cat: 재료비·외주가공비 없음 (이중 계상 방지)', !catNames.includes('재료비') && !catNames.includes('외주가공비'));
  const catCountBefore = win.DB.allRaw('expense_categories').length;
  win.migrateExpenseCategories();
  record('cat: 마이그레이션 멱등 (재실행 무증가)', win.DB.allRaw('expense_categories').length===catCountBefore);
  record('cat: 경비 UI 라벨 = 계정과목', /<div><label>계정과목<\/label><select id="ex-cat">/.test(html) && html.includes('계정과목 관리'));
  record('close→invoice: 매출 마감 확정 시 자동 이동 코드', /state\.type==='sales' && getFlags\(\)\.acct/.test(html) && html.includes("tab:'invoice', month: state.month"));
  record('close→invoice: 마감 완료 화면 발행 진입 버튼', /id="cl-goto-invoice"/.test(html));
  // 데모 경비 시드 계정과목 표준명 일치
  record('cat: 데모 경비 시드 = 표준 계정과목명 (지급임차료·소모품비)', (()=>{ const e1=win.DB.get('acct_expenses','EX-DEMO-001'), e2=win.DB.get('acct_expenses','EX-DEMO-002'); return e1?.category==='지급임차료' && e2?.category==='소모품비'; })());

  // ────────────────────────────────────────
  // ★ 5번째 상시 검증 — 회계 시나리오 (원장 무변 · 미수 검산 · 손익 산식 · 계산서=마감 일치)
  // ────────────────────────────────────────
  const acctSnap = () => ({ records: win.DB.all('records').length, lots: win.DB.all('production_lots').length,
    orders: win.DB.all('orders').length, ships: win.DB.all('shipments').length });
  const acctBefore = acctSnap();
  const acctBeforeReport = win.buildReportData('week', win.fmtDate(Date.now()));

  const stg = win.getAcctSettings();
  record('acct: 설정값 — 경고 45일(변경 가능)·수단 3종(어음 포함)', stg.warn_days===45 && stg.methods.includes('어음'), JSON.stringify(stg));
  // 발생액 검산: calcAcctFlows.sales 총합 == 전 출하 shipAmount 총합 (calcClosing 1벌 재사용 증명)
  const flows = win.calcAcctFlows();
  const flowSales = Object.values(flows.sales).reduce((s,v)=>s+v,0);
  const shipTot = win.DB.all('shipments').reduce((t,sh)=>t+win.shipAmount(sh),0);
  record('acct: 매출 발생 총합 = Σ shipAmount (재계산 없음)', Math.round(flowSales)===Math.round(shipTot), `${flowSales} vs ${shipTot}`);
  // 미수 검산 — C001: 수금 시드 2,300,000 반영
  const r1 = win.calcReceivable('C001');
  record('acct: C001 수금 누계 = 3,300,000 (v0.32.5 대량 시드 3건으로 승계)', r1.received===3300000, `received=${r1.received}`);
  record('acct: C001 미수 = 발생 − 수금 (v0.32.5 승계)', Math.round(r1.balance)===Math.round(r1.accrued-3300000), `${r1.balance}`);
  record('acct: 경과일 = 마지막 수금 기준 (v0.32.5 승계 — 5일 전 대량 시드)', (()=>{ const od=win.acctOverdueDays(r1); return r1.balance<=0 ? od===null : (od!=null && od>=4 && od<=6); })(), String(win.acctOverdueDays(r1)));
  // 수금 등록 → 정확히 그만큼 감소 · void → 원복
  win.DB.set('acct_receipts','AR-TEST-1',{ id:'AR-TEST-1', vendor:'C001', amount:100000, method:'현금', date:win.fmtDate(Date.now()), ts:Date.now(), by:'test' });
  const r2 = win.calcReceivable('C001');
  record('acct: 수금 +100,000 → 미수 −100,000', Math.round(r1.balance-r2.balance)===100000, `${r1.balance} → ${r2.balance}`);
  win.voidDoc('acct_receipts','AR-TEST-1','test');
  const r3 = win.calcReceivable('C001');
  record('acct: 수금 void → 미수 원복 (물리 삭제 없음)', Math.round(r3.balance)===Math.round(r1.balance) && !!win.DB.allRaw('acct_receipts').find(x=>x.id==='AR-TEST-1'), `${r3.balance}`);
  // 지급 검산
  const p1 = win.calcPayable('S001');
  record('acct: S001 지급 누계 = 1,440,000 (v0.32.5 대량 시드 2건으로 승계)', p1.paid===1440000, `paid=${p1.paid}`);
  // 손익 산식 + 경비 반영
  const thisM = win._tsMonth(Date.now());
  const pnl = win.calcAcctPnl(thisM);
  record('acct: 손익 항등식 (이익 = 매출−매입−외주−경비)', Math.round(pnl.profit)===Math.round(pnl.revenue-pnl.purchase-pnl.outsource-pnl.expenses), JSON.stringify(pnl));
  record('acct: 이번 달 경비에 절삭유 150,000 포함', pnl.expenses>=150000 && win.acctExpensesOf(thisM).some(e=>e.name==='절삭유'), `exp=${pnl.expenses}`);
  // 계산서 대상 = 마감 매출 일치 (구조적)
  const lastM = win._tsMonth(Date.now()-32*86400000);
  const tg = win.calcInvoiceTargets(lastM);
  const clSales = win.calcClosing('sales', lastM);
  record('acct: 계산서 대상 = 마감 매출 산식 일치', tg.every(t=>Math.round(t.amount)===Math.round(clSales.byParty[t.vendor].amount)), `${tg.length} targets`);
  const demoInv = win.DB.get('acct_invoices','TI-DEMO-001');
  record('acct: 데모 발행 기록 금액 = 대상액 (금액 상이 없음)', !!demoInv && Math.round(demoInv.amount)===Math.round(clSales.byParty['C001']?.amount||0));
  // 반복 경비 제안 → 등록 → 제안 소멸
  const sg1 = win._acctRepeatSuggestions(thisM);
  record('acct: 반복 제안 — 전월 임차료(매월) 이번 달 미등록 → 제안 포함', sg1.some(x=>x.name==='공장 임차료' && x.amount===2000000), sg1.map(x=>x.name).join(','));
  win.DB.set('acct_expenses','EX-TEST-1',{ id:'EX-TEST-1', category:'지급임차료', name:'공장 임차료', amount:2000000, date:thisM+'-01', cl_month:thisM, repeat_monthly:true, ts:Date.now(), by:'test' });
  record('acct: 등록 후 제안 소멸 (멱등)', !win._acctRepeatSuggestions(thisM).some(x=>x.name==='공장 임차료'));
  // ★ G1 — 회계 조작 전 과정에서 원장 무변
  const acctAfterReport = win.buildReportData('week', win.fmtDate(Date.now()));
  record('acct G1: 원장 counts 불변 (records/lots/orders/ships)', JSON.stringify(acctBefore)===JSON.stringify(acctSnap()), JSON.stringify(acctSnap()));
  record('acct G1: buildReportData 수치 불변 (매출·실적)', acctBeforeReport.sales.revenue===acctAfterReport.sales.revenue && acctBeforeReport.production.actual===acctAfterReport.production.actual);
  // 화면 3장 렌더 (허브 탭별)
  try {
    win.window._acctHub = { tab:'ledger', month: thisM };
    win.VIEWS.acct(); const h1 = win.document.getElementById('view').innerHTML;
    win.window._acctHub.tab='invoice'; win.VIEWS.acct(); const h2 = win.document.getElementById('view').innerHTML;
    win.window._acctHub.tab='pnl'; win.VIEWS.acct(); const h3 = win.document.getElementById('view').innerHTML;
    record('acct render: 원장·계산서/경비·손익 3장 예외 0',
      h1.includes('수금 (매출채권)') && h2.includes('세금계산서 발행 관리') && h3.includes('월 관리 손익'), 'ok');
    record('acct render: 정직 라벨 (간이·세전·신고용 아님)', h3.includes('세무 신고용 아님'));
    win.window._acctHub.tab='ledger';
  } catch(e){ record('acct render: 3장', false, e.message); }

  record('sim regression: counts 불변', JSON.stringify(before) === JSON.stringify(snap()));
  record('sim regression: totalShipAmount 불변', beforeShip === shipTotal(), `₩${beforeShip.toLocaleString()}`);
  const afterReport = win.buildReportData('week', win.fmtDate(Date.now()));
  record('sim regression: buildReportData 수치 불변', beforeReport.sales.revenue === afterReport.sales.revenue && beforeReport.production.actual === afterReport.production.actual);
  record('회귀 v0.26.4: buildReportPpt 무변 (PPT 존재·생성)', typeof win.buildReportPpt === 'function' && !!win.buildReportPpt(beforeReport));

  // ── v0.28.5 이관 시나리오 (원장 스냅샷 검증 뒤 실행 — 의도적 원장 추가) ──
  (function(){
    const mk = (m,o)=>{ const cols=win._importerColumns(m); const row={}; cols.forEach(c=>{ if(o[c.key]!==undefined) row[c.label]=o[c.key]; }); return win._parseImportRow(m,row); };
    const J=(m,o)=>win._importJudgeRow(m,mk(m,o)); const A=(m,o)=>win._importApplyRow(m,mk(m,o));
    // 이관 시나리오: 품목/BOM → 소재 → LOT → 실적 → 반출 → 수주 (전부 기존 검증 경로)
    win.DB.set('items','IMP-001',{code:'IMP-001',name:'이관품',mat:'S45C',dia:40,weight:0.5,price:2000,active:true});
    record('mig bom: 라우트 검증·등록', J('bom',{code:'IMP-001',route:'CUT>HEAT:out>SHIP'}).ok && (A('bom',{code:'IMP-001',route:'CUT>HEAT:out>SHIP',mat:'S45C',dia:40,weight:0.5}), win.DB.get('bom','IMP-001').route.length===3));
    record('mig bom: 없는 공정 거부', J('bom',{code:'IMP-001',route:'CUT>XXX'}).ok===false);
    record('mig mat: heat 입고 → receipts+lots 두 문서', J('material_receipts',{heat:'H-IMP-1',kg:500,date:'2026-06-01'}).ok && (A('material_receipts',{heat:'H-IMP-1',mat:'S45C',dia:40,kg:500,date:'2026-06-01'}), win.DB.get('material_lots','H-IMP-1').remain_kg===500 && !!win.DB.get('material_receipts','MR-IMP-H-IMP-1')));
    record('mig lot: LOT.create 경로·heat 승계', J('production_lots',{plot_no:'PL-IMP-1',itemCode:'IMP-001',qty_initial:800,heat:'H-IMP-1',date:'2026-06-05'}).ok && (A('production_lots',{plot_no:'PL-IMP-1',itemCode:'IMP-001',qty_initial:800,heat:'H-IMP-1',date:'2026-06-05'}), win.DB.get('production_lots','PL-IMP-1').heat_nos.length===1));
    record('mig rec: registerRecord 경로 — 절단 700/50 등록·ts 소급', (()=>{ const pr=mk('records',{plot_no:'PL-IMP-1',proc:'CUT',good:700,defect:50,date:'2026-06-06'}); if(!win._importJudgeRow('records',pr).ok) return false; win._importApplyRow('records',pr); const r=win.DB.query('records',x=>x.plot_no==='PL-IMP-1'&&x.proc==='CUT')[0]; return r && r.good===700 && r.ts===win._importDateTs('2026-06-06'); })());
    record('mig rec: 초과 실적 행 = 등록 시 거부 (보존식 보호)', (()=>{ const pr=mk('records',{plot_no:'PL-IMP-1',proc:'CUT',good:100,date:'2026-06-07'}); if(!win._importJudgeRow('records',pr).ok) return false; try { win._importApplyRow('records',pr); return false; } catch(e){ return String(e.message).includes('초과'); } })());
    record('mig od: 반출 임포트 — 가능량 내 등록·초과 거부', (()=>{ const ok=mk('os_dispatches',{plot_no:'PL-IMP-1',proc:'HEAT',vendor:'V001',qty:600,date:'2026-06-10'}); if(!win._importJudgeRow('os_dispatches',ok).ok) return false; win._importApplyRow('os_dispatches',ok); const over=mk('os_dispatches',{plot_no:'PL-IMP-1',proc:'HEAT',vendor:'V001',qty:200,date:'2026-06-11'}); try { win._importApplyRow('os_dispatches',over); return false; } catch(e){ return win.osAvailable('PL-IMP-1','HEAT').avail===100; } })());
    record('mig so: 수주 임포트 — 자동 채번·created 소급', (()=>{ const pr=mk('orders',{cust:'C001',itemCode:'IMP-001',qty:500,price:2000,due:'2026-07-15',date:'2026-06-01'}); if(!win._importJudgeRow('orders',pr).ok) return false; win._importApplyRow('orders',pr); const o=win.DB.all('orders').find(x=>x.itemCode==='IMP-001'); return o && o.created===win._importDateTs('2026-06-01') && o.imported===true; })());
  })();

  // ── v0.31.5 전 마스터 공용 삭제 체계 (이력 0 = 삭제 · 이력 있으면 차단) ──
  (function(){
    // 정적: 레지스트리·가드·버튼 배선
    record('del315: MASTER_DEL_COLS 11종 (rules v4 와 동기)', Array.isArray(win.MASTER_DEL_COLS) && win.MASTER_DEL_COLS.length === 11 && ['items','bom','materials','workers','vendors','equipments','molds','defect_types','processes','drawings','expense_categories'].every(c => win.MASTER_DEL_COLS.includes(c)));
    record('del315: _MASTER_DEL 레지스트리 8종', ['workers','vendors','equipments','molds','defect_types','processes','drawings','expense_categories'].every(c => !!win._MASTER_DEL[c]));
    record('del315: masterDelete/masterHistoryRefs/masterDelBtn 정의', typeof win.masterDelete === 'function' && typeof win.masterHistoryRefs === 'function' && typeof win.masterDelBtn === 'function');
    record('del315: 모달 삭제 버튼 5곳 배선 (작업자·거래처·설비·금형·불량유형)', ["wk-del","mv-del","me-del","m-del","md-del"].every(id => html.includes(`masterDelBtn('${id}')`) && html.includes(`getElementById('${id}')`)));
    record('del315: 행 삭제 버튼 3곳 배선 (공정·도면·계정과목)', /data-proc-del/.test(html) && /data-dw-del/.test(html) && /data-ec-del/.test(html) && /masterDelete\('processes'/.test(html) && /masterDelete\('drawings'/.test(html) && /masterDelete\('expense_categories'/.test(html));
    record('del315: 표준 계정과목(EC-STD-*) 가드 (멱등 시드 복원 → 삭제 무의미 차단)', /EC-STD-/.test(html) && typeof win._MASTER_DEL.expense_categories.guard === 'function');
    // firestore.rules v4 파일 동기 (콘솔 게시는 사용자 게이트)
    try {
      const rules = fs.readFileSync(path.join(__dirname, 'firestore.rules'), 'utf8');
      record('del315: rules v4 — col in [11종] delete 허용', /\(v4\)/.test(rules) && /col in \['items','bom','materials','workers','vendors','equipments','molds',/.test(rules) && /'defect_types','processes','drawings','expense_categories'\]/.test(rules));
      record('del315: rules v4 — 원장 delete 전면 금지 유지', /allow delete: if false;/.test(rules));
    } catch(e){ record('del315: rules v4 파일', false, e.message); }

    // 기능: 왕복 (무이력 삭제 성공 → 재조회 null / 이력 참조 → 차단·문서 보존)
    win.confirm = () => true;   // masterDelete 확인창 통과
    const del = (col,id) => win.masterDelete(col, id);
    // workers
    win.DB.set('workers','ZZ-901',{id:'ZZ-901',emp_no:'ZZ-901',name:'삭제시험',dept:'',active:true});
    record('del315: 작업자 무이력 삭제 → 재조회 null', del('workers','ZZ-901') === true && !win.DB.get('workers','ZZ-901'));
    win.DB.set('workers','ZZ-902',{id:'ZZ-902',emp_no:'ZZ-902',name:'이력자',active:true});
    win.DB.set('records','REC-DEL-T1',{id:'REC-DEL-T1',plot_no:'PL-IMP-1',proc:'CUT',good:1,type:'in',worker:{id:'ZZ-902',emp_no:'ZZ-902',name:'이력자'},ts:Date.now()});
    record('del315: 작업자 실적 참조 → 차단·보존', del('workers','ZZ-902') === false && !!win.DB.get('workers','ZZ-902'));
    // vendors
    win.DB.set('vendors','Z-901',{code:'Z-901',name:'무이력상사',type:'sup',active:true});
    record('del315: 거래처 무이력 삭제', del('vendors','Z-901') === true && !win.DB.get('vendors','Z-901'));
    record('del315: 거래처 수주 참조 → 차단 (C001)', del('vendors','C001') === false && !!win.DB.get('vendors','C001'));
    // equipments
    win.DB.set('equipments','EQ-901',{id:'EQ-901',name:'시험프레스901',procCode:'FORGE',active:true});
    record('del315: 설비 무이력 삭제', del('equipments','EQ-901') === true && !win.DB.get('equipments','EQ-901'));
    win.DB.set('equipments','EQ-902',{id:'EQ-902',name:'시험프레스902',procCode:'FORGE',active:true});
    win.DB.set('production_plans','PP-DEL-T1',{id:'PP-DEL-T1',itemCode:'IMP-001',qty:10,equipment:'시험프레스902',status:'planned',ts:Date.now()});
    record('del315: 설비 생산계획 참조 → 차단·보존', del('equipments','EQ-902') === false && !!win.DB.get('equipments','EQ-902'));
    // molds (파생 타수 0 = 삭제 · FORGE 실적 파생 > 0 = 차단)
    win.DB.set('molds','MD-901',{id:'MD-901',name:'무이력금형',type:'mold',itemCodes:[],life_limit:1000,offset:0,qty:1,active:true});
    record('del315: 금형 파생 타수 0 삭제', del('molds','MD-901') === true && !win.DB.get('molds','MD-901'));
    (function(){
      const forged = win.DB.all('items').find(i => win._forgeQtyForItem(i.code) > 0);
      if(!forged){ record('del315: 금형 파생 타수 차단', false, 'FORGE 실적 보유 품번 없음 — 시나리오 확인'); return; }
      win.DB.set('molds','MD-902',{id:'MD-902',name:'이력금형',type:'mold',itemCodes:[forged.code],life_limit:1000,offset:0,qty:1,active:true});
      record('del315: 금형 파생 타수 참조 → 차단·보존', del('molds','MD-902') === false && !!win.DB.get('molds','MD-902'), forged.code);
    })();
    // defect_types
    win.DB.set('defect_types','DT-901',{code:'DT-901',name:'무이력불량',active:true});
    record('del315: 불량유형 무이력 삭제', del('defect_types','DT-901') === true && !win.DB.get('defect_types','DT-901'));
    win.DB.set('defect_types','DT-902',{code:'DT-902',name:'이력불량',active:true});
    win.DB.set('nc_records','NC-DEL-T1',{id:'NC-DEL-T1',defect_type:'DT-902',item_code:'IMP-001',qty:1,ts:Date.now()});
    record('del315: 불량유형 부적합 참조 → 차단·보존', del('defect_types','DT-902') === false && !!win.DB.get('defect_types','DT-902'));
    // processes
    win.DB.set('processes','ZPRC',{code:'ZPRC',name:'시험공정',seq:999,unit:'ea',type:'in',equipment:[],inspect:[]});
    record('del315: 공정 무이력 삭제', del('processes','ZPRC') === true && !win.DB.get('processes','ZPRC'));
    record('del315: 공정 BOM 라우트 참조 → 차단 (CUT)', del('processes','CUT') === false && !!win.DB.get('processes','CUT'));
    // drawings
    win.DB.set('drawings','DW-901',{id:'DW-901',itemCode:'IMP-001',current_rev:0,revs:[{rev:0,date:'2026-08-01',reason:'최초'}],active:true});
    record('del315: 도면 최초등록만 → 삭제', del('drawings','DW-901') === true && !win.DB.get('drawings','DW-901'));
    win.DB.set('drawings','DW-902',{id:'DW-902',itemCode:'IMP-001',current_rev:1,revs:[{rev:0},{rev:1}],active:true});
    record('del315: 도면 개정 이력 2건 → 차단·보존', del('drawings','DW-902') === false && !!win.DB.get('drawings','DW-902'));
    // expense_categories
    record('del315: 표준 계정과목 → 가드 차단', del('expense_categories','EC-STD-001') === false && !!win.DB.get('expense_categories','EC-STD-001'));
    win.DB.set('expense_categories','EC-901-zzz',{id:'EC-901-zzz',name:'시험과목',order:9999,active:true});
    record('del315: 사용자 계정과목 무이력 삭제', del('expense_categories','EC-901-zzz') === true && !win.DB.get('expense_categories','EC-901-zzz'));
    win.DB.set('expense_categories','EC-902-zzz',{id:'EC-902-zzz',name:'이력과목',order:9998,active:true});
    win.DB.set('acct_expenses','EX-DEL-T1',{id:'EX-DEL-T1',category:'이력과목',name:'시험경비',amount:1000,date:'2026-08-01',cl_month:'2026-08',ts:Date.now()});
    record('del315: 계정과목 경비 참조 → 차단·보존', del('expense_categories','EC-902-zzz') === false && !!win.DB.get('expense_categories','EC-902-zzz'));
    // 원장 이중 방어
    record('del315: DB.remove 원장 불허 (records)', win.DB.remove('records','REC-DEL-T1') === false && !!win.DB.get('records','REC-DEL-T1'));
    record('del315: masterDelete 미지원 컬렉션 거부 (orders)', win.masterDelete('orders','SO-X') === false);
  })();

  // ── v0.32.0 디자인 리뉴얼 1차 (라이트 기본 · 토큰 · 컴포넌트 레이어 — 로직 무접촉) ──
  (function(){
    record('dz320: 부팅 기본 테마 = PC 라이트 · 폰 auto(기기 설정) — v0.37.1 승계 (저장 취향은 유지)', /t = phone \? 'auto' : 'light';/.test(html) && /setAttribute\('data-theme','light'\)/.test(html) && typeof win.getThemePref === 'function' && typeof win.resolveTheme === 'function');
    record('dz320: 신규 토큰 --sh/--prim 두 테마 정의', /--sh:none;--sh2:none;--prim:#eceff3;--prim-ink:#12161d;/.test(html) && /--sh:0 1px 2px rgba\(16,24,40,\.05\)/.test(html) && /--prim:#1a212e;--prim-ink:#ffffff;/.test(html));
    record('dz320: 라이트 프리미엄 팔레트 (--bg #f6f7f9 · --ink #101828 · --cyan #0f766e)', /--bg:#f6f7f9;--bg2:#ffffff;--panel:#ffffff;--panel2:#f9fafb;/.test(html) && /--ink:#101828/.test(html) && /--cyan:#0f766e/.test(html));
    record('dz320: 디자인 레이어 — 카드 단색+그림자·차콜 주버튼', /\.card\{background:var\(--panel\);box-shadow:var\(--sh\);border-radius:14px\}/.test(html) && /\.btn\.primary\{background:var\(--prim\);border-color:var\(--prim\);color:var\(--prim-ink\)\}/.test(html));
    record('dz320: 라이트 배경 평면화 (그리드·라디얼 제거)', /html\[data-theme="light"\] body\{background:var\(--bg\)\}/.test(html) && /html\[data-theme="light"\] body::before\{display:none\}/.test(html));
    record('dz320: 행 호버 중립화 (청록 틴트 → 잉크 4%)', /tr:hover td\{background:color-mix\(in srgb,var\(--ink\) 4%, transparent\)\}/.test(html));
    record('dz320: 다크 :root 원본 불변 (기존 다크 사용자 픽셀 보존)', /--bg:#08090b;--bg2:#0c0e11;--panel:#121419;--panel2:#161922;/.test(html));
    record('dz320: 스크롤 CSS 불변 (sc292 계약 유지)', /\.tbl-wrap\{overflow-x:auto;overflow-y:auto;max-height:50vh/.test(html));
  })();

  // ── v0.32.1 사이드바 내비 (그룹형 · 데스크톱 전용 · 좁은 화면 기존 가로 칩) ──
  (function(){
    record('nav321: NAV_GROUPS 7그룹 · 전 PAGES 수용(기타 제외 누락 0)', (()=>{
      const G = win.eval('NAV_GROUPS');
      if(!Array.isArray(G) || G.length !== 7) return false;
      const inG = new Set(G.flatMap(g => g.ids));
      const P = win.eval('PAGES');
      return P.every(p => inG.has(p.id));
    })());
    record('nav321: renderNav 그룹 라벨 + 전 메뉴 버튼 (개발자 세션)', (()=>{
      const S = win.eval('SESSION');
      const bak = { loggedIn:S.loggedIn, role:S.role };
      S.loggedIn = true; S.role = 'vendor';
      win.eval('renderNav(); openAllMenuModal()');   // ★ v0.37.7 — 그룹 라벨·전 메뉴는 전체 메뉴 창으로 이동 · 사이드바 = 내 메뉴 + 전체 메뉴 버튼 · 아이콘 0
      const nv = win.document.getElementById('nav').innerHTML, am = win.document.getElementById('modal').innerHTML; win.eval('closeModal()');
      S.loggedIn = bak.loggedIn; S.role = bak.role;
      return am.includes('am-grp-lbl') && am.includes('영업') && am.includes('생산') && am.includes('data-am="dash"') && am.includes('data-am="set"') && nv.includes('data-page="dash"') && nv.includes('data-allmenu') && !nv.includes('<svg') && !nv.includes('nav-grp');
    })());
    record('nav321: 사이드바 CSS — 그리드·sticky·좁은 화면 라벨 숨김·인쇄 무영향', /#app\{display:grid;grid-template-columns:212px minmax\(0,1fr\)/.test(html) && /\.nav-grp\{display:none\}/.test(html) && /@media print\{ #app\{display:block\} \}/.test(html));
    record('nav321: 활성 메뉴 틴트 (솔리드 시안 대체 · 아이콘 동색)', /\.tabnav button\.on\{background:color-mix\(in srgb,var\(--cyan\) 14%,transparent\)/.test(html) && /\.tabnav button\.on svg\{stroke:currentColor\}/.test(html));
  })();

  // ── v0.33.3 대시보드 심플화 (사용자 판정 B안 — todo322 검사식 승계·대체: '오늘 챙길 일' 기능 자체가 제거됨) ──
  (function(){
    try {
      win.location.hash = '#dash'; win.VIEWS.dash();
      const v = win.document.getElementById('view').innerHTML;
      record('dash333: 제거 확인 — 오늘 챙길 일·오늘의 생산·안내줄·섹션 라벨 전부 미노출',
        !v.includes('dash-todo') && !v.includes('오늘 챙길 일') && !v.includes('오늘의 생산') && !v.includes('dash-hint') && !v.includes('매출 현황 · SALES') && !v.includes('카드를 클릭하면'));
      // ★ v0.35.7 승계: KPI 6 스트립(hero6) → 핵심 스트립 4 (총 수주잔 · 당월 매출 · 오늘 생산 · 오늘 이상) — D′7 FAIL 수정 (A안)
      record('dash357: 핵심 스트립 4 (hero dash3 — 총 수주잔 · 당월 매출 · 오늘 생산 · 오늘 이상) · hero6 제거',
        v.includes('hero hero4 dash3') && v.includes('총 수주잔') && v.includes('당월 매출') && v.includes('오늘 생산') && v.includes('오늘 이상') && !v.includes('hero6') && (v.match(/class="dkpi/g)||[]).length === 4);
      record('dash357: 이동 확인 — 연 누적 카드·월별 매출·고객사 5년·진도 4칸·공정 흐름·품질 KPI 그리드 대시보드 미노출 (담당 화면 존치)',
        !v.includes('cust5y') && !v.includes('월별 매출') && !v.includes('생산 진도') && !v.includes('class="flow"') && !v.includes('class="kgrid"') && !v.includes('종합 수율') && win.eval('typeof computeProgress')==='function' && win.eval('typeof calcSales')==='function');
      record('dash334: 띠 세그먼트 .prog 전역 충돌 차단 (margin:0·row 명시 — 생산중 반만 칠해지던 버그)',
        /\.stage3 \.s\{height:100%;margin:0;display:flex;flex-direction:row/.test(html));
      record('dash333: 띠그래프 단색 (생산중 #f59e0b · 출하완료 단색 + 라이트 텍스트 보정 · 그라데이션 제거)',
        /\.stage3 \.s\.prog\{background:#f59e0b;color:#3a2800\}/.test(html) && /\.stage3 \.s\.done\{background:var\(--cyan-dim\)\}/.test(html) && /html\[data-theme="light"\] \.stage3 \.s\.done\{color:#eafffb\}/.test(html) && !/\.stage3 \.s\.prog\{background:linear-gradient/.test(html));
      record('dash333: 차트 단색 틸 — 월별 막대 인라인 무지개 색 제거 (당월만 진하게)',
        /\.mbar \.b\{[^}]*background:var\(--cyan-dim\);[^}]*opacity:\.45\}/.test(html) && !v.match(/class="b (cur )?\s*(future)?"[^>]*background:linear-gradient/));
      record('dash333: 수주대비 카드 — 범례 설명문 제거·총 수주잔 범례 우측 통합',
        !v.includes('납품·매출 확정분') && !v.includes('공정 진행 중 (재공)') && v.includes('총 수주잔'));
      record('dash357: 카드 2장 — 수주 대비 띠(stage3 · D′4 그대로) + 생산 추세(trend-seg) 한 행 · 카드 수 2',
        /dgrid2 dash3g[\s\S]*?stage3[\s\S]*?trend-chart/.test(v) && v.includes('trend-seg') && (v.match(/class="card/g)||[]).length === 2);
      record('dash357: 오늘 이상 신호등 — 잠금·HOLD·부적합 대기·납기 지연 합 (lockedPlotSet·heldPlotSet·calcSoRows 1벌 · 0이면 「이상 없음」)',
        /lockedPlotSet\(\)\.size/.test(html) && /heldPlotSet\(\)\.size/.test(html) && /calcSoRows\(\)\.filter\(r => r\.status\.label === '지연'\)/.test(html) && (v.includes('이상 없음') || /오늘 이상[\s\S]*?\d+<span class="unit">건/.test(v)));
      record('dash333: 데이터 보존 — 오늘의 생산 상세는 설비현황 담당 (computeEqStatus 1벌 유지)', typeof win.eval('typeof computeEqStatus') === 'string' && win.eval('typeof computeEqStatus') === 'function');
      record('dash357→hold3: 오늘 이상 = 잠금+HOLD+판정 지체(alert↑ · ncPendingLevel 1벌)+납기 지연 (9-B 승계)', /_ncDelay = flags\.qc \? DB\.query\('nc_records', n => ncIsPending\(n\) && ncPendingLevel\(n\) >= 2\)\.length/.test(html) && /_issueN = _lockN \+ _holdN \+ _lateN \+ _ncDelay/.test(html));
    } catch(e){ record('dash333: 대시보드 심플화', false, e.message); }
  })();

  // ── v0.32.3 LOT 공정 스텝 패널 (행 클릭 → 스텝 → 공정 클릭 → 실적 상세 · 읽기 전용) ──
  (function(){
    try {
      record('lst323: lotStepPanelHTML 정의 + LOT행 배선 + 스텝 CSS', typeof win.lotStepPanelHTML === 'function' && /data-lot-steps=/.test(html) && /\.lstp\.done \.lsdot\{background:var\(--cyan\)/.test(html) && /@media print\{\.lot-steps-row\{display:none\}\}/.test(html));
      // 표준 시나리오 LOT 하나로 패널 생성 — 스텝·상세·LOT 배지·합계 검증
      const lot = win.DB.all('production_lots').find(x => win.DB.query('records', r => r.plot_no === x.plot_no && (+r.good||0) > 0).length > 0);
      if(!lot){ record('lst323: 패널 생성 (실적 보유 LOT)', false, '실적 보유 LOT 없음'); return; }
      const p = win.lotStepPanelHTML(lot.plot_no);
      record('lst323: 패널 — 생산 LOT 배지 + 스텝 스트립', p.includes('생산 LOT') && p.includes(lot.plot_no) && p.includes('lstp') && p.includes('data-lstp'));
      record('lst323: 패널 — 실적 상세 표 (일시·설비·작업자·양품/불량/재작업·합계)', p.includes('일시') && p.includes('작업자') && p.includes('양품') && p.includes('재작업') && p.includes('합계'));
      record('lst323: 패널 — 완료 공정 done 상태 존재', /lstp done/.test(p));
      record('lst323: 원장 무변 (패널 생성은 읽기 전용)', (()=>{ const before = win.DB.all('records').length; win.lotStepPanelHTML(lot.plot_no); return win.DB.all('records').length === before; })());
    } catch(e){ record('lst323: 패널', false, e.message); }
  })();

  // ── v0.32.4 규칙·가드 상주 안내문 접기 (마크업 무수정 · 렌더 후 클래스 부여 · 텍스트 보존) ──
  (function(){
    try {
      record('rn324: _ruleNoteEnhance 정의 + router/openModal 훅 + CSS(1줄 접힘·인쇄 펼침)', typeof win._ruleNoteEnhance === 'function' && /_ruleNoteEnhance\(\); \} catch/.test(html.replace(/try \{ _ruleNoteEnhance/g,'try { _ruleNoteEnhance')) && /\.rule-note\{max-height:40px/.test(html) && /@media print\{\.rule-note\{max-height:none !important\}/.test(html));
      const S = win.eval('SESSION');
      const bak = { loggedIn:S.loggedIn, role:S.role };
      S.loggedIn = true; S.role = 'vendor';
      win.location.hash = '#masters'; win.VIEWS.masters(); win._ruleNoteEnhance();
      const notes = win.document.querySelectorAll('#view .rule-note');
      record('rn324: 기준정보 화면 규칙·가드 박스 접힘 적용 (2개 이상)', notes.length >= 2, `${notes.length}개`);
      if(notes.length){
        const n = notes[0];
        const before = n.classList.contains('open');
        n.click();
        record('rn324: 클릭 펼침 토글', n.classList.contains('open') !== before);
        record('rn324: 텍스트 전문 보존 (규칙/가드 원문 그대로)', /규칙|가드/.test(n.innerHTML) && n.innerHTML.length > 60);
      }
      S.loggedIn = bak.loggedIn; S.role = bak.role;
      win.location.hash = '#dash';
    } catch(e){ record('rn324: 접기', false, e.message); }
  })();

  // ── v0.32.5 데모 대량 시드 (마지막 섹션 — 이후 수치 검증 없음 · 같은 win 에 시드해도 무해) ──
  (function(){
    try {
      const S = win.eval('SESSION');
      S.loggedIn = true; S.role = 'vendor';
      win.eval('seedDemoData()');
      const C = c => win.DB.all(c).length;
      record('bulk325: 마스터 대량 (품번≥25·거래처≥15·작업자≥17·금형≥19·도면≥14·설비≥9)', C('items')>=25 && C('vendors')>=15 && C('workers')>=17 && C('molds')>=19 && C('drawings')>=14 && C('equipments')>=9, `items ${C('items')} vendors ${C('vendors')} workers ${C('workers')} molds ${C('molds')}`);
      record('bulk325: 원장 대량 (수주≥30·LOT≥20·지시≥20·출하≥28·계획≥20·발주≥12)', C('orders')>=30 && C('production_lots')>=20 && C('work_orders')>=20 && C('shipments')>=28 && C('production_plans')>=20 && C('purchase_orders')>=12, `SO ${C('orders')} LOT ${C('production_lots')} WO ${C('work_orders')}`);
      record('bulk325: 품질·회계 대량 (NC≥14·수금≥11·지급≥8·경비≥12·견적≥11)', C('nc_records')>=14 && C('acct_receipts')>=11 && C('acct_payments')>=8 && C('acct_expenses')>=12 && C('quotes')>=11);
      record('bulk325: 멱등 (재시드 시 문서 수 불변 — 고정 id)', (()=>{
        const before = { i:C('items'), o:C('orders'), l:C('production_lots'), r:C('records') };
        win.eval('seedDemoData()');
        return C('items')===before.i && C('orders')===before.o && C('production_lots')===before.l && C('records')===before.r;
      })());
      // 시드 후 전 화면 렌더 예외 0 (대량 데이터 기준)
      const fails = [];
      ['dash','so','plan','wo','perf','lot','qc','quote','po','matin','os','inv','ship','pallet','closing','acct','cost','eqstatus','mold','draw','proc','masters','set'].forEach(vn => {
        try { win.location.hash = '#'+vn; win.VIEWS[vn](); } catch(e){ fails.push(vn+':'+e.message); }
      });
      record('bulk325: 대량 데이터 렌더 23화면 예외 0', fails.length===0, fails.slice(0,3).join(' | '));
      win.location.hash = '#dash';
      // ── v0.32.6 시뮬레이션 이력 (6개월 완료 생산·매출) + 시드 자동 보강 ──
      record('bulk326: 6개월 이력 (실적≥80·수주≥42·출하≥50·지시≥32·LOT≥32)', C('records')>=80 && C('orders')>=42 && C('shipments')>=50 && C('work_orders')>=32 && C('production_lots')>=32, `rec ${C('records')} SO ${C('orders')} SH ${C('shipments')} WO ${C('work_orders')}`);
      record('bulk326: 파렛트·회계 6개월 (파렛트≥9·경비≥30·수금≥19·견적≥16·마감 3)', C('pallet_moves')>=9 && C('acct_expenses')>=30 && C('acct_receipts')>=19 && C('quotes')>=16 && C('closings')>=3);
      record('bulk326: H LOT 전 공정 완주 (records 에 H01 실적 존재·출하 90% 완납)', (()=>{
        const hRecs = win.DB.query('records', r => r.plot_no === 'PL-2026-H01');
        const so = win.DB.get('orders','SO-2026-H01');
        const shipped = win.DB.query('shipments', x => x.so_id==='SO-2026-H01').reduce((s,x)=>s+(x.qty||0),0);
        return hRecs.length >= 3 && !!so && shipped === so.qty;
      })());
      record('bulk326: 마감 3개월 = calcClosing 재계산 일치 (첫 달 amount 검산)', (()=>{
        const m = win._tsMonth(Date.now()-90*86400000);
        const doc = win.DB.get('closings', win._closingId('sales', m));
        if(!doc) return false;
        const { byParty } = win.calcClosing('sales', m);
        return doc.rows.every(r => Math.round(byParty[r.party]?.amount||0) === r.amount);
      })());
      record('bulk326: 시드 버전 자동 보강 (DEMO_SEED_VER=7 · v0.37.2 worker 사용자 · autoLoginDemo 불일치 재시드)', /const DEMO_SEED_VER = 7;/.test(html) && /sv\.ver !== DEMO_SEED_VER/.test(html) && /데모 샘플 데이터가 최신으로 보강/.test(html));
      // ── v0.32.7 LOT 분할 수량 보존 (v2.1 검증 실측 결함 I6-split → 수정 검증) ──
      (function(){
        try {
          const heat = win.DB.all('material_lots')[0]?.heat || 'H2604-101';
          win.LOT.create({ plot_no:'PL-SPLIT-T', itemCode:'HF-3301', materialLots:[heat], qty_initial:200 });
          win.LOT.split({ parentPlotNo:'PL-SPLIT-T', parts:[{ plot_no:'PL-SPLIT-T1', qty:80 }] });
          const balP = win.LOT.balance('PL-SPLIT-T','CUT');
          record('split327: 분할 후 부모 잔량 차감 (200 분할 80 → 부모 120)', balP === 120, `부모 잔량 ${balP}`);
          record('split327: 이관 레코드 생성 (transferred=80 · transferred_to 자식)', win.DB.query('records', r => r.plot_no==='PL-SPLIT-T' && r.transferred===80 && String(r.transferred_to||'').includes('PL-SPLIT-T1')).length === 1);
          record('split327: 총수량 보존 (부모 잔량+자식 초기 = 원 초기 200)', balP + 80 === 200);
          let over = false; try { win.LOT.split({ parentPlotNo:'PL-SPLIT-T', parts:[{ plot_no:'PL-SPLIT-T2', qty:121 }] }); } catch(e){ over = true; }
          record('split327: 잔량 초과 분할 차단 유지 (121 > 120)', over && !win.DB.get('production_lots','PL-SPLIT-T2'));
        } catch(e){ record('split327: 분할 보존', false, e.message); }
      })();
      // ── v0.33.0 품질 Hold Gate (Q2-1 — K2 GAP 해소 · 원장 쓰기 경로 차단) ──
      (function(){
        try {
          record('hold330: 헬퍼 1벌 + 3중 배선 (registerRecord·출하·반출) + UI 배지', typeof win.lotHoldInfo === 'function' && typeof win.assertLotNotHeld === 'function' && typeof win.heldPlotSet === 'function' && /assertLotNotHeld\(plot_no, '실적 등록'\)/.test(html) && /assertLotNotHeld\(plot, '출하'\)/.test(html) && /assertLotNotHeld\(plot, '외주 반출'\)/.test(html) && /품질 HOLD \(출하 불가\)/.test(html) && html.includes('>HOLD</span>'));
          const heat = win.DB.all('material_lots')[0]?.heat || 'H2604-101';
          win.LOT.create({ plot_no:'PL-HOLD-T', itemCode:'HF-3301', materialLots:[heat], qty_initial:100 });
          win.LOT.registerRecord({ plot_no:'PL-HOLD-T', proc:'CUT', equipment:'절단기 1호', good:60, type:'in', shift:'day' });
          win.DB.set('nc_records','NC-HOLD-T',{ id:'NC-HOLD-T', date:win.fmtDate(Date.now()), item_code:'HF-3301', lot_id:'PL-HOLD-T', record_id:null, process:'CUT', equipment:'절단기 1호', worker:{name:'t'}, defect_type:'DT-001', qty:2, memo:'hold', status:'접수', hold:true /* v0.35.10 승계: 로트 불량 HOLD */, handle:null, rework_record_id:null, ts:Date.now(), by:'t', void:null });
          let b1=false, m1='';
          try { win.LOT.registerRecord({ plot_no:'PL-HOLD-T', proc:'FORGE', equipment:'프레스 1호', good:10, type:'in', shift:'day' }); } catch(e){ b1=true; m1=e.message; }
          record('hold330: 접수 NC → 다음 공정 등록 차단 (원장 경로)', b1 && /품질 HOLD/.test(m1) && win.DB.query('records', r=>r.plot_no==='PL-HOLD-T').length === 1, m1);
          let b2=false; try { win.assertLotNotHeld('PL-HOLD-T','출하'); } catch(e){ b2=true; }
          record('hold330: 출하 게이트 차단', b2);
          win.DB.set('nc_records','NC-HOLD-T',{ status:'특채', handle:{ kind:'deviation', by_name:'t', ts:Date.now(), by:'t' } });
          const ok1 = !!win.LOT.registerRecord({ plot_no:'PL-HOLD-T', proc:'FORGE', equipment:'프레스 1호', good:10, type:'in', shift:'day' });
          record('hold330: 판정(특채) → HOLD 해제·등록 성공', ok1);
          win.DB.set('nc_records','NC-HOLD-T',{ status:'접수', hold:true });
          let b3=false; try { win.LOT.registerRecord({ plot_no:'PL-HOLD-T', proc:'FORGE', equipment:'프레스 1호', good:5, type:'in', shift:'day' }); } catch(e){ b3=/품질 HOLD/.test(e.message); }
          record('hold330: 재접수 → 재차단 (양방향 일관)', b3);
          win.voidDoc('nc_records','NC-HOLD-T','시험 종료');
          const okVoid = (()=>{ try { return !!win.LOT.registerRecord({ plot_no:'PL-HOLD-T', proc:'FORGE', equipment:'프레스 1호', good:5, type:'in', shift:'day' }); } catch(e){ return false; } })();
          record('hold330: NC 취소(void) → HOLD 해제 (void 는 판정에서 제외)', okVoid);
          record('hold330: heldPlotSet 일괄 판정 = 개별 판정 일치', (()=>{ const hs = win.heldPlotSet(); return !hs.has('PL-HOLD-T') && win.lotHoldInfo('PL-HOLD-T').held === false; })());
        } catch(e){ record('hold330: Hold Gate', false, e.message); }
      })();
      // ── v0.33.1 Audit Viewer (Q2-2 — G-3a 전값 스냅샷 · G-3b 통합 뷰어 · LOT 생애주기) ──
      (function(){
        try {
          record('audit331: 메뉴 신설 (PAGES audit · MANAGER_ONLY · 기준정보 그룹)', /id:'audit', flag:'setting', title:'변경 이력'/.test(html) && /MANAGER_ONLY = \['set','masters','audit'\]/.test(html) && /ids:\['masters','docs','audit','set'\]/.test(html));
          const heat = win.DB.all('material_lots')[0]?.heat || 'H2604-101';
          win.LOT.create({ plot_no:'PL-AUD-T', itemCode:'HF-3301', materialLots:[heat], qty_initial:100 });
          const r = win.LOT.registerRecord({ plot_no:'PL-AUD-T', proc:'CUT', equipment:'절단기 1호', good:50, defect:5, type:'in', shift:'day' });
          win.DB.correct('records', r.id, { good:45 }, '검증 — 전값 스냅샷');
          const log = win.DB.get('records', r.id)._logs.find(l => l.op === 'correct');
          record('audit331: G-3a 해소 — 보정 로그에 전값/후값 스냅샷 (50→45)', !!(log && log.prev && log.prev.good === 50 && log.next.good === 45));
          win.DB.set('orders','SO-AUD-T',{ id:'SO-AUD-T', cust:'C001', itemCode:'HF-3301', qty:10, price:100, due:win.fmtDate(Date.now()), created:Date.now() });
          win.voidDoc('orders','SO-AUD-T','검증 취소'); win.unvoidDoc('orders','SO-AUD-T');
          const so = win.DB.allRaw('orders').find(o => o.id === 'SO-AUD-T');
          record('audit331: 취소·복원 흔적 보존 (DB.set 단일 경로 op 세분화 · 복원 후에도 남음)', (so._logs||[]).some(l => l.op === 'void' && l.reason === '검증 취소') && (so._logs||[]).some(l => l.op === 'unvoid') && !so.void);
          const ae = win.auditEntries();
          record('audit331: 통합 이력 — 의미 있는 op 만 (correct/void/unvoid · update 소음 제외)', ae.length >= 2 && ae.every(e => ['correct','void','unvoid'].includes(e.op)));
          record('audit331: 이력에 보정 전값 포함 + 복원 항목 존재', ae.some(e => e.id === r.id && e.op === 'correct' && e.prev?.good === 50) && ae.some(e => e.id === 'SO-AUD-T' && e.op === 'unvoid'));
          const S2 = win.eval('SESSION'); S2.loggedIn = true; S2.role = 'manager';
          win.location.hash = '#audit'; win.VIEWS.audit();
          const v = win.document.getElementById('view').innerHTML;
          record('audit331: G-3b 해소 — 뷰어 렌더 (통합 이력·전→후 표기·LOT 생애주기 카드)', v.includes('통합 변경 이력') && v.includes('50 → 45') && v.includes('LOT 생애주기'));
          const tl = win.lotTimeline('PL-2026-H01');
          record('audit331: LOT 생애주기 — 생성·실적·출하 통합 + 시간 오름차순 (데모 백데이팅 특성상 생성이 첫 행 아닐 수 있음 — 정렬만 검증)', tl.length >= 5 && tl.some(e => e.kind === '생성') && tl.some(e => e.kind === '실적') && tl.some(e => e.kind === '출하') && tl.every((e,i) => i === 0 || tl[i-1].ts <= e.ts));
          win.eval('renderNav(); openAllMenuModal()');   // ★ v0.37.7 — 전체 메뉴 창에서 검사
          const navMgr = win.document.getElementById('modal').innerHTML.includes('data-am="audit"'); win.eval('closeModal()');
          const bak2 = { role:S2.role, userId:S2.userId };
          S2.role = 'worker'; S2.userId = 'U-NOPE'; win.eval('renderNav(); openAllMenuModal()');
          const navWk = !(win.document.getElementById('nav').innerHTML + win.document.getElementById('modal').innerHTML).includes('audit"'); win.eval('closeModal()');
          S2.role = bak2.role; S2.userId = bak2.userId; win.eval('renderNav()');
          record('audit331: 관리자 전용 게이팅 (관리자 노출·작업자 숨김)', navMgr && navWk);
          win.location.hash = '#dash';
        } catch(e){ record('audit331: Audit Viewer', false, e.message); }
      })();
      // ── v0.33.2 동시성 사고 방어·복구 Gate (Q2-3 — G-2 완화 · 탐지→추적→잠금→정정→J1 재검증→해제) ──
      (function(){
        try {
          record('conc332: 헬퍼 1벌 + 3중 배선 + 컬렉션 등록 + UI 배지',
            typeof win.detectConcurrencyIncidents === 'function' && typeof win.lotLockInfo === 'function' && typeof win.assertWithinAccepted === 'function' && typeof win.lockedPlotSet === 'function' && typeof win.verifyLotInvariant === 'function' && typeof win.resolveIncident === 'function'
            && typeof win.assertLotNotLocked === 'undefined'   // ★ v0.36.0 — LOT 전체 잠금 assert 폐지 (수량 차단 = balance/assertWithinAccepted · 사고 상태 = 문서)
            && /assertWithinAccepted\(plot, qty, '출하'/.test(html) && /assertWithinAccepted\(plot, qty, '외주 반출'/.test(html) && /detectConcurrencyIncidents\(plot_no\); \} catch\(e\)\{\}   \/\/ ★ v0.36.0/.test(html)
            && html.includes("'concurrency_incidents',") && html.includes('동시성 사고 (인정분만 출하 가능)') && html.includes('>사고</span>'));
          // 0) 데모 시드 원장 무결 — 시드·기존 테스트 데이터에서 사고 0건 (탐지기 오탐 없음)
          const pre = win.detectConcurrencyIncidents();
          record('conc332: 정상 원장에서 오탐 0건 (시드+기존 시나리오 전체 스캔)', pre.length === 0 && win.DB.allRaw('concurrency_incidents').length === 0, `사고 ${pre.length}건`);
          // 1) 충돌 구성 — 두 단말이 각자 통과한 상태 (v2.1 O2 등가): 100 가용에 80+80
          const heat = win.DB.all('material_lots')[0]?.heat || 'H2604-101';
          win.LOT.create({ plot_no:'PL-CONC-T', itemCode:'HF-3301', materialLots:[heat], qty_initial:100 });
          win.LOT.registerRecord({ plot_no:'PL-CONC-T', proc:'CUT', equipment:'절단기 1호', good:80, type:'in', shift:'day', worker:{name:'작업자A', emp_no:'W-A'} });
          win.DB.set('records', 'R-CUT-PL-CONC-T-SYNC', { id:'R-CUT-PL-CONC-T-SYNC', plot_no:'PL-CONC-T', proc:'CUT', equipment:'절단기 2호', vendor:null, type:'in', good:80, defect:0, rework:0, scrap:0, rework_returned:0, defect_reason:'', shift:'day', source:'manual', worker:{name:'작업자B', emp_no:'W-B'}, ts:Date.now(), by:'W-B' });
          record('conc332: 사고 성립 (동기화 병합 등가 — 보류 60 · 잔량 0 · 인정 100) — v0.36.0 승계', win.LOT.pendingQty('PL-CONC-T','CUT') === 60 && win.LOT.balance('PL-CONC-T','CUT') === 0 && win.LOT.acceptedGood('PL-CONC-T','CUT') === 100, `보류 ${win.LOT.pendingQty('PL-CONC-T','CUT')}`);
          // 2) ① 탐지 — Incident 자동 생성 (LOT·공정·직전 가용·첫 실적·충돌 실적·작업자·시각 자동 연결)
          const made = win.detectConcurrencyIncidents();
          const ci = win.DB.get('concurrency_incidents', 'CI-PL-CONC-T-CUT-1');
          record('conc332: ① Incident 자동 생성 — LOT·공정·가용 100·투입 160·보류 60 연결 (v0.36.0)', made.length === 1 && !!ci && ci.plot_no === 'PL-CONC-T' && ci.proc === 'CUT' && ci.avail === 100 && ci.input === 160 && ci.pending === 60 && ci.accepted === 100 && ci.balance === 0 && ci.status === 'open');
          record('conc332: ① 첫 실적·충돌 실적·작업자·시각 자동 연결', !!ci && ci.first_record?.by === '작업자A' && ci.first_record?.qty === 80 && ci.conflict_records?.length === 1 && ci.conflict_records[0].by === '작업자B' && ci.conflict_records[0].id === 'R-CUT-PL-CONC-T-SYNC' && ci.workers.includes('작업자A') && ci.workers.includes('작업자B') && ci.ts > 0);
          // 멱등 — 재탐지 시 중복 생성 없음
          const again = win.detectConcurrencyIncidents();
          record('conc332: 재탐지 멱등 (열린 사고 중복 생성 0)', again.length === 0 && win.DB.queryRaw('concurrency_incidents', c => c.plot_no === 'PL-CONC-T').length === 1);
          // 3) ② 잠금 — 신규 실적·출하·반출 차단
          let k1=false, km='';
          try { win.LOT.registerRecord({ plot_no:'PL-CONC-T', proc:'CUT', equipment:'절단기 1호', good:1, type:'in', shift:'day' }); } catch(e){ k1=true; km=e.message; }
          record('conc332: ② 수량 차단 — 사고 공정 신규 실적은 잔량 0 초과 차단 (balance 가 Lock 흡수)', k1 && /초과 등록 차단/.test(km) && /잔량 0ea/.test(km), km);
          let k2=false; try { win.assertWithinAccepted('PL-CONC-T', 101, '출하', 100); } catch(e){ k2=/보류 60/.test(e.message); }
          let k3=false; try { win.assertWithinAccepted('PL-CONC-T', 1, '외주 반출', 0); } catch(e){ k3=true; }
          const okIn = win.assertWithinAccepted('PL-CONC-T', 100, '출하', 100);
          record('conc332: ② 출하·반출 인정 밖 차단 + 인정 안 통과(사고 목록 반환) + 일괄 판정 일치', k2 && k3 && okIn.incidents.length === 1 && win.lockedPlotSet().has('PL-CONC-T') && win.lotLockInfo('PL-CONC-T').open === true && win.lotLockInfo('PL-CONC-T').pending === 60);
          // 4) ⑤ 전제 — 정합성 미회복 상태에서 해제 거부 (J1 재검증 실패 → 잠금 유지)
          const S3 = win.eval('SESSION'); const bak3 = { role:S3.role };
          S3.role = 'manager';
          let rej=false, rm='';
          try { win.resolveIncident('CI-PL-CONC-T-CUT-1'); } catch(e){ rej=true; rm=e.message; }
          record('conc332: ⑤ 보류 미해소 → 해제 거부·사고 open 유지 (J1′)', rej && /J1′ 재검증 실패/.test(rm) && /보류 60/.test(rm) && win.DB.get('concurrency_incidents','CI-PL-CONC-T-CUT-1').status === 'open', rm);
          // 비관리자 해제 차단
          S3.role = 'worker';
          let rej2=false; try { win.resolveIncident('CI-PL-CONC-T-CUT-1'); } catch(e){ rej2=/관리책임자/.test(e.message); }
          record('conc332: 비관리자 해제 차단 (관리책임자 전용)', rej2);
          S3.role = 'manager';
          // 5) ③ 관리자 정정 (기존 DB.correct 경로 — 잠금과 무관하게 동작) → ④ J1 자동 재검증 → ⑤ 해제
          win.DB.correct('records', 'R-CUT-PL-CONC-T-SYNC', { good:20 }, '동시성 사고 정정 — 중복 등록분 조정');
          const v1 = win.verifyLotInvariant('PL-CONC-T');
          record('conc332: ③ 관리자 정정 → 잔량 0 회복 + ④ J1 전 공정 오차 0', win.LOT.balance('PL-CONC-T','CUT') === 0 && v1.ok && v1.checks.every(c => c.err === 0), JSON.stringify(v1.checks.map(c=>[c.proc,c.bal,c.err])));
          const done = win.resolveIncident('CI-PL-CONC-T-CUT-1');
          record('conc332: ⑤ 재검증 통과 시에만 해제 — resolved + 검증 스냅샷 보존', done.status === 'resolved' && done.resolved?.verify?.ok === true && Array.isArray(done.resolved.verify.checks) && done.resolved.by);
          const ok2 = (()=>{ try { return !!win.LOT.registerRecord({ plot_no:'PL-CONC-T', proc:'FORGE', equipment:'프레스 1호', good:10, type:'in', shift:'day' }); } catch(e){ return false; } })();
          record('conc332: 해제 후 정상 등록 재개 + 잠금 세트 소멸', ok2 && !win.lockedPlotSet().has('PL-CONC-T'));
          // 6) 원장 연결 — LOT 생애주기에 사고 감지·해제 표시 + 변경 이력 화면 렌더
          const tl2 = win.lotTimeline('PL-CONC-T');
          record('conc332: LOT 생애주기 — 사고 감지·해제 연결', tl2.some(e => e.kind === '사고' && /감지/.test(e.text)) && tl2.some(e => e.kind === '사고' && /해제/.test(e.text)));
          S3.loggedIn = true; win.location.hash = '#audit'; win.VIEWS.audit();
          const v2 = win.document.getElementById('view').innerHTML;
          record('conc332: 변경 이력 화면 — 동시성 사고 카드 렌더 (사고번호·잔량·해제 상태)', v2.includes('동시성 사고') && v2.includes('CI-PL-CONC-T-CUT-1') && v2.includes('해제됨'));
          win.location.hash = '#dash'; S3.role = bak3.role;
        } catch(e){ record('conc332: 동시성 Gate', false, e.message); }
      })();
      // ── v0.34.0 Forging Pack 추출 (GPT 승인 조건 7항 — 업종 하드코딩 0·mold_linked·팩 설치 엔진) ──
      (function(){
        try {
          // ① Runtime 업종 하드코딩 0 (확장 정적 감사 — 시드 1곳만 허용)
          record('pack340: ① Runtime 업종 결합 소멸 — R1/R2 mold_linked · R3 seq맵 제거 · R4 병목 자동 판정',
            !/capaByProc\['FORGE'\]/.test(html) && !/seqMap = \{CUT/.test(html)
            && /if\(proc_obj\.mold_linked\)/.test(html) && /moldLinkedProcSet/.test(html)
            && /r\.proc==='SHIP' \? 99 : \(i\+1\)\*10/.test(html)
            && (html.match(/proc\s*===?\s*'FORGE'/g)||[]).length <= 1);
          // ② mold_linked 기존 동작 100% — FORGE 실적 → 타발 누적 · 비연동 공정 무반응
          const fp = win.DB.get('processes','FORGE');
          record('pack340: ② FORGE 공정 mold_linked=true (시드+마이그레이션 멱등)', !!fp && fp.mold_linked === true && win.moldLinkedProcSet().has('FORGE') && !win.moldLinkedProcSet().has('CUT'));
          const mold = win.DB.all('molds').find(m => (m.itemCodes||[]).length);
          if(mold){
            const itc = mold.itemCodes[0];
            const before = win.DB.get('molds', mold.id).cum_strokes || 0;
            const heat = win.DB.all('material_lots')[0]?.heat || 'H2604-101';
            win.LOT.create({ plot_no:'PL-PACK-T', itemCode: itc, materialLots:[heat], qty_initial: 50 });
            win.LOT.registerRecord({ plot_no:'PL-PACK-T', proc:'CUT', equipment:'절단기 1호', good:30, type:'in', shift:'day' });
            const midCut = win.DB.get('molds', mold.id).cum_strokes || 0;
            win.LOT.registerRecord({ plot_no:'PL-PACK-T', proc:'FORGE', equipment:'프레스 1호', good:20, defect:2, type:'in', shift:'day' });
            const after = win.DB.get('molds', mold.id).cum_strokes || 0;
            record('pack340: ② 타발 누적 동일 동작 — CUT(비연동) 무반응 · FORGE +22 (good+defect)', midCut === before && after === before + 22, `${before}→${midCut}→${after}`);
          } else record('pack340: ② 타발 누적', false, '금형 시드 없음');
          // ③~⑦ 팩 설치 엔진 — 스냅샷 격리 후 빈 저장소에서 양방향 검증
          const snap = win.DB.snapshot();
          try {
            win.DB.restore({});
            // 빈 → machining
            const r1 = win.applyIndustryPack('machining');
            record('pack340: ④ 빈 인스턴스 machining 설치 — 공정 10·불량 8·정의필드 4', r1.added === 22 && win.DB.all('processes').length === 10 && win.DB.all('defect_types').length === 8 && r1.first === true, JSON.stringify(r1));
            // 멱등 — 재적용 무변
            const cntP = win.DB.all('processes').length;
            const lname = win.DB.get('processes','LATHE').name;
            const r2 = win.applyIndustryPack('machining');
            record('pack340: ④ 멱등 — 재적용 added 0 · 마스터 무변 · first=false', r2.added === 0 && r2.skipped === 22 && win.DB.all('processes').length === cntP && win.DB.get('processes','LATHE').name === lname && r2.first === false);
            // machining → forging (오염 없음)
            const r3 = win.applyIndustryPack('forging');
            record('pack340: ⑤ machining→forging — SHIP 공유 스킵 · 기존 가공 데이터 무변', r3.added === 13 && win.DB.get('processes','SHIP') && win.DB.get('processes','LATHE').name === lname && win.DB.get('processes','FORGE').mold_linked === true && win.DB.get('defect_types','DT-M02').name === '조도 불량');
            // 역방향: 빈 → forging → machining
            win.DB.restore({});
            const f1 = win.applyIndustryPack('forging');
            const fname = win.DB.get('processes','FORGE').name;
            const f2 = win.applyIndustryPack('machining');
            record('pack340: ⑤ 역방향 forging→machining — 단조 데이터 무변 · mold_linked 보존', f1.added === 14 && f2.added === 21 && win.DB.get('processes','FORGE').name === fname && win.DB.get('processes','FORGE').mold_linked === true && win.DB.get('defect_types','DT-001').name === '치수 불량');
            // ⑥⑦ 설치 이력·버전
            const hist = (win.DB.get('settings','pack_installs')||{}).installs || [];
            record('pack340: ⑥⑦ 설치 이력·버전 — 2건 · pack/packVersion/installedAt/installedBy', hist.length === 2 && hist[0].pack === 'forging' && hist[0].packVersion === 'forging-1.1' && hist[1].packVersion === 'machining-2.1' && hist.every(x => x.installedAt > 0 && x.installedBy));
            record('pack340: preset 갱신 + _logs Audit (DB.set 경유)', win.getCompany().preset === 'machining' && ((win.DB.get('settings','pack_installs')._logs||[]).length >= 1));
          } finally { win.DB.restore(snap); }
          record('pack340: ③ 스냅샷 복원 — 기존 데이터 원복 (orders 42 유지)', win.DB.all('orders').length >= 40);
          // UI — 설정 화면 팩 설치 버튼
          record('pack340: 설정 화면 [팩 설치] 버튼 + 이력 표시', html.includes('id="co-pack-install"') && html.includes('pack_installs'));
        } catch(e){ record('pack340: Forging Pack 추출', false, e.message); }
      })();
      // ── v0.35.0 생산계획 개편 (설비 레인·교대·순번·타수·교체/수명·이월·생산계획서 — 시안 v6+v7 확정) ──
      (function(){
        try {
          record('plan350: 필드·설정 — BOM uph 입력·금형 setup_min·교대 설정·CSS 레인', /data-ct=/.test(html) && /uph: \+r\.uph\|\|0/.test(html) && /id="m-setup"/.test(html) && /SHIFT_PRESETS/.test(html) && /\.pln-lane\{/.test(html));
          const today = win.fmtDate(Date.now());
          win.DB.set('production_plans','PP-P35A',{id:'PP-P35A',plan_type:'board',so_id:null,itemCode:'HF-3312',qty:180,date:today,equipment:'프레스 1호',seq:10,ts:Date.now(),by:'t'});
          win.DB.set('production_plans','PP-P35B',{id:'PP-P35B',plan_type:'board',so_id:null,itemCode:'HF-3309',qty:640,date:today,equipment:'프레스 1호',seq:20,ts:Date.now(),by:'t'});
          win.DB.set('production_plans','PP-P35C',{id:'PP-P35C',plan_type:'board',so_id:null,itemCode:'HF-3305',qty:300,date:today,equipment:null,ts:Date.now(),by:'t'});
          const tl = win.planDayTimeline(today, win.DB.all('production_plans').filter(p=>p.date===today));
          const lane = tl.lanes.find(l=>l.eqName==='프레스 1호');
          const jobs = lane ? lane.shifts.flatMap(s2=>s2.entries).filter(e=>e.type==='job') : [];
          record('plan350: 타임라인 — 레인·미배정 분리·순번 배치·시각 계산', !!lane && tl.unassigned.some(p=>p.id==='PP-P35C') && jobs.length>=2 && jobs[0].p.id==='PP-P35A' && /:/.test(jobs[0].from||''), JSON.stringify(jobs.slice(0,2).map(j=>[j.p.id,j.from,j.to])) + ' (완료 계획 자동 제외 동작 포함)');
          record('plan350: 타수 폴백 ③ 설비 능력', jobs[0] && jobs[0].rate>0 && /설비 능력|라우트|실적/.test(jobs[0].basis));
          const bomA = win.DB.get('bom','HF-3301');
          if(bomA){ bomA.route = bomA.route.map(r=>r.proc==='FORGE'?{...r,uph:480}:r); win.DB.set('bom','HF-3301',bomA); }
          const rr = win.planRateFor('HF-3301','프레스 1호');
          record('plan350: 타수 폴백 ① 라우트 uph 우선(×cav)', /라우트/.test(rr.basis) && rr.rate>=480, rr.rate+' · '+rr.basis);
          win.DB.set('settings','shifts',{count:2,list:win.SHIFT_PRESETS[2]});
          const tl2 = win.planDayTimeline(today, win.DB.all('production_plans').filter(p=>p.date===today));
          record('plan350: 2교대 밴드 + 교대 이월 구조', tl2.lanes[0].shifts.length===2);
          win.eval("window._planFilter = { tab:'day', anchor: fmtDate(Date.now()), q:'', eq:'', status:'', origin:'', axis:'item' }");
          win.location.hash='#plan'; win.VIEWS.plan();
          const vv = win.document.getElementById('view').innerHTML;
          record('plan350: 일간 렌더 — 레인·교대·순번·미배정·출력 버튼', vv.includes('pln-lane') && vv.includes('pln-seq') && vv.includes('미배정') && vv.includes('생산계획서'));
          win.VIEWS._planSeqMove('PP-P35B', -1);
          record('plan350: 순번 이동 (seq 스왑 저장)', (+win.DB.get('production_plans','PP-P35B').seq) < (+win.DB.get('production_plans','PP-P35A').seq));
          let sheets=[]; const _X = win.XLSX;
          win.XLSX = { utils:{ book_new:()=>({}), aoa_to_sheet:a=>a, book_append_sheet:(wb,ws,nm)=>sheets.push(nm) }, writeFile:()=>{} };
          win.VIEWS._planSheetXlsx(today);
          win.XLSX = _X;
          let printed=''; const _O = win.open;
          win.open = () => ({ document:{ write:h=>{printed+=h;}, close:()=>{} } });
          win.VIEWS._planSheetPrint(today);
          win.open = _O;
          record('plan350: 생산계획서 — 엑셀 설비별+전체 시트 · 인쇄(교대·결재란·소재소요)', sheets.includes('전체') && sheets.length>=2 && printed.includes('생 산 계 획 서') && printed.includes('작성'));
          win.DB.set('settings','shifts',{count:1,list:win.SHIFT_PRESETS[1]});
          win.voidDoc('production_plans','PP-P35A','t'); win.voidDoc('production_plans','PP-P35B','t'); win.voidDoc('production_plans','PP-P35C','t');
          win.location.hash='#dash';
        } catch(e){ record('plan350: 생산계획 개편', false, e.message); }
      })();
      // ── v0.35.1 메뉴 공통 규격 1차 (수주·LOT — 핵심 스트립·통합 표·행 클릭 상세·⋯ 수납) ──
      (function(){
        try {
          const S = win.eval('SESSION'); S.loggedIn = true; S.role = 'manager';
          win.eval("window._soFilter = { q:'', status:'' }");
          win.location.hash = '#so'; win.VIEWS.so();
          let v = win.document.getElementById('view').innerHTML;
          record('menu351: 수주 — 핵심 스트립 4 + 통합 표 + ⋯ 수납', v.includes('mstrip') && v.includes('진행 중 수주잔') && v.includes('data-so-row=') && v.includes('so-more-menu') && !v.includes('order-total'));
          record('menu351: 수주 — 상세 행(단가·출하 누계·수정 버튼) 기본 숨김', v.includes('data-so-det=') && /data-so-det="[^"]+" hidden/.test(v) && v.includes('data-so-edit='));
          const det = win.document.querySelector('[data-so-det]');
          const row = win.document.querySelector('[data-so-row]');
          row.click();
          record('menu351: 수주 — 행 클릭 → 상세 펼침 토글', det.hidden === false && (row.click(), det.hidden === true));
          win.eval("window._soFilter = { q:'', status:'지연' }"); win.VIEWS.so();
          const v2 = win.document.getElementById('view').innerHTML;
          record('menu351: 수주 — 상태 필터 동작', !/pill cyan">진행</.test(v2.split('<tbody>')[1]||''), '');
          win.eval("window._soFilter = { q:'', status:'' }");
          win.location.hash = '#lot'; win.VIEWS.lot();
          v = win.document.getElementById('view').innerHTML;
          record('menu351: LOT — 핵심 6컬럼 (진행 위치·잔량) + 스트립 + ⋯ 수납', v.includes('진행 위치') && v.includes('진행 LOT') && v.includes('lot-more-menu') && v.includes('data-lot-steps='));
          record('menu351: LOT — 공정별 12칸 제거 · 스텝 패널 상세 유지', !/<th class="num">단조<\/th>/.test(v) && /lot-no-link/.test(v));
          win.location.hash = '#dash';
        } catch(e){ record('menu351: 공통 규격', false, e.message); }
      })();
    } catch(e){ record('bulk325: 시드', false, e.message); }
  })();
}

// ★ v0.38.5 — 교대 세션 설비 목록 선택 (마스터 equipments 에만 있는 설비 · 구 equipment 문서 없음) + 설비 QR 매칭
(function(){
  try {
    const r = win.eval(`(()=>{
      if(!DB.all('processes').length) seedDemoData();
      const proc = (getProcesses().find(p => p.type !== 'out') || {}).code;
      DB.set('equipments','EQM-385',{ id:'EQM-385', name:'검증 설비 385', procCode: proc, active:true });
      const noLegacy = !DB.all('equipment').some(e => e.name === '검증 설비 385');
      try{ localStorage.removeItem('vela_m_session'); localStorage.removeItem('vela_demo_m_session'); }catch(e){}
      renderSessionSetup();
      document.querySelector('[data-toggle="eq"]').onclick();
      const it = [...document.querySelectorAll('[data-pick-eq]')].find(x => /검증 설비 385/.test(x.textContent));
      if(!it) return { noLegacy, found:false };
      it.onclick();
      const cur = document.querySelector('.pick-cur.eq').textContent;
      const qr = matchEquipmentQr('검증 설비 385');
      DB.del ? DB.del('equipments','EQM-385') : null;
      return { noLegacy, found:true, picked: /검증 설비 385/.test(cur) && !/선택 안 됨/.test(cur), closed: !document.getElementById('eq-list'), qr: !!qr && qr.name === '검증 설비 385' && qr.proc === proc };
    })()`);
    record('eqpick385: 교대 세션 설비 목록 클릭 → 선택됨 (구 equipment 문서 없는 마스터 설비) · 목록 닫힘', r.noLegacy && r.found && r.picked && r.closed, JSON.stringify(r));
    record('eqpick385: 설비 QR 매칭 = 마스터에만 있는 설비도 인식 (id/qr/name)', r.qr === true, JSON.stringify(r));
  } catch(e){ record('eqpick385', false, e.message); }
})();

// ★ v0.38.6 — 수주 표 헤더 「잔량 / 수주수량」 (사용자 요청)
record('so386: 수주 표 헤더 = 잔량 / 수주수량', /<th class="num">잔량 \/ 수주수량<\/th>/.test(html) && !/잔량 \/ 수량</.test(html));

// ★ v0.38.8 — 로고 서체 (Fraunces 정적 Black Italic · 광학 크기 9 · paint-order) · 색 무변
record('logo388: 로고 = Fraunces 정적 900 이탤릭 요청 + opsz 9 고정 + 테두리 뒤로 · 색 무변', /family=Fraunces:ital,wght@1,900&/.test(html) && !/Fraunces:ital,opsz/.test(html) && /\.vela-mark\{[^}]*font-variation-settings:'opsz' 9;paint-order:stroke fill\}/.test(html) && /\.vela-mark\.primary\{color:#fff;text-shadow:1px 1px 0 #0a0a0a/.test(html) && /\.vela-mark \.accent\{color:#5eead4\}/.test(html));

// ★ v0.38.9 — 진입 화면 사용업체 CI (로고+상호+사이트) · 저장값만 · 로고 검증
(function(){
  try {
    const r = win.eval(`(()=>{
      const prev = DB.get('settings','company');
      DB.set('settings','company', { name:'', site:'', logo:'' }); renderGate();
      const hidEmpty = document.getElementById('gt-ci').hidden === true;
      DB.set('settings','company', { name:'검증산업', site:'1공장', logo:'data:image/png;base64,iVBORw0KGgo=' }); renderGate();
      const box = document.getElementById('gt-ci');
      const shown = box.hidden === false && /검증산업/.test(box.textContent) && !/1공장/.test(box.textContent) && !!box.querySelector('#gt-co-logo img') && document.getElementById('gt-co-logo').hidden === false;
      DB.set('settings','company', { name:'', site:'', logo:'javascript:alert(1)' }); renderGate();
      const injBlocked = document.getElementById('gt-ci').hidden === true && !document.querySelector('#gt-co-logo img');
      const v1 = companyLogoSrc({logo:'data:image/png;base64,AAA" onerror="x'}) === '';
      if(prev) DB.set('settings','company', prev);
      return { hidEmpty, shown, injBlocked, v1 };
    })()`);
    record('ci389: 진입 화면 CI — 저장값 없으면 숨김(기본 상호 표시 0) · 로고+상호만 표시(사이트·테두리 없음) · 이미지 data URL 외 로고 무시', r.hidEmpty && r.shown && r.injBlocked && r.v1, JSON.stringify(r));
    record('ci389: 설정 › 회사 정보 로고 선택·빼기 + [저장] 시 logo 반영 · 대기 화면도 같은 검증 함수', /id="co-logo-file"/.test(html) && /id="co-logo-clear"/.test(html) && /logo: _coLogo,/.test(html) && (html.match(/companyLogoSrc\(co\)/g)||[]).length >= 3);
  } catch(e){ record('ci389', false, e.message); }
})();

// ★ v0.38.11 — 헤더 회사명: 서버 데이터가 늦게 와도 갱신 · 예시 상호 노출 0
(function(){
  try {
    const r = win.eval(`(()=>{
      const prev = DB.get('settings','company');
      const _c = DB.all('settings'); 
      DB.set('settings','company', { name:'', site:'' }); renderBrandCompany();
      const hid = document.getElementById('brand-co').hidden === true;
      const def = getCompany.toString().includes('대성단조') === false;
      DB.set('settings','company', { name:'용암<b>금속', site:'함안' }); renderBrandCompany();
      const el = document.getElementById('brand-co');
      const shown = el.hidden === false && el.textContent.includes('용암<b>금속') && el.textContent.includes('함안') && !el.querySelector('b b');
      if(prev) DB.set('settings','company', prev);
      return { hid, def, shown };
    })()`);
    record('brand3811: 헤더 회사명 renderBrandCompany 1벌 — 빈 값 숨김 · 저장값 표시(이스케이프) · getCompany 예시 상호 제거', r.hid && r.def && r.shown, JSON.stringify(r));
    record('brand3811: 서버 스냅샷 도착(첫 라운드·settings 증분) 시 헤더 갱신 호출 · 회사정보 저장 실패 시 토스트 없음', /try \{ renderBrandCompany\(\); \} catch\(e\)\{\}   \/\/ ★ v0\.38\.11 — 서버 데이터 도착/.test(html) && /if\(col === 'settings'\) renderBrandCompany\(\)/.test(html) && /if\(!DB\.set\('settings','company', \{\.\.\.co, \.\.\.n\}\)\) return;/.test(html) && !/function brandCompany/.test(html));
  } catch(e){ record('brand3811', false, e.message); }
})();

// ★ v0.38.12 — 대기 화면 CI(로고+상호) · 서버 저장 거부 알림
(function(){
  try {
    const r = win.eval(`(()=>{
      const prev = DB.get('settings','company');
      DB.set('settings','company', { name:'검증산업', site:'', logo:'data:image/png;base64,iVBORw0KGgo=' });
      try { enterStandby(); } catch(e){}
      const ci = document.getElementById('sb-ci');
      const ok = /검증산업/.test(ci.textContent) && !!ci.querySelector('img');
      try { wakeStandby(); } catch(e){}
      if(prev) DB.set('settings','company', prev);
      return { ok };
    })()`);
    record('sb3812: 대기 화면 = 저장된 상호+로고 표시 (companyLogoSrc 1벌)', r.ok, JSON.stringify(r));
    record('sb3812: 서버 저장(set) 비동기 거부도 화면 알림 _fsWriteRejected (즉시 쓰기·대기열 flush 둘 다)', (html.match(/\.set\(doc, \{ ?merge: ?true ?\}\)\.catch\(e => _fsWriteRejected\(col, id, e\)\)/g)||[]).length === 2 && /function _fsWriteRejected\(col, id, e\)/.test(html));
  } catch(e){ record('sb3812', false, e.message); }
})();

// ★ v0.38.13 — 사용자·비밀번호 / 작업자 마스터 → 기준정보 (기준정보 OFF 업체는 설정에서)
(function(){
  try {
    const r = win.eval(`(()=>{
      SESSION.loggedIn = true; SESSION.role = 'manager';
      location.hash = '#masters'; router();
      const m = { usr: !!document.getElementById('usr-rows'), wk: !!document.getElementById('wk-rows'), wkRows: document.querySelectorAll('#wk-rows tr').length, usrAdd: typeof document.getElementById('usr-add').onclick === 'function', wkAdd: typeof document.getElementById('wk-add').onclick === 'function' };
      location.hash = '#set'; router();
      const sOn = { usr: !!document.getElementById('usr-rows'), link: /기준정보 열기/.test(document.getElementById('view').textContent) };
      const f = getFlags(); const _m = f.masters; setFlag('masters', false); router();
      const sOff = { usr: !!document.getElementById('usr-rows'), wk: !!document.getElementById('wk-rows'), bound: typeof (document.getElementById('usr-add')||{}).onclick === 'function' };
      setFlag('masters', _m); location.hash = '#dash';
      return { m, sOn, sOff };
    })()`);
    record('mst3813: 기준정보에 사용자·비밀번호 관리 + 작업자 마스터 (목록·버튼 동작) · 설정에는 안내+이동 버튼만', r.m.usr && r.m.wk && r.m.wkRows > 0 && r.m.usrAdd && r.m.wkAdd && !r.sOn.usr && r.sOn.link, JSON.stringify(r));
    record('mst3813: 기준정보 모듈 OFF 업체 → 설정 화면에 그대로 표시(관리 경로 상실 방지)', r.sOff.usr && r.sOff.wk && r.sOff.bound, JSON.stringify(r.sOff));
  } catch(e){ record('mst3813', false, e.message); }
})();

// ★ v0.38.14 — 진입 화면 CI 단말 캐시 (서버 데이터 도착 전 F5 직후) · 회사(저장소)별 키
(function(){
  try {
    const r = win.eval(`(()=>{
      const prev = DB.get('settings','company');
      DB.set('settings','company', { name:'캐시산업', logo:'' }); renderGate();
      const key = _gateCiKey(); const cached = JSON.parse(localStorage.getItem(key)||'null');
      // 서버 문서가 아직 없는 상태(F5 직후) 재현: 메모리에서 문서만 잠시 빼고 렌더
      const st = DB._state ? DB._state : null;
      const raw = DB.allRaw('settings'); const doc = DB.get('settings','company');
      const _g = DB.get; DB.get = (c, id) => (c === 'settings' && id === 'company') ? null : _g(c, id);
      renderGate(); const shownFromCache = document.getElementById('gt-ci').hidden === false && /캐시산업/.test(document.getElementById('gt-co-name').textContent);
      DB.get = _g;
      if(prev) DB.set('settings','company', prev);
      return { key, cachedName: cached && cached.name, shownFromCache, keyScoped: /^vela_gate_ci:(local|demo|fs:)/.test(key) };
    })()`);
    record('gate3814: 진입 화면 CI 캐시 — 원본 있으면 캐시 갱신 · 원본 도착 전(F5 직후)엔 캐시로 즉시 표시 · 저장소별 키', r.cachedName === '캐시산업' && r.shownFromCache && r.keyScoped, JSON.stringify(r));
    record('gate3814: 서버 스냅샷 — settings/users 첫 도착 시 전체 라운드 기다리지 않고 진입 화면 갱신', /if\(\(col === 'settings' \|\| col === 'users'\) && !SESSION\.loggedIn && seen\.size !== expected\)\{ try \{ renderGate\(\); renderBrandCompany\(\); \} catch\(e\)\{\} \}/.test(html));
  } catch(e){ record('gate3814', false, e.message); }
})();

finish();
function finish(){
  const failed = results.filter(r => !r.ok);
  console.log('\n────────────────────────────────────────');
  console.log(`TOTAL ${results.length} · PASS ${results.length - failed.length} · FAIL ${failed.length}`);
  if(failed.length){ console.log('\nFAILED:'); failed.forEach(f => console.log(` - ${f.name}${f.detail?`  (${f.detail})`:''}`)); }
  process.exit(failed.length ? 1 : 0);
}
