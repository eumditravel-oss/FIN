/* =========================
   FIN 산출자료(Web) app.js - FINAL (Sections + Vars)
   ✅ 섹션(구분) = Ctrl+F3 (철골/부자재/동바리 탭에서)
   ✅ 행 추가 = Shift+Ctrl+F3 (산출표에서)
   ✅ Ctrl+. 코드피커, Ctrl+Delete 행삭제, F2 편집모드
   ✅ <...> 주석 제외 계산 유지
   ✅ 변수(A/AB/A1/AB1… 최대 3자) 정의 → 산출식에서 바로 사용
   ✅ 섹션 이동 시 해당 섹션의 변수표/산출표가 분리되어 저장/로드
   ========================= */

const STORAGE_KEY = "FIN_WEB_V10";

/* ===== Seed ===== */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
];

const SECTION_DEFAULT = { name: "구분 1", count: "" };

function makeEmptyCalcRow(){
  return {
    code: "", name: "", spec: "", unit: "",
    formulaExpr: "", value: 0,
    note: "",
    surchargeMul: "", convUnit: "", convFactor: "",
    convQty: 0, finalQty: 0
  };
}
function makeEmptyCodeRow(){
  return {code:"", name:"", spec:"", unit:"", surcharge:"", conv_unit:"", conv_factor:"", note:""};
}
function makeEmptyVarRow(){
  return { key:"", expr:"", value:0, note:"" };
}

function makeTabState(){
  return {
    sections: [ { ...SECTION_DEFAULT } ],
    activeSection: 0,
    calcBySection: [ Array.from({length: 20}, makeEmptyCalcRow) ],
    varsBySection: [ Array.from({length: 12}, makeEmptyVarRow) ],
  };
}

function makeState(){
  return {
    codes: SEED_CODES,
    tabs: {
      steel: makeTabState(),
      steelSub: makeTabState(),
      support: makeTabState(),
    }
  };
}

function loadState(){
  const raw = localStorage.getItem(STORAGE_KEY);
  if(!raw) return makeState();

  try{
    const s = JSON.parse(raw);

    // codes
    s.codes = Array.isArray(s.codes) ? s.codes : SEED_CODES;

    // tabs
    if(!s.tabs) s.tabs = {};
    for(const k of ["steel","steelSub","support"]){
      if(!s.tabs[k]) s.tabs[k] = makeTabState();

      const t = s.tabs[k];

      // sections
      if(!Array.isArray(t.sections) || t.sections.length === 0) t.sections = [{...SECTION_DEFAULT}];
      if(typeof t.activeSection !== "number") t.activeSection = 0;
      t.activeSection = Math.max(0, Math.min(t.activeSection, t.sections.length-1));

      // calcBySection
      if(!Array.isArray(t.calcBySection)) t.calcBySection = [];
      // 맞춰서 늘림
      while(t.calcBySection.length < t.sections.length){
        t.calcBySection.push(Array.from({length: 20}, makeEmptyCalcRow));
      }
      // 줄이면 같이 줄임(데이터 정리)
      if(t.calcBySection.length > t.sections.length){
        t.calcBySection = t.calcBySection.slice(0, t.sections.length);
      }
      // 각 섹션 rows 보정
      t.calcBySection = t.calcBySection.map(arr => Array.isArray(arr) ? arr : Array.from({length: 20}, makeEmptyCalcRow));

      // varsBySection
      if(!Array.isArray(t.varsBySection)) t.varsBySection = [];
      while(t.varsBySection.length < t.sections.length){
        t.varsBySection.push(Array.from({length: 12}, makeEmptyVarRow));
      }
      if(t.varsBySection.length > t.sections.length){
        t.varsBySection = t.varsBySection.slice(0, t.sections.length);
      }
      t.varsBySection = t.varsBySection.map(arr => Array.isArray(arr) ? arr : Array.from({length: 12}, makeEmptyVarRow));
    }

    return s;
  }catch{
    return makeState();
  }
}

let state = loadState();
function saveState(){ localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }

/* ===== Utils ===== */
function num(v){
  const x = (v ?? "").toString().trim();
  if(x === "") return 0;
  const n = Number(x.replaceAll(",",""));
  return Number.isFinite(n) ? n : 0;
}
function roundUp3(x){
  const n = Number(x);
  if(!Number.isFinite(n)) return 0;
  return Math.ceil(n * 1000) / 1000;
}
function escapeHtml(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;");
}
function escapeAttr(s){ return escapeHtml(s).replaceAll("\n","&#10;"); }

function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}
function surchargeToMul(p){
  const x = num(p);
  return x ? (1 + x/100) : "";
}

/* =========================
   Expr eval
   - <...> 제거
   - 숫자/연산자만 허용(치환 후)
   ========================= */
function stripAngleTags(exprRaw){
  return (exprRaw ?? "").toString().replace(/<[^>]*>/g, "");
}
function safeEvalNumeric(expr){
  const s = (expr ?? "").toString().trim();
  if(!s) return 0;
  if(!/^[0-9+\-*/().\s,]+$/.test(s)) return 0;
  try{
    // eslint-disable-next-line no-new-func
    const f = new Function(`return (${s.replaceAll(",","")});`);
    const out = f();
    return Number.isFinite(out) ? out : 0;
  }catch{
    return 0;
  }
}

/* ===== Variables =====
   - key: ^[A-Za-z][A-Za-z0-9]{0,2}$  (최대3자, 시작은 영문)
   - expr: 숫자/연산자 + 변수 사용 가능
   - expr에서 <...>는 제외
   - 변수간 참조 허용(순서/반복 계산)
*/
const VAR_KEY_RE = /^[A-Za-z][A-Za-z0-9]{0,2}$/;
const VAR_TOKEN_RE = /\b[A-Za-z][A-Za-z0-9]{0,2}\b/g;

function buildVarMap(tabId){
  const t = state.tabs[tabId];
  if(!t) return {};
  const si = t.activeSection || 0;
  const vars = t.varsBySection?.[si] ?? [];

  // 초기 값 0으로
  const map = {};
  for(const v of vars){
    const k = (v.key ?? "").trim();
    if(!VAR_KEY_RE.test(k)) continue;
    map[k] = 0;
  }

  // 반복 계산(서로 참조)
  for(let pass=0; pass<6; pass++){
    let changed = false;

    for(const v of vars){
      const k = (v.key ?? "").trim();
      if(!VAR_KEY_RE.test(k)) continue;

      const raw = stripAngleTags(v.expr ?? "");
      // 변수 토큰 치환
      const replaced = raw.replace(VAR_TOKEN_RE, (tok)=>{
        if(Object.prototype.hasOwnProperty.call(map, tok)) return String(map[tok] ?? 0);
        return tok; // 미정의 변수면 그대로(=> 안전검증에서 0 처리됨)
      });

      // 치환 후에도 변수 토큰이 남아있으면(미정의) 계산 불가 -> 0
      if(/[A-Za-z]/.test(replaced)){
        if(v.value !== 0){ v.value = 0; changed = true; }
        map[k] = 0;
        continue;
      }

      const val = safeEvalNumeric(replaced);
      if((v.value ?? 0) !== val){
        v.value = val;
        changed = true;
      }
      map[k] = val;
    }

    if(!changed) break;
  }

  return map;
}

function evalExprWithVars(exprRaw, varMap){
  const raw = stripAngleTags(exprRaw ?? "");
  const replaced = raw.replace(VAR_TOKEN_RE, (tok)=>{
    if(Object.prototype.hasOwnProperty.call(varMap, tok)) return String(varMap[tok] ?? 0);
    return tok;
  });

  // 변수 토큰이 남아있으면(미정의 변수) 계산 불가 -> 0
  if(/[A-Za-z]/.test(replaced)) return 0;

  return safeEvalNumeric(replaced);
}

/* ===== Recalc ===== */
function recalcRow(row, varMap){
  const m = findCode(row.code);
  if(m){
    row.name = m.name ?? "";
    row.spec = m.spec ?? "";
    row.unit = m.unit ?? "";
    const mul = surchargeToMul(m.surcharge);
    row.surchargeMul = (mul === "" ? "" : mul);
    row.convUnit = m.conv_unit ?? "";
    if((row.convFactor ?? "").toString().trim() === "") row.convFactor = (m.conv_factor ?? "");
  }else{
    // 코드 없으면 자동값 그대로 둠(수동 입력 없음 구조)
    row.name = row.name ?? "";
    row.spec = row.spec ?? "";
    row.unit = row.unit ?? "";
  }

  row.value = evalExprWithVars(row.formulaExpr, varMap);

  const E = num(row.value);
  const K = num(row.convFactor);
  const I = num(row.surchargeMul);

  row.convQty = (K === 0 ? E : E*K);
  row.finalQty = (I === 0 ? row.convQty : row.convQty * I);
}

function recalcTabAllSections(tabId){
  const t = state.tabs[tabId];
  if(!t) return;

  for(let si=0; si<t.sections.length; si++){
    // 섹션 si를 활성으로 두고 계산(변수맵이 섹션별이기 때문)
    const prev = t.activeSection;
    t.activeSection = si;

    const varMap = buildVarMap(tabId);
    const rows = t.calcBySection?.[si] ?? [];
    rows.forEach(r => recalcRow(r, varMap));

    t.activeSection = prev;
  }
}

function recalcAll(){
  recalcTabAllSections("steel");
  recalcTabAllSections("steelSub");
  recalcTabAllSections("support");
}

/* =========================
   Mouse click -> focus cell (delegation, once)
   ========================= */
let mouseFocusWired = false;
function wireMouseFocus(){
  if(mouseFocusWired) return;
  mouseFocusWired = true;

  document.addEventListener("click", (e)=>{
    const sel = window.getSelection?.();
    if(sel && !sel.isCollapsed) return;

    const t = e.target;
    if(t?.closest?.("button,a,label,select,option")) return;

    // input/textarea 클릭은 그대로
    if(t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA")) return;

    const td = t?.closest?.("td");
    if(!td) return;

    const cell = td.querySelector("input.cell:not(.readonly), textarea.cell");
    if(!cell) return;

    editMode = false;
    setEditingClass(false);

    cell.focus();
  });
}

/* =========================
   Tabs
   ========================= */
const tabsDef = [
  { id:"codes", label:'코드(Ctrl+".")' },
  { id:"steel", label:"철골(Steel)" },
  { id:"steelTotal", label:"철골_집계" },
  { id:"steelSub", label:"철골_부자재" },
  { id:"support", label:"동바리(support)" },
  { id:"supportTotal", label:"동바리_집계" }
];

let activeTabId = "steel";

/* ===== DOM ===== */
const $tabs = document.getElementById("tabs");
const $view = document.getElementById("view");

/* ===== Panel helper ===== */
function panel(title, desc){
  const wrap = document.createElement("div");
  wrap.className = "panel";
  const h = document.createElement("div");
  h.className = "panel-header";
  const left = document.createElement("div");
  left.innerHTML = `<div class="panel-title">${title}</div><div class="panel-desc">${desc || ""}</div>`;
  h.appendChild(left);
  wrap.appendChild(h);
  return {wrap, header:h};
}

function renderTabs(){
  $tabs.innerHTML = "";
  for(const t of tabsDef){
    const b = document.createElement("button");
    b.className = "tab" + (t.id===activeTabId ? " active" : "");
    b.textContent = t.label;
    b.onclick = () => go(t.id);
    $tabs.appendChild(b);
  }
}

/* =========================
   Grid registry + wiring
   ========================= */
const cellRegistry = [];
const lastFocusCell = {
  codes: { row: 0, col: 0 },
  steel: { row: 0, col: 0 },
  steelSub: { row: 0, col: 0 },
  support: { row: 0, col: 0 },
};

function gridAttrString(gridAttrs){
  if(!gridAttrs) return "";
  const {tabId, rowIdx, colIdx} = gridAttrs;
  return `data-grid="1" data-tab="${tabId}" data-row="${rowIdx}" data-col="${colIdx}"`;
}

function inputCell(value, onChange, placeholder="", gridAttrs=null, extraAttrs=""){
  const id = crypto.randomUUID();
  cellRegistry.push({id, onChange});
  const g = gridAttrString(gridAttrs);
  return `<input class="cell" data-cell="${id}" ${g} value="${escapeAttr(value ?? "")}" placeholder="${escapeAttr(placeholder)}" ${extraAttrs}/>`;
}
function textAreaCell(value, onChange, gridAttrs=null){
  const id = crypto.randomUUID();
  cellRegistry.push({id, onChange});
  const g = gridAttrString(gridAttrs);
  return `<textarea class="cell" data-cell="${id}" ${g}>${escapeHtml(value ?? "")}</textarea>`;
}
function readonlyCell(value){
  return `<input class="cell readonly" value="${escapeAttr(value ?? "")}" readonly />`;
}

function wireCells(){
  document.querySelectorAll("[data-cell]").forEach(el=>{
    const id = el.getAttribute("data-cell");
    const meta = cellRegistry.find(x=>x.id===id);
    if(!meta) return;

    const handler = ()=>{
      meta.onChange(el.value);

      // 계산/저장
      recalcAll();
      saveState();

      // 같은 행 readonly 즉시 업데이트(산출표만)
      refreshReadonlyInSameRow(el);
    };

    el.addEventListener("input", handler);
    el.addEventListener("blur", handler);
    el.addEventListener("change", handler);

    // 산출식(col=1) Enter = 계산 + 아래행 이동
    el.addEventListener("keydown", (e)=>{
      if(e.key !== "Enter") return;
      if(el.tagName.toLowerCase() === "textarea") return;

      const col = Number(el.getAttribute("data-col") || -1);
      const tab = el.getAttribute("data-tab");
      // 산출표의 산출식 칸만 처리
      if(col !== 1) return;
      if(!["steel","steelSub","support"].includes(tab)) return;

      e.preventDefault();
      handler();
      moveGridFrom(el, +1, 0);
    });
  });

  cellRegistry.length = 0;
}

function wireFocusTracking(){
  document.querySelectorAll('[data-grid="1"]').forEach(el=>{
    el.addEventListener("focus", ()=>{
      const tab = el.getAttribute("data-tab");
      const row = Number(el.getAttribute("data-row") || 0);
      const col = Number(el.getAttribute("data-col") || 0);
      if(tab && lastFocusCell[tab]){
        lastFocusCell[tab] = { row, col };
      }
    });
  });
}

function focusGrid(tab, row, col){
  const selector = `[data-grid="1"][data-tab="${tab}"][data-row="${row}"][data-col="${col}"]`;
  const el = document.querySelector(selector);
  if(el){
    el.focus();
    if(el.scrollIntoView) el.scrollIntoView({block:"nearest", inline:"nearest"});
    return true;
  }
  return false;
}

function moveGridFrom(el, dRow, dCol){
  const tab = el.getAttribute("data-tab");
  const row = Number(el.getAttribute("data-row") || 0);
  const col = Number(el.getAttribute("data-col") || 0);

  const targetRow = row + dRow;
  const targetCol = col + dCol;

  if(focusGrid(tab, targetRow, targetCol)) return true;

  for(let offset=1; offset<=6; offset++){
    if(focusGrid(tab, targetRow, targetCol - offset)) return true;
    if(focusGrid(tab, targetRow, targetCol + offset)) return true;
  }
  return false;
}

/* ===== Readonly refresh (산출표에서만) ===== */
function getActiveCalcRows(tabId){
  const t = state.tabs[tabId];
  if(!t) return null;
  const si = t.activeSection || 0;
  return t.calcBySection?.[si] ?? null;
}
function refreshReadonlyInSameRow(activeEl){
  const tabId = activeEl?.getAttribute?.("data-tab");
  if(!tabId) return;
  if(!["steel","steelSub","support"].includes(tabId)) return;

  const rowIdx = Number(activeEl.getAttribute("data-row") || -1);
  if(rowIdx < 0) return;

  const rows = getActiveCalcRows(tabId);
  if(!rows || !rows[rowIdx]) return;
  const r = rows[rowIdx];

  const tr = activeEl.closest("tr");
  if(!tr) return;

  // readonly 순서: [name, spec, unit, value, surchargeMul, convUnit, convFactor, convQty, finalQty]
  const ro = tr.querySelectorAll("input.cell.readonly");
  if(ro.length < 9) return;

  ro[0].value = r.name ?? "";
  ro[1].value = r.spec ?? "";
  ro[2].value = r.unit ?? "";
  ro[3].value = String(roundUp3(r.value));
  ro[4].value = (r.surchargeMul === "" ? "" : String(r.surchargeMul));
  ro[5].value = r.convUnit ?? "";
  ro[6].value = (r.convFactor ?? "").toString();
  ro[7].value = String(roundUp3(r.convQty));
  ro[8].value = String(roundUp3(r.finalQty));
}

/* =========================
   Sections (구분)
   ========================= */
function ensureTabStructures(tabId){
  if(!state.tabs) state.tabs = {};
  if(!state.tabs[tabId]) state.tabs[tabId] = makeTabState();

  const t = state.tabs[tabId];

  if(!Array.isArray(t.sections) || t.sections.length === 0) t.sections = [{...SECTION_DEFAULT}];
  if(typeof t.activeSection !== "number") t.activeSection = 0;
  t.activeSection = Math.max(0, Math.min(t.activeSection, t.sections.length-1));

  if(!Array.isArray(t.calcBySection)) t.calcBySection = [];
  while(t.calcBySection.length < t.sections.length){
    t.calcBySection.push(Array.from({length: 20}, makeEmptyCalcRow));
  }
  if(t.calcBySection.length > t.sections.length){
    t.calcBySection = t.calcBySection.slice(0, t.sections.length);
  }

  if(!Array.isArray(t.varsBySection)) t.varsBySection = [];
  while(t.varsBySection.length < t.sections.length){
    t.varsBySection.push(Array.from({length: 12}, makeEmptyVarRow));
  }
  if(t.varsBySection.length > t.sections.length){
    t.varsBySection = t.varsBySection.slice(0, t.sections.length);
  }
}

function addSection(tabId){
  ensureTabStructures(tabId);
  const t = state.tabs[tabId];
  const nextIdx = t.sections.length + 1;

  t.sections.push({ name:`구분 ${nextIdx}`, count:"" });
  t.calcBySection.push(Array.from({length: 20}, makeEmptyCalcRow));
  t.varsBySection.push(Array.from({length: 12}, makeEmptyVarRow));
  t.activeSection = t.sections.length - 1;

  saveState();
}

function deleteActiveSection(tabId){
  ensureTabStructures(tabId);
  const t = state.tabs[tabId];
  if(t.sections.length <= 1) return;

  const i = t.activeSection || 0;
  t.sections.splice(i,1);
  t.calcBySection.splice(i,1);
  t.varsBySection.splice(i,1);
  t.activeSection = Math.max(0, i-1);

  saveState();
}

function setActiveSection(tabId, idx){
  ensureTabStructures(tabId);
  const t = state.tabs[tabId];
  t.activeSection = Math.max(0, Math.min(idx, t.sections.length-1));
  saveState();
}

/* =========================
   Row insert/delete
   ========================= */
function insertCalcRowBelow(tabId){
  if(!["steel","steelSub","support"].includes(tabId)) return;

  const rows = getActiveCalcRows(tabId);
  if(!rows) return;

  const {row, col} = lastFocusCell[tabId] ?? {row:0, col:0};
  const insertAt = Math.min(row + 1, rows.length);

  rows.splice(insertAt, 0, makeEmptyCalcRow());
  saveState();
  go(tabId);
  setTimeout(()=>focusGrid(tabId, insertAt, col), 0);
}

function deleteRowAtActiveFocus(){
  if(!["codes","steel","steelSub","support"].includes(activeTabId)) return;

  const { row } = lastFocusCell[activeTabId] ?? { row: 0 };
  const r = Math.max(0, Number(row) || 0);

  const ok = confirm("선택된 행을 정말 삭제할까요?");
  if(!ok) return;

  if(activeTabId === "codes"){
    if(state.codes.length === 0) return;
    if(r >= state.codes.length) return;
    state.codes.splice(r, 1);
    saveState();
    go("codes");
    const newRow = Math.min(r, state.codes.length - 1);
    if(newRow >= 0) setTimeout(()=>focusGrid("codes", newRow, 0), 0);
    return;
  }

  const rows = getActiveCalcRows(activeTabId);
  if(!rows || rows.length === 0) return;
  if(r >= rows.length) return;

  rows.splice(r, 1);
  saveState();
  go(activeTabId);

  const newRow = Math.min(r, rows.length - 1);
  if(newRow >= 0) setTimeout(()=>focusGrid(activeTabId, newRow, 0), 0);
}

/* =========================
   Views
   ========================= */
function renderCodes(){
  $view.innerHTML = "";
  const {wrap, header} = panel('코드(Ctrl+".")', "코드 마스터(수정/추가). 엑셀 업로드(.xlsx)로 한 번에 등록 가능. (행 추가: Ctrl+F3)");

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가 (Ctrl+F3)";
  addBtn.onclick = ()=>{
    const {row, col} = lastFocusCell.codes ?? {row:0, col:0};
    const insertAt = Math.min(row + 1, state.codes.length);
    state.codes.splice(insertAt, 0, makeEmptyCodeRow());
    saveState();
    go("codes");
    setTimeout(()=>focusGrid("codes", insertAt, col), 0);
  };

  const uploadLabel = document.createElement("label");
  uploadLabel.className="smallbtn";
  uploadLabel.textContent="엑셀 업로드(.xlsx)";
  const uploadInput = document.createElement("input");
  uploadInput.type="file";
  uploadInput.accept=".xlsx,.xls";
  uploadInput.hidden = true;
  uploadLabel.appendChild(uploadInput);

  uploadInput.addEventListener("change", async (e)=>{
    const file = e.target.files?.[0];
    if(!file) return;

    if(!window.XLSX){
      alert("엑셀 업로드 라이브러리(XLSX)가 로드되지 않았습니다.\nindex.html에 SheetJS 스크립트를 추가해 주세요.");
      e.target.value = "";
      return;
    }

    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, {type:"array"});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {defval:""});

      const mapRow = (r) => ({
        code: (r["코드"] ?? r["code"] ?? "").toString().trim(),
        name: (r["품명"] ?? r["name"] ?? "").toString().trim(),
        spec: (r["규격"] ?? r["spec"] ?? "").toString().trim(),
        unit: (r["단위"] ?? r["unit"] ?? "").toString().trim(),
        surcharge: (r["할증"] ?? r["surcharge"] ?? "").toString().trim(),
        conv_unit: (r["환산단위"] ?? r["conv_unit"] ?? "").toString().trim(),
        conv_factor: (r["환산계수"] ?? r["conv_factor"] ?? "").toString().trim(),
        note: (r["비고"] ?? r["note"] ?? "").toString().trim(),
      });

      const mapped = rows.map(mapRow).filter(x=>x.code);
      if(mapped.length === 0){
        alert("엑셀에서 유효한 '코드' 행을 찾지 못했습니다.\n헤더(코드/품명/규격/단위/할증/환산단위/환산계수/비고)를 확인해 주세요.");
        e.target.value = "";
        return;
      }

      if(!confirm(`엑셀에서 ${mapped.length}개 코드를 불러옵니다.\n기존 코드 마스터를 엑셀 값으로 덮어쓸까요?`)){
        e.target.value = "";
        return;
      }

      state.codes = mapped;
      saveState();
      go("codes");
    }catch(err){
      console.error(err);
      alert("엑셀 업로드 처리 중 오류가 발생했습니다.\n콘솔 로그를 확인해 주세요.");
    }finally{
      e.target.value = "";
    }
  });

  right.appendChild(addBtn);
  right.appendChild(uploadLabel);
  header.appendChild(right);

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:160px;">코드</th>
          <th style="min-width:220px;">품명</th>
          <th style="min-width:220px;">규격</th>
          <th style="min-width:90px;">단위</th>
          <th style="min-width:110px;">할증</th>
          <th style="min-width:120px;">환산단위</th>
          <th style="min-width:140px;">환산계수</th>
          <th style="min-width:260px;">비고</th>
          <th style="min-width:120px;">작업</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;
  const tbody = tableWrap.querySelector("tbody");

  state.codes.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${inputCell(r.code, v=>{ r.code=v; }, "", {tabId:"codes", rowIdx:idx, colIdx:0})}</td>
      <td>${inputCell(r.name, v=>{ r.name=v; }, "", {tabId:"codes", rowIdx:idx, colIdx:1})}</td>
      <td>${inputCell(r.spec, v=>{ r.spec=v; }, "", {tabId:"codes", rowIdx:idx, colIdx:2})}</td>
      <td>${inputCell(r.unit, v=>{ r.unit=v; }, "", {tabId:"codes", rowIdx:idx, colIdx:3})}</td>
      <td>${inputCell(r.surcharge, v=>{ r.surcharge=v; }, "예: 7", {tabId:"codes", rowIdx:idx, colIdx:4})}</td>
      <td>${inputCell(r.conv_unit, v=>{ r.conv_unit=v; }, "", {tabId:"codes", rowIdx:idx, colIdx:5})}</td>
      <td>${inputCell(r.conv_factor, v=>{ r.conv_factor=v; }, "", {tabId:"codes", rowIdx:idx, colIdx:6})}</td>
      <td>${textAreaCell(r.note, v=>{ r.note=v; }, {tabId:"codes", rowIdx:idx, colIdx:7})}</td>
      <td></td>
    `;
    const tdAct = tr.lastElementChild;
    const act = document.createElement("div");
    act.className="row-actions";
    const del = document.createElement("button");
    del.className="smallbtn"; del.textContent="삭제";
    del.onclick=()=>{ state.codes.splice(idx,1); saveState(); go("codes"); };
    act.appendChild(del);
    tdAct.appendChild(act);
    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);

  wireCells();
  wireFocusTracking();
  wireMouseFocus();

  const {row, col} = lastFocusCell.codes;
  setTimeout(()=>focusGrid("codes", row, col), 0);
}

function renderCalcSheet(title, tabId){
  ensureTabStructures(tabId);
  $view.innerHTML = "";

  const t = state.tabs[tabId];
  const si = t.activeSection || 0;
  const sections = t.sections;
  const rows = t.calcBySection[si];
  const vars = t.varsBySection[si];

  // ===== TOP (구분 + 변수표) : top-split (CSS로 위 배치 조정)
  const top = document.createElement("div");
  top.className = "top-split";

  // LEFT: section list + editor + buttons
  const leftBox = document.createElement("div");
  leftBox.className = "rail-box";

  leftBox.innerHTML = `
    <div class="rail-title">구분명 리스트 (↑/↓ 이동)</div>
    <div class="section-list" id="sectionList"></div>

    <div class="section-editor">
      <input id="secName" placeholder="구분명(예: 2층 바닥 철골보)" />
      <input id="secCount" placeholder="개소(예: 0,1,2...)" />
      <button class="smallbtn" id="btnSecSave">저장</button>
    </div>

    <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
      <button class="smallbtn" id="btnSecAdd">구분 추가 (Ctrl+F3)</button>
      <button class="smallbtn" id="btnSecDel">구분 삭제</button>
    </div>
  `;

  // RIGHT: variables table
  const rightBox = document.createElement("div");
  rightBox.className = "rail-box";
  rightBox.innerHTML = `
    <div class="rail-title">변수표 (A, AB, A1, AB1... 최대 3자)</div>
    <div class="var-tablewrap">
      <table class="var-table">
        <thead>
          <tr>
            <th style="min-width:140px;">변수</th>
            <th style="min-width:220px;">산식</th>
            <th style="min-width:120px;">값</th>
            <th style="min-width:220px;">비고</th>
          </tr>
        </thead>
        <tbody id="varTbody"></tbody>
      </table>
    </div>
    <div class="var-hint">
      • 변수명은 영문으로 시작, 최대 3자 (예: A, AB, A1)<br/>
      • 산출식에서 변수 사용 가능 (예: (A+0.5)*2 )<br/>
      • &lt;...&gt; 안은 주석(계산 제외)
    </div>
  `;

  top.appendChild(leftBox);
  top.appendChild(rightBox);
  $view.appendChild(top);

  // ===== Calc panel
  const desc = '구분(↑/↓) 이동 → 해당 구분의 변수/산출표로 전환 | 산출식 Enter 계산 | Ctrl+. 코드선택 | Ctrl+F3 구분추가 | Shift+Ctrl+F3 행추가';
  const {wrap, header} = panel(title, desc);

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addRowBtn = document.createElement("button");
  addRowBtn.className="smallbtn";
  addRowBtn.textContent="행 추가 (Shift+Ctrl+F3)";
  addRowBtn.onclick=()=>insertCalcRowBelow(tabId);

  const add10Btn = document.createElement("button");
  add10Btn.className="smallbtn";
  add10Btn.textContent="+10행";
  add10Btn.onclick=()=>{
    for(let i=0;i<10;i++) rows.push(makeEmptyCalcRow());
    saveState();
    go(tabId);
  };

  right.appendChild(addRowBtn);
  right.appendChild(add10Btn);
  header.appendChild(right);

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:70px;">No</th>
          <th style="min-width:170px;">코드</th>
          <th style="min-width:220px;">품명(자동)</th>
          <th style="min-width:220px;">규격(자동)</th>
          <th style="min-width:90px;">단위(자동)</th>
          <th style="min-width:550px;">산출식</th>
          <th style="min-width:110px;">물량(Value)</th>
          <th style="min-width:90px;">할증(배수)</th>
          <th style="min-width:90px;">환산단위</th>
          <th style="min-width:110px;">환산계수</th>
          <th style="min-width:120px;">환산수량</th>
          <th style="min-width:130px;">할증후수량</th>
          <th style="min-width:120px;">작업</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;

  // 현재 섹션 변수맵/행 계산
  const varMap = buildVarMap(tabId);
  rows.forEach(r => recalcRow(r, varMap));

  const tbody = tableWrap.querySelector("tbody");
  rows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${inputCell(r.code, v=>{ r.code=v; }, "코드 입력", {tabId, rowIdx:idx, colIdx:0})}</td>
      <td>${readonlyCell(r.name)}</td>
      <td>${readonlyCell(r.spec)}</td>
      <td>${readonlyCell(r.unit)}</td>

      <td>${inputCell(r.formulaExpr, v=>{ r.formulaExpr=v; }, "예: (A+0.5)*2  (<...>는 주석)", {tabId, rowIdx:idx, colIdx:1})}</td>
      <td>${readonlyCell(String(roundUp3(r.value)))}</td>

      <td>${readonlyCell(r.surchargeMul === "" ? "" : String(r.surchargeMul))}</td>
      <td>${readonlyCell(r.convUnit)}</td>
      <td>${readonlyCell((r.convFactor ?? "").toString())}</td>

      <td>${readonlyCell(String(roundUp3(r.convQty)))}</td>
      <td>${readonlyCell(String(roundUp3(r.finalQty)))}</td>
      <td></td>
    `;

    const tdAct = tr.lastElementChild;
    const act = document.createElement("div");
    act.className="row-actions";

    const dup = document.createElement("button");
    dup.className="smallbtn"; dup.textContent="복제";
    dup.onclick=()=>{ rows.splice(idx+1,0, JSON.parse(JSON.stringify(r))); saveState(); go(tabId); };

    const del = document.createElement("button");
    del.className="smallbtn"; del.textContent="삭제";
    del.onclick=()=>{ rows.splice(idx,1); saveState(); go(tabId); };

    act.appendChild(dup);
    act.appendChild(del);
    tdAct.appendChild(act);

    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);

  // 합계
  const sumBox = document.createElement("div");
  sumBox.style.marginTop="10px";
  const sumVal = (tabId === "support")
    ? rows.reduce((a,b)=> a + num(b.value), 0)
    : rows.reduce((a,b)=> a + num(b.finalQty), 0);
  sumBox.innerHTML = `<span class="badge">합계: ${roundUp3(sumVal)}</span>`;
  wrap.appendChild(sumBox);

  $view.appendChild(wrap);

  // ===== Wire section + vars UI
  const $sectionList = document.getElementById("sectionList");
  const $secName = document.getElementById("secName");
  const $secCount = document.getElementById("secCount");
  const $btnSecSave = document.getElementById("btnSecSave");
  const $btnSecAdd = document.getElementById("btnSecAdd");
  const $btnSecDel = document.getElementById("btnSecDel");
  const $varTbody = document.getElementById("varTbody");

  function renderSectionList(){
    $sectionList.innerHTML = "";
    sections.forEach((s, i)=>{
      const div = document.createElement("div");
      div.className = "section-item" + (i === t.activeSection ? " active" : "");
      div.tabIndex = 0;
      div.innerHTML = `
        <div>
          <div style="font-weight:900;">${escapeHtml(s.name || `구분 ${i+1}`)}</div>
          <div class="meta">개소: ${escapeHtml(s.count ?? "")}</div>
        </div>
        <div class="meta">선택</div>
      `;
      div.addEventListener("click", ()=>{
        setActiveSection(tabId, i);
        go(tabId);
      });
      div.addEventListener("keydown", (e)=>{
        if(e.key === "Enter"){
          setActiveSection(tabId, i);
          go(tabId);
        }
      });
      $sectionList.appendChild(div);
    });

    const cur = sections[t.activeSection] || {name:"",count:""};
    $secName.value = cur.name ?? "";
    $secCount.value = cur.count ?? "";
  }

  function renderVarTable(){
    $varTbody.innerHTML = "";
    vars.forEach((v)=>{
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><input class="cell" value="${escapeAttr(v.key ?? "")}" placeholder="A / AB / A1" /></td>
        <td><input class="cell" value="${escapeAttr(v.expr ?? "")}" placeholder="산식" /></td>
        <td><input class="cell readonly" value="${escapeAttr(String(v.value ?? 0))}" readonly /></td>
        <td><input class="cell" value="${escapeAttr(v.note ?? "")}" placeholder="비고" /></td>
      `;

      const $key  = tr.children[0].querySelector("input");
      const $expr = tr.children[1].querySelector("input");
      const $val  = tr.children[2].querySelector("input");
      const $note = tr.children[3].querySelector("input");

      const recalcVarsAndRows = ()=>{
        v.key  = ($key.value ?? "").trim();
        v.expr = ($expr.value ?? "").trim();
        v.note = ($note.value ?? "").trim();

        // 변수 전체 재계산
        const map = buildVarMap(tabId);

        // 값 UI 갱신
        // buildVarMap에서 v.value도 업데이트됨
        $val.value = String(v.value ?? 0);

        // 산출표도 현재 섹션만 즉시 재계산 + readonly 갱신 위해 go 대신 전체 재렌더
        // (행 수 많아도 안정적으로 동기화)
        saveState();
        go(tabId);
      };

      $key.addEventListener("input", recalcVarsAndRows);
      $expr.addEventListener("input", recalcVarsAndRows);
      $note.addEventListener("input", ()=>{
        v.note = ($note.value ?? "").trim();
        saveState();
      });

      $varTbody.appendChild(tr);
    });
  }

  // 섹션 리스트 키보드 ↑/↓ 이동 (입력창과 충돌 없음)
  $sectionList.addEventListener("keydown", (e)=>{
    if(e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
    e.preventDefault();
    const next = t.activeSection + (e.key === "ArrowDown" ? 1 : -1);
    const clamped = Math.max(0, Math.min(next, sections.length-1));
    if(clamped !== t.activeSection){
      setActiveSection(tabId, clamped);
      go(tabId);
    }
  });

  $btnSecSave.addEventListener("click", ()=>{
    const cur = sections[t.activeSection];
    cur.name = ($secName.value ?? "").toString();
    cur.count = ($secCount.value ?? "").toString();
    saveState();
    renderSectionList();
  });

  $btnSecAdd.addEventListener("click", ()=>{
    addSection(tabId);
    go(tabId);
  });

  $btnSecDel.addEventListener("click", ()=>{
    deleteActiveSection(tabId);
    go(tabId);
  });

  renderSectionList();
  renderVarTable();

  // ===== Wire grids
  wireCells();
  wireFocusTracking();
  wireMouseFocus();

  // 산출표 포커스 복원
  const last = lastFocusCell[tabId] ?? {row:0,col:0};
  setTimeout(()=>focusGrid(tabId, last.row, last.col), 0);
}

/* ===== Totals ===== */
function groupSum(rows, valueSelector){
  const map = new Map();
  for(const r of rows){
    const code = (r.code ?? "").toString().trim();
    if(!code) continue;
    const m = findCode(code);
    const cur = map.get(code) ?? {
      code,
      name: m?.name ?? r.name ?? "",
      spec: m?.spec ?? r.spec ?? "",
      unit: ((m?.conv_unit ?? "").toString().trim() !== "" ? m.conv_unit : (m?.unit ?? r.unit ?? "")),
      sum: 0
    };
    cur.sum += valueSelector(r);
    map.set(code, cur);
  }
  return Array.from(map.values()).sort((a,b)=>a.code.localeCompare(b.code));
}

function flattenAllRows(tabId){
  const t = state.tabs[tabId];
  if(!t) return [];
  const out = [];
  for(let si=0; si<t.sections.length; si++){
    out.push(...(t.calcBySection?.[si] ?? []));
  }
  return out;
}

function renderSteelTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(철골+부자재의 할증후수량 합산)");
  recalcAll();

  const steelRows = flattenAllRows("steel");
  const subRows   = flattenAllRows("steelSub");
  const grouped = groupSum([...steelRows, ...subRows], r => num(r.finalQty));

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table><thead><tr>
      <th style="min-width:170px;">코드</th>
      <th style="min-width:220px;">품명</th>
      <th style="min-width:220px;">규격</th>
      <th style="min-width:90px;">단위</th>
      <th style="min-width:160px;">할증후수량</th>
    </tr></thead><tbody></tbody></table>
  `;
  const tbody = tableWrap.querySelector("tbody");
  grouped.forEach(g=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(g.code)}</td>
      <td>${escapeHtml(g.name)}</td>
      <td>${escapeHtml(g.spec)}</td>
      <td>${escapeHtml(g.unit)}</td>
      <td>${roundUp3(g.sum)}</td>
    `;
    tbody.appendChild(tr);
  });
  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);
}

function renderSupportTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("동바리_집계(Support_Total quantity)", "코드별 합계(동바리의 물량(Value) 합계)");
  recalcAll();

  const supportRows = flattenAllRows("support");
  const grouped = groupSum(supportRows, r => num(r.value));

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table><thead><tr>
      <th style="min-width:170px;">코드</th>
      <th style="min-width:220px;">품명</th>
      <th style="min-width:220px;">규격</th>
      <th style="min-width:90px;">단위</th>
      <th style="min-width:160px;">물량(Value)</th>
    </tr></thead><tbody></tbody></table>
  `;
  const tbody = tableWrap.querySelector("tbody");
  grouped.forEach(g=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(g.code)}</td>
      <td>${escapeHtml(g.name)}</td>
      <td>${escapeHtml(g.spec)}</td>
      <td>${escapeHtml(g.unit)}</td>
      <td>${roundUp3(g.sum)}</td>
    `;
    tbody.appendChild(tr);
  });
  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);
}

/* =========================
   go()
   ========================= */
function go(id, opts={silentTabRender:false}){
  activeTabId = id;
  recalcAll();
  saveState();
  if(!opts.silentTabRender) renderTabs();

  if(id==="codes") renderCodes();
  else if(id==="steel") renderCalcSheet("철골(Steel)", "steel");
  else if(id==="steelSub") renderCalcSheet("철골_부자재(Processing and assembly)", "steelSub");
  else if(id==="support") renderCalcSheet("동바리(support)", "support");
  else if(id==="steelTotal") renderSteelTotal();
  else if(id==="supportTotal") renderSupportTotal();
}

renderTabs();
go(activeTabId);

/* =========================
   Code Picker window (Ctrl+.)
   ========================= */
let pickerWin = null;

function openPickerWindow(){
  let origin = activeTabId;
  if(!["steel","steelSub","support"].includes(origin)) origin = "steel";

  const focusRow = (lastFocusCell[origin]?.row ?? 0);
  const w = 1100, h = 720;
  const x = Math.max(0, (window.screenX || 0) + (window.outerWidth - w) / 2);
  const y = Math.max(0, (window.screenY || 0) + (window.outerHeight - h) / 2);

  pickerWin = window.open(
    "picker.html",
    "FIN_CODE_PICKER",
    `width=${w},height=${h},left=${x},top=${y},resizable=yes,scrollbars=yes`
  );

  if(!pickerWin){
    alert("팝업이 차단되어 새 창을 열 수 없습니다. 브라우저에서 팝업 허용 후 다시 시도해 주세요.");
    return;
  }

  const payload = {
    type: "INIT",
    originTab: origin,
    focusRow,
    codes: state.codes
  };

  const sendInit = () => { try { pickerWin.postMessage(payload, window.location.origin); } catch {} };
  setTimeout(sendInit, 80);
  setTimeout(sendInit, 250);
  setTimeout(sendInit, 600);
}

function insertCodesBelow(tab, focusRow, codeList){
  ensureTabStructures(tab);

  const rows = getActiveCalcRows(tab);
  if(!rows) return;

  const idx = Math.min(Math.max(0, Number(focusRow) || 0), rows.length);
  const insertAt = Math.min(idx + 1, rows.length);

  const newRows = codeList.map(code=>{
    const r = makeEmptyCalcRow();
    r.code = code;
    return r;
  });

  rows.splice(insertAt, 0, ...newRows);
  recalcAll();
  saveState();
  go(tab);
  setTimeout(()=>focusGrid(tab, insertAt, 0), 0);
}

window.addEventListener("message", (event)=>{
  if(event.origin !== window.location.origin) return;
  const msg = event.data;
  if(!msg || typeof msg !== "object") return;

  if(msg.type === "INSERT_SELECTED"){
    const { originTab, focusRow, selectedCodes } = msg;
    if(Array.isArray(selectedCodes) && selectedCodes.length){
      insertCodesBelow(originTab, focusRow, selectedCodes);
    }
    return;
  }

  if(msg.type === "UPDATE_CODES"){
    const next = msg.codes;
    if(!Array.isArray(next) || next.length === 0){
      alert("코드 반영 실패: 전달된 codes가 비어있습니다.");
      return;
    }

    const cleaned = next
      .map(r => ({
        code: (r.code ?? "").toString().trim(),
        name: (r.name ?? "").toString().trim(),
        spec: (r.spec ?? "").toString().trim(),
        unit: (r.unit ?? "").toString().trim(),
        surcharge: (r.surcharge ?? "").toString().trim(),
        conv_unit: (r.conv_unit ?? "").toString().trim(),
        conv_factor: (r.conv_factor ?? "").toString().trim(),
        note: (r.note ?? "").toString().trim(),
      }))
      .filter(r => r.code);

    if(cleaned.length === 0){
      alert("코드 반영 실패: 유효한 code가 없습니다.");
      return;
    }

    const seen = new Set();
    const dup = [];
    for(const r of cleaned){
      if(seen.has(r.code)) dup.push(r.code);
      else seen.add(r.code);
    }
    if(dup.length){
      alert(`코드 반영 실패: 중복 코드 존재\n${dup.slice(0,20).join(", ")}${dup.length>20 ? "..." : ""}`);
      return;
    }

    state.codes = cleaned;
    recalcAll();
    saveState();
    go(activeTabId);
    return;
  }

  if(msg.type === "CLOSE_PICKER"){
    try{ pickerWin?.close(); }catch{}
    pickerWin = null;
    return;
  }
});

/* =========================
   HOTKEYS + 방향키/편집모드 (capture)
   ========================= */
let editMode = false;

function setEditingClass(on){
  document.querySelectorAll('.cell.editing').forEach(x=>x.classList.remove('editing'));
  if(on){
    const el = document.activeElement;
    if(el && el.classList && el.classList.contains("cell")) el.classList.add("editing");
  }
}
function isGridEl(el){
  return el && el.getAttribute && el.getAttribute("data-grid") === "1";
}
function isTextareaEl(el){
  return (el?.tagName || "").toLowerCase() === "textarea";
}
function caretAtStart(el){
  try{ return (el.selectionStart ?? 0) === 0 && (el.selectionEnd ?? 0) === 0; }catch{ return false; }
}
function caretAtEnd(el){
  try{
    const len = (el.value ?? "").length;
    return (el.selectionStart ?? 0) === len && (el.selectionEnd ?? 0) === len;
  }catch{ return false; }
}
function textareaAtTop(el){
  const v = el.value ?? "";
  const pos = el.selectionStart ?? 0;
  return !v.slice(0, pos).includes("\n");
}
function textareaAtBottom(el){
  const v = el.value ?? "";
  const pos = el.selectionStart ?? 0;
  return !v.slice(pos).includes("\n");
}

document.addEventListener("keydown", (e)=>{
  // Ctrl+. picker
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "." || e.code === "Period")){
    e.preventDefault();
    openPickerWindow();
    return;
  }

  // Ctrl+Delete = 행삭제
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "Delete" || e.code === "Delete")){
    e.preventDefault();
    deleteRowAtActiveFocus();
    return;
  }

  // Ctrl+F3 = (codes: 행추가) / (calc tabs: 구분추가)
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3" && !e.shiftKey){
    e.preventDefault();
    if(activeTabId === "codes"){
      const {row, col} = lastFocusCell.codes ?? {row:0, col:0};
      const insertAt = Math.min(row + 1, state.codes.length);
      state.codes.splice(insertAt, 0, makeEmptyCodeRow());
      saveState();
      go("codes");
      setTimeout(()=>focusGrid("codes", insertAt, col), 0);
      return;
    }
    if(["steel","steelSub","support"].includes(activeTabId)){
      addSection(activeTabId);
      go(activeTabId);
      return;
    }
    return;
  }

  // Shift+Ctrl+F3 = 행추가 (산출표)
  if(e.ctrlKey && e.shiftKey && !e.altKey && !e.metaKey && e.code === "F3"){
    e.preventDefault();
    if(["steel","steelSub","support"].includes(activeTabId)){
      insertCalcRowBelow(activeTabId);
    }
    return;
  }

  const el = document.activeElement;
  if(!isGridEl(el)) return;

  // F2 편집모드
  if(e.key === "F2"){
    e.preventDefault();
    editMode = true;
    setEditingClass(true);
    if(el.setSelectionRange){
      const len = (el.value ?? "").length;
      el.setSelectionRange(len, len);
    }
    return;
  }

  if(editMode && e.key === "Escape"){
    e.preventDefault();
    editMode = false;
    setEditingClass(false);
    return;
  }

  if(editMode && e.key === "Enter" && !isTextareaEl(el)){
    e.preventDefault();
    editMode = false;
    setEditingClass(false);
    moveGridFrom(el, +1, 0);
    return;
  }

  if(editMode) return;

  // textarea: 커서가 끝/처음일 때만 이동
  if(isTextareaEl(el)){
    if(e.key === "ArrowUp"){
      if(!textareaAtTop(el)) return;
      e.preventDefault();
      moveGridFrom(el, -1, 0);
      return;
    }
    if(e.key === "ArrowDown"){
      if(!textareaAtBottom(el)) return;
      e.preventDefault();
      moveGridFrom(el, +1, 0);
      return;
    }
    if(e.key === "ArrowLeft"){
      if(!caretAtStart(el)) return;
      e.preventDefault();
      moveGridFrom(el, 0, -1);
      return;
    }
    if(e.key === "ArrowRight"){
      if(!caretAtEnd(el)) return;
      e.preventDefault();
      moveGridFrom(el, 0, +1);
      return;
    }
    return;
  }

  // Excel-like arrows
  if(e.key === "ArrowUp"){
    e.preventDefault();
    moveGridFrom(el, -1, 0);
    return;
  }
  if(e.key === "ArrowDown"){
    e.preventDefault();
    moveGridFrom(el, +1, 0);
    return;
  }
  if(e.key === "ArrowLeft"){
    e.preventDefault();
    moveGridFrom(el, 0, -1);
    return;
  }
  if(e.key === "ArrowRight"){
    e.preventDefault();
    moveGridFrom(el, 0, +1);
    return;
  }
}, true); // capture

/* ===== Buttons ===== */
document.getElementById("btnOpenPicker")?.addEventListener("click", openPickerWindow);

document.getElementById("btnExport")?.addEventListener("click", ()=>{
  recalcAll(); saveState();
  const blob = new Blob([JSON.stringify(state, null, 2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `FIN_WEB_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
});

document.getElementById("fileImport")?.addEventListener("change", async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  const txt = await f.text();
  try{
    const parsed = JSON.parse(txt);
    state = parsed;
    // 로드 보정
    state = loadStateFromObject(state);
    recalcAll();
    saveState();
    go("steel");
  }catch(err){
    console.error(err);
    alert("JSON 파싱 실패: 파일 내용을 확인해 주세요.");
  }finally{
    e.target.value = "";
  }
});

// JSON import 보정(파일 import시)
function loadStateFromObject(obj){
  try{
    const tmp = obj ?? {};
    if(!tmp.codes) tmp.codes = SEED_CODES;
    if(!Array.isArray(tmp.codes)) tmp.codes = SEED_CODES;

    if(!tmp.tabs) tmp.tabs = {};
    for(const k of ["steel","steelSub","support"]){
      if(!tmp.tabs[k]) tmp.tabs[k] = makeTabState();
      // ensure
      state = tmp; // 임시로 state에 넣고 ensure 돌려도 됨
      ensureTabStructures(k);
    }
    return tmp;
  }catch{
    return makeState();
  }
}

document.getElementById("btnReset")?.addEventListener("click", ()=>{
  if(!confirm("정말 초기화할까요? (로컬 저장 데이터가 삭제됩니다)")) return;
  localStorage.removeItem(STORAGE_KEY);
  state = makeState();
  recalcAll();
  saveState();
  go("steel");
});
