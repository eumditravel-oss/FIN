/* =========================
   FIN 산출자료(Web) app.js - "구분(세로) + 변수표 + 산출표" 구조
   - ✅ 노란(구분명 리스트) : 세로 목록 + ↑/↓로 이동하면서 산출 위치 전환
   - ✅ 빨강(변수표) : 변수명(최대3자) + 산식 + Enter -> 값 계산
   - ✅ 초록(산출표) : 기존 구현 + 산출식에서 변수 사용 가능
   - ✅ <...> 주석은 변수/산출식 모두 계산 제외
   - ✅ Ctrl+F3 : 구분(섹션) 추가 + 그 섹션으로 즉시 이동
   - ✅ Ctrl+Shift+F3 : 산출표 행 추가
   - ✅ Ctrl+. : 코드 picker (기존 유지)
   ========================= */

const STORAGE_KEY = "FIN_WEB_V10";

/* ===== Seed codes ===== */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
];

/* =========================
   Mouse click -> focus cell (delegation, once)
   ========================= */
let mouseFocusWired = false;
let editMode = false;

function setEditingClass(on){
  document.querySelectorAll('.cell.editing').forEach(x=>x.classList.remove('editing'));
  if(on){
    const el = document.activeElement;
    if(el && el.classList && el.classList.contains("cell")) el.classList.add("editing");
  }
}

function wireMouseFocus(){
  if(mouseFocusWired) return;
  mouseFocusWired = true;

  document.addEventListener("click", (e)=>{
    const sel = window.getSelection?.();
    if(sel && !sel.isCollapsed) return;

    const t = e.target;
    if(t?.closest?.("button,a,label,select,option")) return;
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

/* ===== Factories ===== */
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
function makeDefaultRows(n=20){
  return Array.from({length:n}, makeEmptyCalcRow);
}

/* ===== 변수표 row ===== */
function makeVarRow(){
  return { key:"", expr:"", value:0, note:"" };
}
function makeDefaultVars(n=30){
  return Array.from({length:n}, makeVarRow);
}

/* ===== Section model ===== */
function makeDefaultSection(){
  return { label:"", count:"", vars: makeDefaultVars(30), rows: makeDefaultRows(20) };
}
function makeSectionsPackFromLegacyRows(rows){
  return {
    activeIndex: 0,
    sections: [{ label:"", count:"", vars: makeDefaultVars(30), rows: Array.isArray(rows)? rows : makeDefaultRows(20) }]
  };
}

/* ===== State ===== */
function makeState(){
  return {
    codes: SEED_CODES,
    sheets: {
      steel:    { activeIndex: 0, sections: [ makeDefaultSection() ] },
      steelSub: { activeIndex: 0, sections: [ makeDefaultSection() ] },
      support:  { activeIndex: 0, sections: [ makeDefaultSection() ] },
    }
  };
}

function loadState(){
  const raw = localStorage.getItem(STORAGE_KEY);
  if(!raw) return makeState();

  try{
    const s = JSON.parse(raw);

    s.codes = Array.isArray(s.codes) ? s.codes : SEED_CODES;

    // migrate (이전 버전 steel/steelSub/support 있으면 섹션 1개로 감싸기)
    if(!s.sheets){
      s.sheets = {};
      s.sheets.steel    = makeSectionsPackFromLegacyRows(s.steel);
      s.sheets.steelSub = makeSectionsPackFromLegacyRows(s.steelSub);
      s.sheets.support  = makeSectionsPackFromLegacyRows(s.support);
      delete s.steel; delete s.steelSub; delete s.support;
    }

    for(const k of ["steel","steelSub","support"]){
      if(!s.sheets[k]) s.sheets[k] = { activeIndex:0, sections:[makeDefaultSection()] };
      if(!Array.isArray(s.sheets[k].sections) || s.sheets[k].sections.length===0){
        s.sheets[k].sections = [makeDefaultSection()];
      }
      if(typeof s.sheets[k].activeIndex !== "number") s.sheets[k].activeIndex = 0;
      if(s.sheets[k].activeIndex < 0) s.sheets[k].activeIndex = 0;
      if(s.sheets[k].activeIndex >= s.sheets[k].sections.length) s.sheets[k].activeIndex = 0;

      s.sheets[k].sections.forEach(sec=>{
        if(!sec || typeof sec !== "object") return;
        if(sec.label == null) sec.label = "";
        if(sec.count == null) sec.count = "";
        if(!Array.isArray(sec.rows)) sec.rows = makeDefaultRows(20);
        if(!Array.isArray(sec.vars)) sec.vars = makeDefaultVars(30);
      });
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

/* ===== Code master ===== */
function surchargeToMul(p){ const x = num(p); return x ? (1 + x/100) : ""; }

function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}

/* ===== Sheet helpers ===== */
function getPack(tab){ return state.sheets?.[tab] ?? null; }
function getActiveSection(tab){
  const pack = getPack(tab);
  if(!pack) return null;
  const i = Math.min(Math.max(0, pack.activeIndex|0), pack.sections.length-1);
  return pack.sections[i] ?? null;
}
function getActiveRows(tab){ return getActiveSection(tab)?.rows ?? null; }
function getActiveVars(tab){ return getActiveSection(tab)?.vars ?? null; }
function setActiveSectionIndex(tab, idx){
  const pack = getPack(tab);
  if(!pack) return;
  const n = pack.sections.length;
  const next = Math.min(Math.max(0, idx|0), n-1);
  pack.activeIndex = next;
  saveState();
}
function addSection(tab){
  const pack = getPack(tab);
  if(!pack) return;
  const cur = pack.activeIndex|0;
  const insertAt = Math.min(cur+1, pack.sections.length);
  pack.sections.splice(insertAt, 0, makeDefaultSection());
  pack.activeIndex = insertAt;
  saveState();
}

/* =========================================================
   ✅ 식 평가: 변수 지원 + <...> 주석 제거
   - 허용문자: 숫자, + - * / ( ) . , 공백, 변수토큰(영문/숫자)
   - 변수토큰 규칙(요구): 최대 3자, 첫 글자는 영문, 나머지는 영문/숫자
     예) A, AB, A1, AB1, ABC
========================================================= */
const VAR_TOKEN_RE = /\b[A-Za-z][A-Za-z0-9]{0,2}\b/g;
function isValidVarName(s){
  const t = (s ?? "").toString().trim();
  return /^[A-Za-z][A-Za-z0-9]{0,2}$/.test(t);
}

function stripAngleComment(exprRaw){
  return (exprRaw ?? "").toString().replace(/<[^>]*>/g, "");
}

function evalExprWithVars(exprRaw, varMap){
  const withoutTags = stripAngleComment(exprRaw);
  const s = withoutTags.trim();
  if(!s) return 0;

  // 허용 문자 검사(변수 포함)
  if(!/^[0-9A-Za-z+\-*/().\s,]+$/.test(s)) return 0;

  // 변수 치환
  const replaced = s.replace(VAR_TOKEN_RE, (tok)=>{
    const key = tok.toUpperCase();
    const v = varMap?.[key];
    const n = Number(v);
    return Number.isFinite(n) ? String(n) : "0";
  });

  // 치환 후 숫자/연산자만 남아야 함
  if(!/^[0-9+\-*/().\s,]+$/.test(replaced)) return 0;

  try{
    // eslint-disable-next-line no-new-func
    const f = new Function(`return (${replaced.replaceAll(",","")});`);
    const out = f();
    return Number.isFinite(out) ? out : 0;
  }catch{
    return 0;
  }
}

/* ===== 변수 재계산(상단부터, 다른 변수 참조 허용) ===== */
function buildVarMapFromRows(varRows){
  const map = {};
  (varRows || []).forEach(r=>{
    const k = (r.key ?? "").toString().trim().toUpperCase();
    if(!isValidVarName(k)) return;
    map[k] = num(r.value);
  });
  return map;
}

function recalcVars(tab){
  const vars = getActiveVars(tab);
  if(!vars) return {map:{}};

  // 여러 번 반복해서 의존성 해소(간단 고정점)
  let map = buildVarMapFromRows(vars);
  for(let pass=0; pass<5; pass++){
    for(const r of vars){
      const k = (r.key ?? "").toString().trim().toUpperCase();
      if(!isValidVarName(k)) continue;

      const val = evalExprWithVars(r.expr, map);
      r.value = val;
      map[k] = val;
    }
  }
  return {map};
}

/* ===== 산출행 계산(변수 반영) ===== */
function recalcRow(row, varMap){
  const m = findCode(row.code);
  if(m){
    row.name = m.name ?? "";
    row.spec = m.spec ?? "";
    row.unit = m.unit ?? "";
    const mul = surchargeToMul(m.surcharge);
    row.surchargeMul = mul === "" ? "" : mul;
    row.convUnit = m.conv_unit ?? "";
    if((row.convFactor ?? "").toString().trim() === "") row.convFactor = (m.conv_factor ?? "");
  }

  row.value = evalExprWithVars(row.formulaExpr, varMap);

  const E = num(row.value);
  const K = num(row.convFactor);
  const I = num(row.surchargeMul);

  row.convQty = (K === 0 ? E : E*K);
  row.finalQty = (I === 0 ? row.convQty : row.convQty * I);
}

function recalcAll(){
  for(const tab of ["steel","steelSub","support"]){
    const pack = getPack(tab);
    if(!pack) continue;
    for(const sec of pack.sections){
      // sec 단위로 변수 계산 -> row 계산
      const tmpMap = {};
      // 섹션 vars 임시 계산
      let map = {};
      for(let pass=0; pass<5; pass++){
        for(const vr of (sec.vars || [])){
          const k = (vr.key ?? "").toString().trim().toUpperCase();
          if(!isValidVarName(k)) continue;
          const val = evalExprWithVars(vr.expr, map);
          vr.value = val;
          map[k] = val;
        }
      }
      for(const r of (sec.rows || [])){
        recalcRow(r, map);
      }
      Object.assign(tmpMap, map);
    }
  }
}

/* ===== readonly 즉시 갱신 (현재 섹션 기준) ===== */
function refreshReadonlyInSameRow(activeEl){
  const tabId = activeEl?.getAttribute?.("data-tab");
  if(!tabId) return;

  // vars/calc 모두 처리
  const area = activeEl.getAttribute("data-area"); // "vars" | "calc" | "codes"
  if(area === "calc"){
    const rows = getActiveRows(tabId);
    if(!rows) return;
    const rowIdx = Number(activeEl.getAttribute("data-row") || -1);
    if(rowIdx < 0) return;
    const r = rows[rowIdx];
    if(!r) return;

    const tr = activeEl.closest("tr");
    if(!tr) return;

    // readonly 순서: [name, spec, unit, value, convQty, finalQty, surchargeMul, convUnit]
    const ro = tr.querySelectorAll("input.cell.readonly");
    if(ro.length < 8) return;

    ro[0].value = r.name ?? "";
    ro[1].value = r.spec ?? "";
    ro[2].value = r.unit ?? "";
    ro[3].value = String(roundUp3(r.value));
    ro[4].value = String(roundUp3(r.convQty));
    ro[5].value = String(roundUp3(r.finalQty));
    ro[6].value = (r.surchargeMul === "" ? "" : String(r.surchargeMul));
    ro[7].value = r.convUnit ?? "";
    return;
  }

  if(area === "vars"){
    // vars는 값 column readonly 1개 갱신
    const vars = getActiveVars(tabId);
    if(!vars) return;

    const rowIdx = Number(activeEl.getAttribute("data-row") || -1);
    if(rowIdx < 0) return;
    const r = vars[rowIdx];
    if(!r) return;

    const tr = activeEl.closest("tr");
    if(!tr) return;
    const ro = tr.querySelector("input.var-value");
    if(ro) ro.value = String(roundUp3(r.value));
  }
}

/* ===== Tabs ===== */
const tabsDef = [
  { id:"codes", label:'코드(Ctrl+".")' },
  { id:"steel", label:"철골(Steel)" },
  { id:"steelTotal", label:"철골_집계" },
  { id:"steelSub", label:"철골_부자재" },
  { id:"support", label:"동바리(support)" },
  { id:"supportTotal", label:"동바리_집계" }
];
let activeTabId = "steel";

const $tabs = document.getElementById("tabs");
const $view = document.getElementById("view");

/* ===== Focus 기억(산출표용) ===== */
const lastFocusCell = {
  codes: { row: 0, col: 0 },
  steel: { row: 0, col: 0 },
  steelSub: { row: 0, col: 0 },
  support: { row: 0, col: 0 },
};

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

/* ===== Render tabs ===== */
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

/* ===== Cells registry ===== */
const cellRegistry = [];

function gridAttrString(gridAttrs){
  if(!gridAttrs) return "";
  const {tabId, rowIdx, colIdx, area} = gridAttrs;
  return `data-grid="1" data-area="${area}" data-tab="${tabId}" data-row="${rowIdx}" data-col="${colIdx}"`;
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
function readonlyCell(value, extraClass=""){
  return `<input class="cell readonly ${extraClass}" value="${escapeAttr(value ?? "")}" readonly />`;
}

/* ===== Wire cells ===== */
function wireCells(){
  document.querySelectorAll("[data-cell]").forEach(el=>{
    const id = el.getAttribute("data-cell");
    const meta = cellRegistry.find(x=>x.id===id);
    if(!meta) return;

    const handler = ()=>{
      meta.onChange(el.value);

      // ✅ vars 먼저 계산 -> calc 계산
      if(["steel","steelSub","support"].includes(activeTabId)){
        recalcVars(activeTabId);
        const {map} = recalcVars(activeTabId);
        // 현재 섹션 rows만 다시 계산
        const rows = getActiveRows(activeTabId) || [];
        rows.forEach(r=>recalcRow(r, map));
      }

      saveState();

      // readonly 즉시 갱신
      refreshReadonlyInSameRow(el);

      // calc 쪽은 영향을 많이 받으니 화면 잔상 최소화를 위해 부분 리렌더 대신 go
      // 단, 입력 중 과도 리렌더 방지: Enter/blur에서만 go 하도록
    };

    el.addEventListener("input", handler);
    el.addEventListener("change", handler);
    el.addEventListener("blur", ()=>{
      handler();
      // blur 시 전체 갱신(값/합계/readonly 안정화)
      go(activeTabId);
    });

    // Enter 동작
    el.addEventListener("keydown", (e)=>{
      if(e.key !== "Enter") return;

      const area = el.getAttribute("data-area");
      const col = Number(el.getAttribute("data-col") || -1);

      // textarea Enter는 그대로
      if(el.tagName.toLowerCase() === "textarea") return;

      // ✅ vars: expr(col=1)에서 Enter -> 값 계산 + 아래로
      if(area === "vars" && col === 1){
        e.preventDefault();
        handler();
        moveGridFrom(el, +1, 0);
        return;
      }

      // ✅ calc: 산출식(col=1) Enter -> 계산 + 아래로
      if(area === "calc" && col === 1){
        e.preventDefault();
        handler();
        moveGridFrom(el, +1, 0);
        return;
      }
    });
  });

  cellRegistry.length = 0;
}

/* ===== Focus tracking ===== */
function wireFocusTracking(){
  document.querySelectorAll('[data-grid="1"]').forEach(el=>{
    el.addEventListener("focus", ()=>{
      const tab = el.getAttribute("data-tab");
      const area = el.getAttribute("data-area");
      const row = Number(el.getAttribute("data-row") || 0);
      const col = Number(el.getAttribute("data-col") || 0);

      // calc 영역만 lastFocusCell에 기록(구분 이동 시 산출표 포커스 복원용)
      if(tab && lastFocusCell[tab] && area === "calc"){
        lastFocusCell[tab] = { row, col };
      }
    });
  });
}

/* ===== Grid nav ===== */
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

function focusGrid(tab, area, row, col){
  const selector = `[data-grid="1"][data-tab="${tab}"][data-area="${area}"][data-row="${row}"][data-col="${col}"]`;
  const el = document.querySelector(selector);
  if(el){
    el.focus();
    el.scrollIntoView?.({block:"nearest", inline:"nearest"});
    return true;
  }
  return false;
}
function moveGridFrom(el, dRow, dCol){
  const tab = el.getAttribute("data-tab");
  const area = el.getAttribute("data-area");
  const row = Number(el.getAttribute("data-row") || 0);
  const col = Number(el.getAttribute("data-col") || 0);

  const targetRow = row + dRow;
  const targetCol = col + dCol;

  if(focusGrid(tab, area, targetRow, targetCol)) return true;

  for(let offset=1; offset<=6; offset++){
    if(focusGrid(tab, area, targetRow, targetCol - offset)) return true;
    if(focusGrid(tab, area, targetRow, targetCol + offset)) return true;
  }
  return false;
}

/* ===== Insert calc row ===== */
function insertCalcRowBelowActive(){
  if(!["steel","steelSub","support"].includes(activeTabId)) return;
  const rows = getActiveRows(activeTabId);
  if(!rows) return;

  const {row, col} = lastFocusCell[activeTabId] ?? {row:0, col:0};
  const insertAt = Math.min((row|0) + 1, rows.length);

  rows.splice(insertAt, 0, makeEmptyCalcRow());
  saveState();
  go(activeTabId);
  setTimeout(()=>focusGrid(activeTabId, "calc", insertAt, col), 0);
}

/* =========================================================
   ✅ Render: "좌(구분+변수) + 우(산출표)"
========================================================= */
function renderSectionList(tabId, leftRail){
  const pack = getPack(tabId);
  if(!pack) return;

  const activeIdx = pack.activeIndex|0;

  const box = document.createElement("div");
  box.className = "rail-box";
  box.innerHTML = `<div class="rail-title">구분명 리스트 (↑/↓ 이동)</div>`;

  const list = document.createElement("div");
  list.className = "section-list";
  list.id = "sectionList";

  pack.sections.forEach((sec, i)=>{
    const title = (sec.label || "").trim() ? sec.label.trim() : `구분 ${i+1}`;
    const count = (sec.count ?? "").toString().trim();
    const meta = count !== "" ? `개소: ${count}` : "";

    const item = document.createElement("button");
    item.type = "button";
    item.className = "section-item" + (i===activeIdx ? " active" : "");
    item.setAttribute("data-sec-index", String(i));
    item.tabIndex = 0;
    item.innerHTML = `
      <div style="text-align:left;">
        <div style="font-weight:900;">${escapeHtml(title)}</div>
        <div class="meta">${escapeHtml(meta)}</div>
      </div>
      <div style="opacity:.75;font-weight:900;">${i===activeIdx ? "선택" : ""}</div>
    `;

    item.addEventListener("click", ()=>{
      setActiveSectionIndex(tabId, i);
      go(tabId);
      // 구분 전환 시 산출표 첫 행으로 포커스
      setTimeout(()=>focusGrid(tabId, "calc", 0, 0), 0);
    });

    list.appendChild(item);
  });

  box.appendChild(list);

  // 구분 편집(현재 섹션 label/count)
  const sec = getActiveSection(tabId);
  const editor = document.createElement("div");
  editor.className = "section-editor";
  editor.innerHTML = `
    <input id="secLabel" placeholder="구분명(예: 2층 바닥 철골보)" value="${escapeAttr(sec?.label ?? "")}">
    <input id="secCount" placeholder="개소(예: 0,1,2...)" value="${escapeAttr(sec?.count ?? "")}">
    <button class="smallbtn" id="btnSecAdd">구분 추가 (Ctrl+F3)</button>
    <button class="smallbtn" id="btnSecRename">저장</button>
  `;

  box.appendChild(editor);
  leftRail.appendChild(box);

  setTimeout(()=>{
    const $label = document.getElementById("secLabel");
    const $count = document.getElementById("secCount");
    const $add = document.getElementById("btnSecAdd");
    const $save = document.getElementById("btnSecRename");

    const saveMeta = ()=>{
      const cur = getActiveSection(tabId);
      if(!cur) return;
      cur.label = ($label?.value ?? "").toString();
      cur.count = ($count?.value ?? "").toString();
      saveState();
      go(tabId);
    };

    $add?.addEventListener("click", ()=>{
      addSection(tabId);
      go(tabId);
      setTimeout(()=>{
        // 새 섹션 리스트 아이템 포커스
        const idx = getPack(tabId)?.activeIndex ?? 0;
        document.querySelector(`button.section-item[data-sec-index="${idx}"]`)?.focus?.();
      }, 0);
    });

    $save?.addEventListener("click", saveMeta);
  }, 0);
}

function renderVarsTable(tabId, leftRail){
  const vars = getActiveVars(tabId) || [];
  const box = document.createElement("div");
  box.className = "rail-box";
  box.innerHTML = `<div class="rail-title">변수표 (A, AB, A1, AB1… 최대 3자)</div>`;

  const wrap = document.createElement("div");
  wrap.className = "var-tablewrap";

  const table = document.createElement("table");
  table.className = "var-table";
  table.innerHTML = `
    <thead>
      <tr>
        <th style="width:90px;">변수</th>
        <th>산식</th>
        <th style="width:110px;">값</th>
        <th style="width:120px;">비고</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;

  const tbody = table.querySelector("tbody");

  // vars 재계산
  const {map} = recalcVars(tabId);
  // 현재 섹션 rows도 변수 반영으로 재계산
  const rows = getActiveRows(tabId) || [];
  rows.forEach(r=>recalcRow(r, map));

  vars.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${inputCell(r.key, v=>{ r.key=v; }, "A / AB / A1", {tabId, area:"vars", rowIdx:idx, colIdx:0})}</td>
      <td>${inputCell(r.expr, v=>{ r.expr=v; }, "예: 1.2+0.3  (<>는 주석)", {tabId, area:"vars", rowIdx:idx, colIdx:1})}</td>
      <td>${readonlyCell(String(roundUp3(r.value)), "var-value")}</td>
      <td>${inputCell(r.note, v=>{ r.note=v; }, "", {tabId, area:"vars", rowIdx:idx, colIdx:2})}</td>
    `;
    // 값칸 readonly에 class 부여(갱신용)
    tr.querySelector("input.cell.readonly")?.classList.add("var-value");

    tbody.appendChild(tr);
  });

  wrap.appendChild(table);
  box.appendChild(wrap);

  const hint = document.createElement("div");
  hint.className = "var-hint";
  hint.innerHTML = `
    • 변수명은 영문으로 시작하고 최대 3자까지 가능합니다. (예: A, AB, A1, AB1, ABC)<br/>
    • 산식/산출식에서 <b>&lt;...&gt;</b> 안은 입력돼도 계산에 포함되지 않습니다.<br/>
    • 산출식에서 변수명을 쓰면 자동으로 값이 반영됩니다.
  `;
  box.appendChild(hint);

  leftRail.appendChild(box);
}

function renderCalcSheet(title, tabId, mode){
  $view.innerHTML = "";

  const layout = document.createElement("div");
  layout.className = "calc-layout";

  const leftRail = document.createElement("div");
  leftRail.className = "left-rail";

  const rightMain = document.createElement("div");
  rightMain.className = "right-main";

  // 좌측(구분 리스트 + 변수표)
  renderSectionList(tabId, leftRail);
  renderVarsTable(tabId, leftRail);

  // 우측(산출표)
  const desc = '구분(↑/↓) 이동 → 해당 구분의 변수/산출표로 전환 | 산출식 Enter 계산 | Ctrl+. 코드 선택 | Ctrl+F3 구분추가 | Ctrl+Shift+F3 행추가';
  const {wrap, header} = panel(title, desc);

  const rightBtns = document.createElement("div");
  rightBtns.style.display="flex";
  rightBtns.style.gap="8px";
  rightBtns.style.flexWrap="wrap";

  const addRowBtn = document.createElement("button");
  addRowBtn.className="smallbtn";
  addRowBtn.textContent="행 추가 (Ctrl+Shift+F3)";
  addRowBtn.onclick=()=>insertCalcRowBelowActive();

  const add10Btn = document.createElement("button");
  add10Btn.className="smallbtn";
  add10Btn.textContent="+10행";
  add10Btn.onclick=()=>{
    const rows = getActiveRows(tabId) || [];
    for(let i=0;i<10;i++) rows.push(makeEmptyCalcRow());
    saveState();
    go(tabId);
  };

  rightBtns.appendChild(addRowBtn);
  rightBtns.appendChild(add10Btn);
  header.appendChild(rightBtns);

  const rows = getActiveRows(tabId) || makeDefaultRows(20);

  // 변수맵(현재 섹션)
  const {map} = recalcVars(tabId);
  rows.forEach(r=>recalcRow(r, map));

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
          <th style="min-width:90px;">값</th>

          <th style="min-width:220px;">비고</th>
          <th style="min-width:90px;">할증(배수)</th>
          <th style="min-width:110px;">환산단위</th>
          <th style="min-width:120px;">환산계수</th>
          <th style="min-width:120px;">환산수량</th>
          <th style="min-width:130px;">할증후수량</th>
          <th style="min-width:120px;">작업</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;

  const tbody = tableWrap.querySelector("tbody");

  rows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${inputCell(r.code, v=>{ r.code=v; }, "코드 입력", {tabId, area:"calc", rowIdx:idx, colIdx:0})}</td>
      <td>${readonlyCell(r.name)}</td>
      <td>${readonlyCell(r.spec)}</td>
      <td>${readonlyCell(r.unit)}</td>

      <td>${inputCell(r.formulaExpr, v=>{ r.formulaExpr=v; }, "예: (A+0.5)*2  (<...>는 주석)", {tabId, area:"calc", rowIdx:idx, colIdx:1})}</td>
      <td>${readonlyCell(String(roundUp3(r.value)))}</td>

      <td>${textAreaCell(r.note, v=>{ r.note=v; }, {tabId, area:"calc", rowIdx:idx, colIdx:2})}</td>

      <td>${readonlyCell(r.surchargeMul === "" ? "" : String(r.surchargeMul))}</td>
      <td>${readonlyCell(r.convUnit)}</td>
      <td>${inputCell(r.convFactor, v=>{ r.convFactor=v; }, "비워도 됨", {tabId, area:"calc", rowIdx:idx, colIdx:3})}</td>
      <td>${readonlyCell(String(roundUp3(r.convQty)))}</td>
      <td>${readonlyCell(String(roundUp3(r.finalQty)))}</td>
      <td></td>
    `;

    // readonly 갱신용 순서 맞추기 위해 class로 구분
    const ro = tr.querySelectorAll("input.cell.readonly");
    // [name,spec,unit,value,surchargeMul,convUnit,convQty,finalQty] 순으로 쓰고 싶지만
    // refreshReadonlyInSameRow는 [name,spec,unit,value,convQty,finalQty,surchargeMul,convUnit]로 기대
    // -> 현재 DOM 순서를 그에 맞게 재배치할 수 없어, refresh에서 ro 인덱스를 맞춘 형태로 처리해 둠

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

  const sumBox = document.createElement("div");
  sumBox.style.marginTop="10px";
  const sumVal = (mode==="steel")
    ? rows.reduce((a,b)=> a + roundUp3(b.finalQty), 0)
    : rows.reduce((a,b)=> a + num(b.value), 0);
  sumBox.innerHTML = `<span class="badge">합계: ${roundUp3(sumVal)}</span>`;
  wrap.appendChild(sumBox);

  rightMain.appendChild(wrap);

  layout.appendChild(leftRail);
  layout.appendChild(rightMain);
  $view.appendChild(layout);

  wireCells();
  wireFocusTracking();
  wireMouseFocus();

  // 산출표 포커스 복원
  const last = lastFocusCell[tabId] ?? {row:0,col:0};
  setTimeout(()=>focusGrid(tabId, "calc", last.row, last.col), 0);

  // 구분 리스트 키보드 ↑/↓ 이동
  setTimeout(()=>wireSectionArrowNav(tabId), 0);
}

/* ===== 섹션 리스트 방향키 이동 ===== */
function wireSectionArrowNav(tabId){
  const list = document.getElementById("sectionList");
  if(!list) return;

  list.addEventListener("keydown", (e)=>{
    if(e.key !== "ArrowUp" && e.key !== "ArrowDown") return;

    e.preventDefault();

    const pack = getPack(tabId);
    if(!pack) return;

    const cur = pack.activeIndex|0;
    const next = e.key === "ArrowDown" ? Math.min(cur+1, pack.sections.length-1) : Math.max(cur-1, 0);

    if(next === cur) return;

    setActiveSectionIndex(tabId, next);
    go(tabId);

    setTimeout(()=>{
      document.querySelector(`button.section-item[data-sec-index="${next}"]`)?.focus?.();
      // 구분 이동 후 산출표는 새 값 준비(기본은 기존 입력 유지지만 섹션별 별도이므로 결과적으로 새 산출표)
      focusGrid(tabId, "calc", 0, 0);
    }, 0);
  });

  // 현재 선택 구분으로 포커스(처음)
  const idx = getPack(tabId)?.activeIndex ?? 0;
  document.querySelector(`button.section-item[data-sec-index="${idx}"]`)?.focus?.();
}

/* ===== Total helpers (모든 섹션 합산) ===== */
function flattenAllRows(tab){
  const pack = getPack(tab);
  if(!pack) return [];
  const out = [];
  for(const sec of pack.sections){
    out.push(...(sec.rows || []));
  }
  return out;
}
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
function renderSteelTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(철골+부자재 모든 구분의 할증후수량 합산)");
  recalcAll();

  const grouped = groupSum(
    [...flattenAllRows("steel"), ...flattenAllRows("steelSub")],
    r => roundUp3(r.finalQty)
  );

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
  const {wrap} = panel("동바리_집계(Support_Total quantity)", "코드별 합계(동바리 모든 구분의 물량(Value) 합계)");
  recalcAll();

  const grouped = groupSum(flattenAllRows("support"), r => num(r.value));

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

/* ===== Codes view (간단 유지) ===== */
function renderCodes(){
  $view.innerHTML = "";
  const {wrap, header} = panel('코드(Ctrl+".")', "코드 마스터(수정/추가). 엑셀 업로드(.xlsx)로 등록 가능.");

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가";
  addBtn.onclick = ()=>{
    const {row, col} = lastFocusCell.codes ?? {row:0,col:0};
    const insertAt = Math.min((row|0)+1, state.codes.length);
    state.codes.splice(insertAt, 0, makeEmptyCodeRow());
    saveState();
    go("codes");
    setTimeout(()=>focusGrid("codes", "codes", insertAt, col), 0);
  };

  right.appendChild(addBtn);
  header.appendChild(right);

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:70px;">No</th>
          <th style="min-width:170px;">코드</th>
          <th style="min-width:220px;">품명</th>
          <th style="min-width:220px;">규격</th>
          <th style="min-width:90px;">단위</th>
          <th style="min-width:90px;">할증</th>
          <th style="min-width:120px;">환산단위</th>
          <th style="min-width:140px;">환산계수</th>
          <th style="min-width:260px;">비고</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;
  const tbody = tableWrap.querySelector("tbody");

  state.codes.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${inputCell(r.code, v=>{ r.code=v; }, "", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:0})}</td>
      <td>${inputCell(r.name, v=>{ r.name=v; }, "", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:1})}</td>
      <td>${inputCell(r.spec, v=>{ r.spec=v; }, "", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:2})}</td>
      <td>${inputCell(r.unit, v=>{ r.unit=v; }, "", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:3})}</td>
      <td>${inputCell(r.surcharge, v=>{ r.surcharge=v; }, "예: 7", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:4})}</td>
      <td>${inputCell(r.conv_unit, v=>{ r.conv_unit=v; }, "", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:5})}</td>
      <td>${inputCell(r.conv_factor, v=>{ r.conv_factor=v; }, "", {tabId:"codes", area:"codes", rowIdx:idx, colIdx:6})}</td>
      <td>${textAreaCell(r.note, v=>{ r.note=v; }, {tabId:"codes", area:"codes", rowIdx:idx, colIdx:7})}</td>
    `;
    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);

  wireCells();
  wireFocusTracking();
  wireMouseFocus();
}

/* ===== Router ===== */
function go(id){
  activeTabId = id;
  recalcAll();
  saveState();
  renderTabs();

  if(id==="codes") renderCodes();
  else if(id==="steel") renderCalcSheet("철골(Steel)", "steel", "steel");
  else if(id==="steelSub") renderCalcSheet("철골_부자재(Processing and assembly)", "steelSub", "steel");
  else if(id==="support") renderCalcSheet("동바리(support)", "support", "support");
  else if(id==="steelTotal") renderSteelTotal();
  else if(id==="supportTotal") renderSupportTotal();
}

/* ===== Picker (Ctrl+.) ===== */
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
  const rows = getActiveRows(tab);
  if(!rows) return;

  const idx = Math.min(Math.max(0, Number(focusRow) || 0), rows.length);
  const insertAt = Math.min(idx + 1, rows.length);

  const newRows = codeList.map(code=>{
    const r = makeEmptyCalcRow();
    r.code = code;
    return r;
  });

  rows.splice(insertAt, 0, ...newRows);
  saveState();
  go(tab);
  setTimeout(()=>focusGrid(tab, "calc", insertAt, 0), 0);
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

  if(msg.type === "CLOSE_PICKER"){
    try{ pickerWin?.close(); }catch{}
    pickerWin = null;
    return;
  }
});

/* =========================
   HOTKEYS
   ========================= */
document.addEventListener("keydown", (e)=>{
  // Ctrl+. picker
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "." || e.code === "Period")){
    e.preventDefault();
    openPickerWindow();
    return;
  }

  // Ctrl+F3 : 구분 추가
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3" && !e.shiftKey){
    if(["steel","steelSub","support"].includes(activeTabId)){
      e.preventDefault();
      addSection(activeTabId);
      go(activeTabId);
      return;
    }
  }

  // Ctrl+Shift+F3 : 산출표 행 추가
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3" && e.shiftKey){
    if(["steel","steelSub","support"].includes(activeTabId)){
      e.preventDefault();
      insertCalcRowBelowActive();
      return;
    }
  }

  // Grid nav / edit mode
  const el = document.activeElement;
  if(!isGridEl(el)) return;

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

  if(isTextareaEl(el)){
    if(e.key === "ArrowUp"){ if(!textareaAtTop(el)) return; e.preventDefault(); moveGridFrom(el, -1, 0); return; }
    if(e.key === "ArrowDown"){ if(!textareaAtBottom(el)) return; e.preventDefault(); moveGridFrom(el, +1, 0); return; }
    if(e.key === "ArrowLeft"){ if(!caretAtStart(el)) return; e.preventDefault(); moveGridFrom(el, 0, -1); return; }
    if(e.key === "ArrowRight"){ if(!caretAtEnd(el)) return; e.preventDefault(); moveGridFrom(el, 0, +1); return; }
    return;
  }

  if(e.key === "ArrowUp"){ e.preventDefault(); moveGridFrom(el, -1, 0); return; }
  if(e.key === "ArrowDown"){ e.preventDefault(); moveGridFrom(el, +1, 0); return; }
  if(e.key === "ArrowLeft"){ e.preventDefault(); moveGridFrom(el, 0, -1); return; }
  if(e.key === "ArrowRight"){ e.preventDefault(); moveGridFrom(el, 0, +1); return; }
}, true);

/* ===== 상단 버튼들(기존 index.html에 존재한다고 가정) ===== */
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
    // 최소 보정
    if(!state.sheets) state = loadState();
    recalcAll();
    saveState();
    go("steel");
  }catch{
    alert("JSON 파싱 실패: 파일 내용을 확인해 주세요.");
  }finally{
    e.target.value = "";
  }
});

document.getElementById("btnReset")?.addEventListener("click", ()=>{
  if(!confirm("정말 초기화할까요? (로컬 저장 데이터가 삭제됩니다)")) return;
  localStorage.removeItem(STORAGE_KEY);
  state = makeState();
  recalcAll();
  saveState();
  go("steel");
});

/* ===== Boot ===== */
renderTabs();
go(activeTabId);
