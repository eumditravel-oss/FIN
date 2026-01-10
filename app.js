/* =========================
   FIN 산출자료(Web) app.js - FINAL (syntax fixed + section UI on TOP)
   - ✅ 방향키 이동(Excel-like) (capture)
   - ✅ F2 편집모드
   - ✅ 마우스로 td 클릭해도 셀 포커스 가능 (delegation 1회)
   - ✅ 구분(섹션) 리스트/변수표를 산출표 "위"에 배치
   - ✅ 구분 리스트(↑/↓) 이동 + 선택 시 산출표가 해당 구분 데이터로 전환
   - ✅ <...> 주석은 계산 제외
   ========================= */

const STORAGE_KEY = "FIN_WEB_V9";

/* ===== Seed Codes ===== */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
];

/* ===== State makers ===== */
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
function makeEmptySection(){
  return {
    label: "",
    count: "",
    vars: Array.from({length: 12}, makeEmptyVarRow), // 변수 12줄(원하면 늘려도 됨)
    rows: Array.from({length: 20}, makeEmptyCalcRow),
  };
}
function makeState(){
  return {
    codes: SEED_CODES,
    // ✅ 탭별 섹션(구분) 묶음
    sections: {
      steel:    [makeEmptySection()],
      steelSub: [makeEmptySection()],
      support:  [makeEmptySection()],
    },
    // ✅ 탭별 현재 선택된 섹션 인덱스
    activeSectionIndex: {
      steel: 0,
      steelSub: 0,
      support: 0,
    }
  };
}

/* ===== Load/Save ===== */
function loadState(){
  const raw = localStorage.getItem(STORAGE_KEY);
  if(!raw) return makeState();
  try{
    const s = JSON.parse(raw);

    // codes
    s.codes = Array.isArray(s.codes) ? s.codes : SEED_CODES;

    // sections
    if(!s.sections) s.sections = {};
    for(const k of ["steel","steelSub","support"]){
      if(!Array.isArray(s.sections[k]) || s.sections[k].length === 0){
        s.sections[k] = [makeEmptySection()];
      }else{
        // 보정
        s.sections[k] = s.sections[k].map(sec => ({
          label: (sec.label ?? "").toString(),
          count: (sec.count ?? "").toString(),
          vars: Array.isArray(sec.vars) ? sec.vars.map(v=>({
            key:(v.key??"").toString(),
            expr:(v.expr??"").toString(),
            value:Number(v.value||0)||0,
            note:(v.note??"").toString(),
          })) : Array.from({length:12}, makeEmptyVarRow),
          rows: Array.isArray(sec.rows) ? sec.rows : Array.from({length:20}, makeEmptyCalcRow),
        }));
      }
    }

    if(!s.activeSectionIndex) s.activeSectionIndex = {steel:0, steelSub:0, support:0};
    for(const k of ["steel","steelSub","support"]){
      const n = Number(s.activeSectionIndex[k] ?? 0);
      s.activeSectionIndex[k] = Number.isFinite(n) ? Math.max(0, Math.min(n, s.sections[k].length-1)) : 0;
    }

    return s;
  }catch{
    return makeState();
  }
}

let state = loadState();
function saveState(){ localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }

/* ===== Utils ===== */
function escapeHtml(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;");
}
function escapeAttr(s){ return escapeHtml(s).replaceAll("\n","&#10;"); }

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
function surchargeToMul(p){
  const x = num(p);
  return x ? (1 + x/100) : "";
}
function normalizeKey(s){
  return (s ?? "").toString().trim().toUpperCase();
}
function isValidVarName(name){
  // ✅ A, AB, A1, AB1 ... 최대 3자(영문 시작, 영문/숫자 조합)
  const n = normalizeKey(name);
  return /^[A-Z][A-Z0-9]{0,2}$/.test(n);
}

/* ===== Find code ===== */
function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}

/* ===== Variable evaluation ===== */
function evalExprWithVars(exprRaw, varMap){
  const withoutTags = (exprRaw ?? "").toString().replace(/<[^>]*>/g, "");
  let s = withoutTags.trim();
  if(!s) return 0;

  // 변수명 치환(가장 긴 키부터)
  const keys = Array.from(varMap.keys()).sort((a,b)=>b.length-a.length);
  for(const k of keys){
    const v = varMap.get(k);
    const safe = String(Number(v)||0);
    // 경계: 알파벳/숫자 덩어리로만 매칭
    s = s.replace(new RegExp(`\\b${k}\\b`, "g"), safe);
  }

  // 숫자/연산자만 허용
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

function buildVarMap(tabId){
  const idx = state.activeSectionIndex[tabId] ?? 0;
  const sec = state.sections[tabId]?.[idx];
  const map = new Map();
  if(!sec) return map;

  for(const v of (sec.vars ?? [])){
    const k = normalizeKey(v.key);
    if(!k) continue;
    if(!isValidVarName(k)) continue;
    map.set(k, Number(v.value)||0);
  }
  return map;
}

function recalcRow(row, tabId){
  const m = findCode(row.code);
  if(m){
    row.name = m.name ?? "";
    row.spec = m.spec ?? "";
    row.unit = m.unit ?? "";
    const mul = surchargeToMul(m.surcharge);
    row.surchargeMul = mul === "" ? "" : mul;
    row.convUnit = m.conv_unit ?? "";
    if((row.convFactor ?? "").toString().trim() === "") row.convFactor = (m.conv_factor ?? "");
  }else{
    // 코드가 없으면 자동필드 유지(사용자 입력/공백)
  }

  const varMap = buildVarMap(tabId);
  row.value = evalExprWithVars(row.formulaExpr, varMap);

  const E = num(row.value);
  const K = num(row.convFactor);
  const I = num(row.surchargeMul);

  row.convQty = (K === 0 ? E : E*K);
  row.finalQty = (I === 0 ? row.convQty : row.convQty * I);
}

function recalcAll(){
  for(const tabId of ["steel","steelSub","support"]){
    const secIdx = state.activeSectionIndex[tabId] ?? 0;
    const sec = state.sections[tabId]?.[secIdx];
    if(!sec) continue;

    // 1) 변수 먼저 계산 (변수 expr는 "산식"으로 계산)
    const varMap = new Map();
    for(const v of sec.vars){
      const k = normalizeKey(v.key);
      if(k && isValidVarName(k)) varMap.set(k, Number(v.value)||0);
    }
    // 변수 expr -> value 재계산(순서대로, 앞에서 계산된 변수 사용 가능)
    for(const v of sec.vars){
      const k = normalizeKey(v.key);
      if(!k || !isValidVarName(k)) continue;
      const val = evalExprWithVars(v.expr, varMap);
      v.value = val;
      varMap.set(k, val);
    }

    // 2) 산출표 계산
    sec.rows.forEach(r => recalcRow(r, tabId));
  }
}

/* ===== Focus grid (Excel-like) ===== */
const lastFocusCell = {
  codes: { row: 0, col: 0 },
  steel: { row: 0, col: 0 },
  steelSub: { row: 0, col: 0 },
  support: { row: 0, col: 0 },
};
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

function focusGrid(tab, row, col){
  const selector = `[data-grid="1"][data-tab="${tab}"][data-row="${row}"][data-col="${col}"]`;
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

/* ===== DOM ===== */
const $tabs = document.getElementById("tabs");
const $view = document.getElementById("view");

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

/* ===== Panels ===== */
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

/* ===== Cells registry ===== */
const cellRegistry = [];
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

/* ===== Wire cells ===== */
function refreshReadonlyInSameRow(activeEl){
  const tabId = activeEl?.getAttribute?.("data-tab");
  if(!tabId) return;
  if(tabId === "codes") return;

  const rowIdx = Number(activeEl.getAttribute("data-row") || -1);
  if(rowIdx < 0) return;

  const secIdx = state.activeSectionIndex[tabId] ?? 0;
  const sec = state.sections[tabId]?.[secIdx];
  if(!sec) return;
  const r = sec.rows?.[rowIdx];
  if(!r) return;

  const tr = activeEl.closest("tr");
  if(!tr) return;

  const ro = tr.querySelectorAll("input.cell.readonly");
  // readonly 순서: [name, spec, unit, value, surchargeMul, convUnit, convFactor, convQty, finalQty]
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

function wireCells(){
  document.querySelectorAll("[data-cell]").forEach(el=>{
    const id = el.getAttribute("data-cell");
    const meta = cellRegistry.find(x=>x.id===id);
    if(!meta) return;

    const handler = ()=>{
      meta.onChange(el.value);
      recalcAll();
      saveState();
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
      if(col !== 1) return;

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

/* ===== Mouse click -> focus cell (delegation, once) ===== */
let mouseFocusWired = false;
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

/* ===== Rows helpers ===== */
function getActiveSection(tabId){
  const idx = state.activeSectionIndex[tabId] ?? 0;
  const list = state.sections[tabId] ?? [];
  return { idx, list, sec: list[idx] };
}

function ensureSection(tabId){
  if(!state.sections[tabId] || state.sections[tabId].length === 0){
    state.sections[tabId] = [makeEmptySection()];
    state.activeSectionIndex[tabId] = 0;
  }
}

function insertSectionBelow(tabId){
  ensureSection(tabId);
  const { idx, list } = getActiveSection(tabId);
  const insertAt = Math.min(idx + 1, list.length);
  list.splice(insertAt, 0, makeEmptySection());
  state.activeSectionIndex[tabId] = insertAt;
  saveState();
}

function deleteActiveSection(tabId){
  ensureSection(tabId);
  const { idx, list } = getActiveSection(tabId);
  if(list.length <= 1){
    // 최소 1개 유지
    list[0] = makeEmptySection();
    state.activeSectionIndex[tabId] = 0;
  }else{
    list.splice(idx, 1);
    state.activeSectionIndex[tabId] = Math.max(0, idx - 1);
  }
  saveState();
}

function insertRowBelowActive(){
  if(activeTabId === "codes"){
    const {row, col} = lastFocusCell.codes ?? {row:0,col:0};
    const insertAt = Math.min(row + 1, state.codes.length);
    state.codes.splice(insertAt, 0, makeEmptyCodeRow());
    saveState();
    go("codes");
    setTimeout(()=>focusGrid("codes", insertAt, col), 0);
    return;
  }

  if(!["steel","steelSub","support"].includes(activeTabId)) return;

  ensureSection(activeTabId);
  const { sec } = getActiveSection(activeTabId);
  if(!sec) return;

  const {row, col} = lastFocusCell[activeTabId] ?? {row:0,col:0};
  const insertAt = Math.min(row + 1, sec.rows.length);
  sec.rows.splice(insertAt, 0, makeEmptyCalcRow());
  saveState();
  go(activeTabId);
  setTimeout(()=>focusGrid(activeTabId, insertAt, col), 0);
}

/* ===== Render Tabs ===== */
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

/* ===== Render Codes ===== */
function renderCodes(){
  $view.innerHTML = "";
  const {wrap, header} = panel('코드(Ctrl+".")', "코드 마스터(수정/추가). 엑셀 업로드(.xlsx)로 한 번에 등록 가능. (행 추가: Ctrl+F3)");

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가 (Ctrl+F3)";
  addBtn.onclick = ()=> insertRowBelowActive();

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

/* ===== Render Calc Sheet (with section+vars on TOP) ===== */
function renderCalcSheet(title, tabId){
  $view.innerHTML = "";

  ensureSection(tabId);
  recalcAll();

  const { idx, list, sec } = getActiveSection(tabId);

  // ✅ (초록) 위 영역 + (노란) 아래 산출표 형태로 쌓기
  const layout = document.createElement("div");
  layout.className = "calc-layout";

  const left = document.createElement("div");
  left.className = "left-rail";

  const right = document.createElement("div");
  right.className = "right-main";

  /* ---------- 구분명 리스트 ---------- */
  const box1 = document.createElement("div");
  box1.className = "rail-box";

  const sectionItems = list.map((s,i)=>{
    const name = (s.label ?? "").trim() || `구분 ${i+1}`;
    return `
      <div class="section-item ${i===idx ? "active":""}" tabindex="0" data-sec="${i}">
        <div>
          <div>${escapeHtml(name)}</div>
          <div class="meta">개소: ${escapeHtml((s.count ?? "").toString())}</div>
        </div>
        <div class="meta">${i===idx ? "선택" : ""}</div>
      </div>
    `;
  }).join("");

  box1.innerHTML = `
    <div class="rail-title">구분명 리스트 (↑/↓ 이동)</div>
    <div class="section-list" id="sectionList">${sectionItems}</div>

    <div class="section-editor">
      <input id="secLabel" placeholder="구분명(예: 2층 바닥 철골보)" value="${escapeAttr(sec.label)}" />
      <input id="secCount" placeholder="개소(예: 0,1,2...)" value="${escapeAttr(sec.count)}" />
      <button class="smallbtn" id="btnSecSave">저장</button>
    </div>

    <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
      <button class="smallbtn" id="btnSecAdd">구분 추가 (Ctrl+F3)</button>
      <button class="smallbtn" id="btnSecDel">구분 삭제</button>
    </div>
  `;

  /* ---------- 변수표 ---------- */
  const box2 = document.createElement("div");
  box2.className = "rail-box";

  const varRowsHtml = sec.vars.map((v,i)=>{
    return `
      <tr>
        <td>${inputCell(v.key, val=>{ v.key = val; }, "A/AB/A1", null, `data-var="key" data-vi="${i}"`)}</td>
        <td>${inputCell(v.expr, val=>{ v.expr = val; }, "산식", null, `data-var="expr" data-vi="${i}"`)}</td>
        <td><input class="cell readonly" value="${escapeAttr(String(roundUp3(v.value)))}" readonly /></td>
        <td>${inputCell(v.note, val=>{ v.note = val; }, "비고", null, `data-var="note" data-vi="${i}"`)}</td>
      </tr>
    `;
  }).join("");

  box2.innerHTML = `
    <div class="rail-title">변수표 (A, AB, A1, AB1... 최대 3자)</div>
    <div class="var-tablewrap">
      <table class="var-table">
        <thead>
          <tr><th>변수</th><th>산식</th><th>값</th><th>비고</th></tr>
        </thead>
        <tbody>${varRowsHtml}</tbody>
      </table>
    </div>
    <div class="var-hint">
      • 변수명은 영문으로 시작, 최대 3자 (예: A, AB, A1)<br/>
      • 산식에 변수 사용 가능 (예: (A+0.5)*2 )<br/>
      • &lt;...&gt; 안은 주석으로 계산 제외
    </div>
  `;

  left.appendChild(box1);
  left.appendChild(box2);

  /* ---------- (노란) 산출표 ---------- */
  const desc = '구분(↑/↓) 이동 → 해당 구분의 변수/산출표로 전환 | 산출식 Enter 계산 | Ctrl+. 코드선택 | Ctrl+F3 구분추가/행추가';
  const {wrap, header} = panel(title, desc);

  const headerRight = document.createElement("div");
  headerRight.style.display="flex"; headerRight.style.gap="8px"; headerRight.style.flexWrap="wrap";

  const addRowBtn = document.createElement("button");
  addRowBtn.className="smallbtn";
  addRowBtn.textContent="행 추가 (Shift+Ctrl+F3)";
  addRowBtn.onclick=()=>insertRowBelowActive();

  const add10Btn = document.createElement("button");
  add10Btn.className="smallbtn";
  add10Btn.textContent="+10행";
  add10Btn.onclick=()=>{
    sec.rows.push(...Array.from({length:10}, makeEmptyCalcRow));
    saveState();
    go(tabId);
  };

  headerRight.appendChild(addRowBtn);
  headerRight.appendChild(add10Btn);
  header.appendChild(headerRight);

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
          <th style="min-width:80px;">물량(Value)</th>
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

  const tbody = tableWrap.querySelector("tbody");

  sec.rows.forEach((r, idxRow)=>{
    recalcRow(r, tabId);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idxRow+1}</td>
      <td>${inputCell(r.code, v=>{ r.code=v; }, "코드 입력", {tabId, rowIdx:idxRow, colIdx:0})}</td>
      <td>${readonlyCell(r.name)}</td>
      <td>${readonlyCell(r.spec)}</td>
      <td>${readonlyCell(r.unit)}</td>
      <td>${inputCell(r.formulaExpr, v=>{ r.formulaExpr=v; }, "예: (A+0.5)*2  (<...>는 주석)", {tabId, rowIdx:idxRow, colIdx:1})}</td>
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
    dup.onclick=()=>{
      sec.rows.splice(idxRow+1, 0, JSON.parse(JSON.stringify(r)));
      saveState(); go(tabId);
    };

    const del = document.createElement("button");
    del.className="smallbtn"; del.textContent="삭제";
    del.onclick=()=>{
      sec.rows.splice(idxRow, 1);
      saveState(); go(tabId);
    };

    act.appendChild(dup);
    act.appendChild(del);
    tdAct.appendChild(act);

    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);

  // 합계
  const sumBox = document.createElement("div");
  sumBox.style.marginTop="10px";
  const sumVal = (tabId==="support")
    ? sec.rows.reduce((a,b)=> a + num(b.value), 0)
    : sec.rows.reduce((a,b)=> a + roundUp3(b.finalQty), 0);
  sumBox.innerHTML = `<span class="badge">합계: ${roundUp3(sumVal)}</span>`;
  wrap.appendChild(sumBox);

  right.appendChild(wrap);

  layout.appendChild(left);
  layout.appendChild(right);
  $view.appendChild(layout);

  /* ---------- 섹션 이벤트 ---------- */
  const $sectionList = document.getElementById("sectionList");
  const $label = document.getElementById("secLabel");
  const $count = document.getElementById("secCount");

  function setActiveSection(nextIdx){
    const max = state.sections[tabId].length - 1;
    const ni = Math.max(0, Math.min(nextIdx, max));
    state.activeSectionIndex[tabId] = ni;
    saveState();
    go(tabId);
  }

  $sectionList?.addEventListener("click", (e)=>{
    const item = e.target.closest("[data-sec]");
    if(!item) return;
    const i = Number(item.getAttribute("data-sec"));
    setActiveSection(i);
  });

  // ✅ 섹션 리스트에서 ↑/↓로 이동
  $sectionList?.addEventListener("keydown", (e)=>{
    if(e.key === "ArrowUp"){
      e.preventDefault();
      setActiveSection(idx - 1);
    }
    if(e.key === "ArrowDown"){
      e.preventDefault();
      setActiveSection(idx + 1);
    }
  });

  document.getElementById("btnSecSave")?.addEventListener("click", ()=>{
    sec.label = ($label?.value ?? "").toString();
    sec.count = ($count?.value ?? "").toString();
    saveState();
    go(tabId);
  });

  document.getElementById("btnSecAdd")?.addEventListener("click", ()=>{
    insertSectionBelow(tabId);
    go(tabId);
  });

  document.getElementById("btnSecDel")?.addEventListener("click", ()=>{
    if(!confirm("현재 구분을 삭제할까요?")) return;
    deleteActiveSection(tabId);
    go(tabId);
  });

  /* ---------- 변수표 입력 이벤트: 별도 바인딩(간단) ---------- */
  // 변수표는 data-cell registry가 아니라 단순 inputCell로 만들었으므로 wireCells로도 들어감
  // (inputCell로 만들었기 때문에 wireCells가 자동으로 값 반영 + recalcAll 수행)

  wireCells();
  wireFocusTracking();
  wireMouseFocus();

  const last = lastFocusCell[tabId] ?? {row:0,col:0};
  setTimeout(()=>focusGrid(tabId, last.row, last.col), 0);
}

/* ===== Group sum (Totals) ===== */
function groupSumCalc(rows, valueSelector){
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
  recalcAll();
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(철골+부자재의 할증후수량 합산)");

  const rowsAll = [];
  for(const tabId of ["steel","steelSub"]){
    ensureSection(tabId);
    for(const sec of state.sections[tabId]){
      sec.rows.forEach(r=>rowsAll.push(r));
    }
  }
  const grouped = groupSumCalc(rowsAll, r => roundUp3(r.finalQty));

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
  recalcAll();
  const {wrap} = panel("동바리_집계(Support_Total quantity)", "코드별 합계(동바리의 물량(Value) 합계)");

  const rowsAll = [];
  ensureSection("support");
  for(const sec of state.sections.support){
    sec.rows.forEach(r=>rowsAll.push(r));
  }
  const grouped = groupSumCalc(rowsAll, r => num(r.value));

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

/* ===== Router ===== */
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

/* ===== Picker window (Ctrl+.) ===== */
let pickerWin = null;

function openPickerWindow(){
  let origin = activeTabId;
  if(!["steel","steelSub","support"].includes(origin)) origin = "steel";

  // 현재 섹션의 포커스 행
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
  ensureSection(tab);
  const { sec } = getActiveSection(tab);
  if(!sec) return;

  const idx = Math.min(Math.max(0, Number(focusRow) || 0), sec.rows.length);
  const insertAt = Math.min(idx + 1, sec.rows.length);

  const newRows = codeList.map(code=>{
    const r = makeEmptyCalcRow();
    r.code = code;
    return r;
  });

  sec.rows.splice(insertAt, 0, ...newRows);
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

  if(msg.type === "CLOSE_PICKER"){
    try{ pickerWin?.close(); }catch{}
    pickerWin = null;
    return;
  }
});

/* ===== Hotkeys + Excel nav ===== */
let editMode = false;

function setEditingClass(on){
  document.querySelectorAll('.cell.editing').forEach(x=>x.classList.remove('editing'));
  if(on){
    const el = document.activeElement;
    if(el && el.classList && el.classList.contains("cell")) el.classList.add("editing");
  }
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

  ensureSection(activeTabId);
  const { sec } = getActiveSection(activeTabId);
  if(!sec) return;
  if(r >= sec.rows.length) return;

  sec.rows.splice(r, 1);
  saveState();
  go(activeTabId);

  const newRow = Math.min(r, sec.rows.length - 1);
  if(newRow >= 0) setTimeout(()=>focusGrid(activeTabId, newRow, 0), 0);
}

document.addEventListener("keydown", (e)=>{
  // Ctrl+. picker
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "." || e.code === "Period")){
    e.preventDefault();
    openPickerWindow();
    return;
  }

  // Ctrl+Delete delete row
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "Delete" || e.code === "Delete")){
    e.preventDefault();
    deleteRowAtActiveFocus();
    return;
  }

  // Ctrl+F3 : 구분 추가(steel/steelSub/support에서)
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3"){
    if(["steel","steelSub","support"].includes(activeTabId)){
      e.preventDefault();
      insertSectionBelow(activeTabId);
      go(activeTabId);
      return;
    }
  }

  // Ctrl+Shift+F3 : 행 추가
  if(e.ctrlKey && e.shiftKey && !e.altKey && !e.metaKey && e.code === "F3"){
    e.preventDefault();
    insertRowBelowActive();
    return;
  }

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
    // 최소 보정
    state = loadState(); // 안전하게 다시 로드/보정 로직 태우기 위해 저장 후 재로딩 방식 사용
    saveState();
    state = loadState();
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
