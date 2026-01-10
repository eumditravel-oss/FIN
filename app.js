/* =========================
   FIN 산출자료(Web) app.js (최종 통합본)
   - ✅ 탭: 코드 / 철골 / 철골_부자재 / 동바리 / 집계
   - ✅ (요청사항 반영)
     1) "구분(섹션)"은 산출표 위에서 관리 (Ctrl+F3 = 구분 추가)
     2) 구분 선택(↑↓) → 해당 구분의 변수표/산출표가 독립적으로 전환
     3) 변수표: 변수명(영문 시작, 최대 3자: A, AB, A1, AB1 등) / 산식 / 값 / 비고
        - 산식 Enter → 값 계산
        - 산출식에서 변수 사용 가능: 예) (A+0.5)*2
     4) 산출식의 "<...>" 는 주석(계산 제외)
     5) 방향키/마우스 포커스 먹통 방지:
        - 엑셀식 방향키 이동은 "산출표(.cell + data-grid=1)"에만 적용
        - 변수표/구분편집 입력은 data-grid를 안 달아서 "A" 입력으로 강제 이동되는 문제 방지
   ========================= */

const STORAGE_KEY = "FIN_WEB_V10";

/* ===== Seed Codes ===== */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
];

/* =========================
   State 구조
   state = {
     codes: [...],
     tabs: {
       steel:   { activeSection: 0, sections:[ {name,count, vars:[...], rows:[...]}, ... ] },
       steelSub:{ ... },
       support: { ... },
     }
   }
   ========================= */

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
  return { key:"", expr:"", val:0, note:"" };
}
function makeSection(name="구분", count=""){
  return {
    name,
    count,
    vars: Array.from({length: 12}, makeEmptyVarRow),
    rows: Array.from({length: 20}, makeEmptyCalcRow),
  };
}
function makeTabState(){
  return { activeSection: 0, sections: [ makeSection("구분 1","") ] };
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
    if(!Array.isArray(s.codes)) s.codes = SEED_CODES;

    if(!s.tabs) s.tabs = {};
    for(const k of ["steel","steelSub","support"]){
      if(!s.tabs[k]) s.tabs[k] = makeTabState();
      if(!Array.isArray(s.tabs[k].sections) || s.tabs[k].sections.length===0){
        s.tabs[k].sections = [ makeSection("구분 1","") ];
      }
      if(typeof s.tabs[k].activeSection !== "number") s.tabs[k].activeSection = 0;

      // 보정
      s.tabs[k].sections.forEach(sec=>{
        if(!Array.isArray(sec.vars)) sec.vars = Array.from({length: 12}, makeEmptyVarRow);
        if(!Array.isArray(sec.rows)) sec.rows = Array.from({length: 20}, makeEmptyCalcRow);
        if(typeof sec.name !== "string") sec.name = "구분";
        if(typeof sec.count !== "string") sec.count = "";
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

/* ===== Variables =====
   - 변수명 규칙: 영문 시작, 영문/숫자 조합, 최대 3자
   - 산식에서 변수 토큰을 치환하여 계산
*/
const VAR_KEY_RE = /^[A-Za-z][A-Za-z0-9]{0,2}$/;
const VAR_TOKEN_RE = /\b[A-Za-z][A-Za-z0-9]{0,2}\b/g;

function stripComments(exprRaw){
  return (exprRaw ?? "").toString().replace(/<[^>]*>/g, ""); // <...> 제거
}

function safeEvalNumeric(expr){
  const s = stripComments(expr).trim();
  if(!s) return 0;
  // 숫자/연산자/괄호/공백/소수점/콤마만(변수 치환 후)
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

// 섹션 변수값 계산 (상호 참조 가능, 순환 방지)
function computeVars(section){
  const mapExpr = new Map();
  section.vars.forEach(v=>{
    const key = (v.key ?? "").toString().trim();
    if(VAR_KEY_RE.test(key)){
      mapExpr.set(key, (v.expr ?? "").toString());
    }
  });

  const memo = new Map();
  const visiting = new Set();

  const evalVar = (key)=>{
    if(memo.has(key)) return memo.get(key);
    if(visiting.has(key)) return 0; // cycle -> 0
    visiting.add(key);

    const raw = mapExpr.get(key) ?? "";
    // 다른 변수 토큰 치환
    const replaced = stripComments(raw).replace(VAR_TOKEN_RE, (tok)=>{
      if(tok === key) return "0";
      if(mapExpr.has(tok)) return String(evalVar(tok));
      return tok; // 숫자면 그대로(하지만 regex상 변수형태만 매치됨)
    });

    const val = safeEvalNumeric(replaced);
    memo.set(key, val);
    visiting.delete(key);
    return val;
  };

  // 섹션 vars의 표시용 val 업데이트
  section.vars.forEach(v=>{
    const key = (v.key ?? "").toString().trim();
    if(VAR_KEY_RE.test(key) && mapExpr.has(key)){
      v.val = roundUp3(evalVar(key));
    }else{
      v.val = 0;
    }
  });

  // 계산용 치환 맵
  const outMap = {};
  for(const [k] of mapExpr.entries()){
    outMap[k] = memo.has(k) ? memo.get(k) : evalVar(k);
  }
  return outMap;
}

/* ===== Codes ===== */
function surchargeToMul(p){ const x = num(p); return x ? (1 + x/100) : ""; }
function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}

/* ===== Row calc ===== */
function evalExprWithVars(exprRaw, varMap){
  const raw = (exprRaw ?? "").toString();
  const without = stripComments(raw);

  // 변수 치환
  const replaced = without.replace(VAR_TOKEN_RE, (tok)=>{
    if(Object.prototype.hasOwnProperty.call(varMap, tok)){
      return String(varMap[tok] ?? 0);
    }
    return tok; // 미정의 변수는 그대로(=> safeEval에서 걸러져 0 처리될 수 있음)
  });

  // 치환 후 안전 계산 (변수 없는 경우도 여기서)
  return safeEvalNumeric(replaced);
}

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
  }else{
    // 코드 비어있거나 없으면 자동값은 건드리지 않음
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

function recalcTab(tabKey){
  const tab = state.tabs[tabKey];
  if(!tab) return;

  tab.sections.forEach(sec=>{
    const varMap = computeVars(sec);
    sec.rows.forEach(r=>recalcRow(r, varMap));
  });
}

function recalcAll(){
  recalcTab("steel");
  recalcTab("steelSub");
  recalcTab("support");
}

/* =========================
   DOM
   ========================= */
const $tabs = document.getElementById("tabs");
const $view = document.getElementById("view");

/* ===== Tabs def ===== */
const tabsDef = [
  { id:"codes", label:'코드(Ctrl+".")' },
  { id:"steel", label:"철골(Steel)" },
  { id:"steelTotal", label:"철골_집계" },
  { id:"steelSub", label:"철골_부자재" },
  { id:"support", label:"동바리(support)" },
  { id:"supportTotal", label:"동바리_집계" }
];
let activeTabId = "steel";

/* ===== Focus tracking for calc grid only ===== */
const lastFocusCell = {
  codes: { row: 0, col: 0 },
  steel: { row: 0, col: 0 },
  steelSub: { row: 0, col: 0 },
  support: { row: 0, col: 0 },
};

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

/* ===== Cells registry (calc/codes only) ===== */
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

/* ===== Mouse focus (delegation, once) ===== */
let mouseFocusWired = false;
function wireMouseFocus(){
  if(mouseFocusWired) return;
  mouseFocusWired = true;

  document.addEventListener("click", (e)=>{
    const sel = window.getSelection?.();
    if(sel && !sel.isCollapsed) return;

    const t = e.target;

    // UI 요소는 제외
    if(t?.closest?.("button,a,label,select,option")) return;

    // input 직접 클릭은 그대로
    if(t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA")) return;

    // td 클릭만 처리
    const td = t?.closest?.("td");
    if(!td) return;

    const cell = td.querySelector("input.cell:not(.readonly), textarea.cell");
    if(!cell) return;

    editMode = false;
    setEditingClass(false);
    cell.focus();
  });
}

/* ===== wireCells (calc/codes only) ===== */
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
  document.querySelectorAll('[data-grid="1"].cell').forEach(el=>{
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

/* ===== Grid navigation helpers (calc/codes only) ===== */
function focusGrid(tab, row, col){
  const selector = `[data-grid="1"][data-tab="${tab}"][data-row="${row}"][data-col="${col}"].cell`;
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

/* ===== readonly 즉시 갱신 (같은 tr 기준) ===== */
function refreshReadonlyInSameRow(activeEl){
  const tabId = activeEl?.getAttribute?.("data-tab");
  if(!tabId) return;
  if(tabId === "codes") return;

  const rowIdx = Number(activeEl.getAttribute("data-row") || -1);
  if(rowIdx < 0) return;

  const tabKey = tabId;
  const sec = getActiveSection(tabKey);
  if(!sec) return;

  const r = sec.rows[rowIdx];
  if(!r) return;

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
   Section / Var UI helpers
   ========================= */
function getTabState(tabKey){
  return state.tabs[tabKey];
}
function getActiveSection(tabKey){
  const tab = getTabState(tabKey);
  if(!tab) return null;
  const idx = Math.min(Math.max(0, tab.activeSection), tab.sections.length-1);
  tab.activeSection = idx;
  return tab.sections[idx];
}
function setActiveSection(tabKey, idx){
  const tab = getTabState(tabKey);
  if(!tab) return;
  tab.activeSection = Math.min(Math.max(0, idx), tab.sections.length-1);
  saveState();
  go(tabKey);
}

function addSection(tabKey){
  const tab = getTabState(tabKey);
  if(!tab) return;
  const n = tab.sections.length + 1;
  tab.sections.push(makeSection(`구분 ${n}`, ""));
  tab.activeSection = tab.sections.length - 1;
  saveState();
  go(tabKey);
}

function deleteActiveSection(tabKey){
  const tab = getTabState(tabKey);
  if(!tab) return;
  if(tab.sections.length <= 1){
    alert("구분은 최소 1개 이상이어야 합니다.");
    return;
  }
  const idx = tab.activeSection;
  const ok = confirm("선택된 구분을 삭제할까요?");
  if(!ok) return;
  tab.sections.splice(idx,1);
  tab.activeSection = Math.max(0, Math.min(tab.activeSection, tab.sections.length-1));
  saveState();
  go(tabKey);
}

/* ===== Rows insert (Shift+Ctrl+F3) ===== */
function addRowToActiveSection(tabKey, count=1){
  const sec = getActiveSection(tabKey);
  if(!sec) return;
  for(let i=0;i<count;i++) sec.rows.push(makeEmptyCalcRow());
  saveState();
  go(tabKey);
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
    const {row, col} = lastFocusCell.codes ?? {row:0,col:0};
    const insertAt = Math.min((row ?? 0) + 1, state.codes.length);
    state.codes.splice(insertAt, 0, makeEmptyCodeRow());
    saveState();
    go("codes");
    setTimeout(()=>focusGrid("codes", insertAt, col ?? 0), 0);
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

/* ===== TOP SPLIT (구분/변수) ===== */
function renderTopSplit(tabKey){
  const tab = getTabState(tabKey);
  const sec = getActiveSection(tabKey);
  if(!tab || !sec) return;

  const top = document.createElement("div");
  top.className = "top-split";

  // LEFT: section list
  const leftBox = document.createElement("div");
  leftBox.className = "rail-box";
  leftBox.innerHTML = `
    <div class="rail-title">구분명 리스트 (↑/↓ 이동)</div>
    <div class="section-list" id="secList"></div>

    <div class="section-editor" style="margin-top:10px;">
      <input id="secName" placeholder="구분명(예: 2층 바닥 철골보)" value="${escapeAttr(sec.name)}">
      <input id="secCount" placeholder="개소(예: 0,1,2...)" value="${escapeAttr(sec.count)}">
      <button class="smallbtn" id="secSave">저장</button>
    </div>

    <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
      <button class="smallbtn" id="secAdd">구분 추가 (Ctrl+F3)</button>
      <button class="smallbtn" id="secDel">구분 삭제</button>
    </div>
  `;
  top.appendChild(leftBox);

  // RIGHT: var table
  const rightBox = document.createElement("div");
  rightBox.className = "rail-box";
  rightBox.innerHTML = `
    <div class="rail-title">변수표 (A, AB, A1, AB1... 최대 3자)</div>
    <div class="var-tablewrap">
      <table class="var-table">
        <thead>
          <tr>
            <th style="min-width:90px;">변수</th>
            <th style="min-width:160px;">산식</th>
            <th style="min-width:90px;">값</th>
            <th style="min-width:140px;">비고</th>
          </tr>
        </thead>
        <tbody id="varBody"></tbody>
      </table>
    </div>
  `;
  top.appendChild(rightBox);

  $view.appendChild(top);

  // build section list
  const $secList = leftBox.querySelector("#secList");
  $secList.innerHTML = "";
  tab.sections.forEach((s, idx)=>{
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "section-item" + (idx === tab.activeSection ? " active" : "");
    btn.setAttribute("data-idx", String(idx));
    btn.innerHTML = `
      <div style="display:flex; gap:8px; align-items:center; min-width:0;">
        <div style="font-weight:900; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width: 220px;">
          ${escapeHtml(s.name || `구분 ${idx+1}`)}
        </div>
        <div class="meta">개소: ${escapeHtml((s.count ?? "").toString())}</div>
      </div>
      <div class="meta">선택</div>
    `;
    btn.onclick = ()=> setActiveSection(tabKey, idx);
    btn.addEventListener("keydown", (e)=>{
      if(e.key === "ArrowUp"){
        e.preventDefault();
        setActiveSection(tabKey, Math.max(0, tab.activeSection - 1));
      }else if(e.key === "ArrowDown"){
        e.preventDefault();
        setActiveSection(tabKey, Math.min(tab.sections.length-1, tab.activeSection + 1));
      }else if(e.key === "Enter"){
        e.preventDefault();
        setActiveSection(tabKey, idx);
      }
    });
    $secList.appendChild(btn);
  });

  // vars render
  const $varBody = rightBox.querySelector("#varBody");
  $varBody.innerHTML = "";
  sec.vars.forEach((v, i)=>{
    const tr = document.createElement("tr");
    // ⚠️ 변수표 입력들은 .cell 을 사용하지 않음(강제 이동/방향키 캡쳐 방지)
    tr.innerHTML = `
      <td><input class="varcell" data-var="key" data-i="${i}" value="${escapeAttr(v.key)}" placeholder="A / AB / A1"></td>
      <td><input class="varcell" data-var="expr" data-i="${i}" value="${escapeAttr(v.expr)}" placeholder="예: 0.5+0.5"></td>
      <td><input class="varcell readonly" data-var="val" value="${escapeAttr(String(v.val ?? 0))}" readonly></td>
      <td><input class="varcell" data-var="note" data-i="${i}" value="${escapeAttr(v.note)}" placeholder="비고"></td>
    `;
    $varBody.appendChild(tr);
  });

  // var events
  rightBox.querySelectorAll("input.varcell").forEach(inp=>{
    const kind = inp.getAttribute("data-var");
    const i = Number(inp.getAttribute("data-i") || -1);

    const commit = ()=>{
      if(i < 0) return;
      const row = sec.vars[i];
      if(!row) return;

      if(kind === "key"){
        row.key = inp.value.toString().trim().toUpperCase();
      }else if(kind === "expr"){
        row.expr = inp.value.toString();
      }else if(kind === "note"){
        row.note = inp.value.toString();
      }

      // 계산 갱신
      recalcTab(tabKey);
      saveState();

      // 값만 즉시 반영(전체 리렌더 없이)
      // (현재 섹션 vars의 val은 computeVars에서 업데이트 됨)
      // 표의 readonly 값 갱신
      const valInputs = rightBox.querySelectorAll('input[data-var="val"]');
      valInputs.forEach((vInp, idx)=>{
        vInp.value = String(roundUp3(sec.vars[idx]?.val ?? 0));
      });

      // 산출표의 readonly도 즉시 갱신을 위해 현재 포커스 행이면 refresh
      const el = document.activeElement;
      if(el && el.classList.contains("cell")) refreshReadonlyInSameRow(el);
    };

    inp.addEventListener("input", commit);
    inp.addEventListener("change", commit);
    inp.addEventListener("blur", commit);

    // 산식 Enter -> 값 계산
    inp.addEventListener("keydown", (e)=>{
      if(kind === "expr" && e.key === "Enter"){
        e.preventDefault();
        commit();
      }
    });
  });

  // section editor events
  const $secName = leftBox.querySelector("#secName");
  const $secCount = leftBox.querySelector("#secCount");
  const $secSave = leftBox.querySelector("#secSave");
  const $secAdd = leftBox.querySelector("#secAdd");
  const $secDel = leftBox.querySelector("#secDel");

  const saveSectionMeta = ()=>{
    sec.name = ($secName?.value ?? "").toString();
    sec.count = ($secCount?.value ?? "").toString();
    saveState();
    go(tabKey);
  };

  $secSave.onclick = saveSectionMeta;
  $secName.addEventListener("keydown", (e)=>{ if(e.key==="Enter"){ e.preventDefault(); saveSectionMeta(); } });
  $secCount.addEventListener("keydown", (e)=>{ if(e.key==="Enter"){ e.preventDefault(); saveSectionMeta(); } });

  $secAdd.onclick = ()=> addSection(tabKey);
  $secDel.onclick = ()=> deleteActiveSection(tabKey);

  // 리스트에서 현재 선택 항목에 포커스
  setTimeout(()=>{
    const cur = $secList.querySelector(".section-item.active");
    cur?.focus?.();
  }, 0);
}

/* ===== calc sheet ===== */
function renderCalcSheet(tabKey, title, mode){
  $view.innerHTML = "";

  // 상단: 구분/변수
  renderTopSplit(tabKey);

  // 하단: 산출표
  const sec = getActiveSection(tabKey);
  const desc = '구분(↑/↓) 이동 → 해당 구분의 변수/산출표로 전환 | 산출식 Enter 계산 | Ctrl+. 코드선택 | Ctrl+F3 구분추가 | Shift+Ctrl+F3 행추가';
  const {wrap, header} = panel(title, desc);

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addRowBtn = document.createElement("button");
  addRowBtn.className="smallbtn";
  addRowBtn.textContent="행 추가 (Shift+Ctrl+F3)";
  addRowBtn.onclick=()=>addRowToActiveSection(tabKey, 1);

  const add10Btn = document.createElement("button");
  add10Btn.className="smallbtn";
  add10Btn.textContent="+10행";
  add10Btn.onclick=()=>addRowToActiveSection(tabKey, 10);

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
          <th style="min-width:520px;">산출식</th>
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

  const tbody = tableWrap.querySelector("tbody");

  // 현재 섹션 varMap로 계산
  const varMap = computeVars(sec);
  sec.rows.forEach((r, idx)=>{
    recalcRow(r, varMap);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${inputCell(r.code, v=>{ r.code=v; }, "코드 입력", {tabId:tabKey, rowIdx:idx, colIdx:0})}</td>
      <td>${readonlyCell(r.name)}</td>
      <td>${readonlyCell(r.spec)}</td>
      <td>${readonlyCell(r.unit)}</td>

      <td>${inputCell(r.formulaExpr, v=>{ r.formulaExpr=v; }, '예: (A+0.5)*2  ( <...> 는 주석 )', {tabId:tabKey, rowIdx:idx, colIdx:1})}</td>
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
    dup.onclick=()=>{ sec.rows.splice(idx+1,0, JSON.parse(JSON.stringify(r))); saveState(); go(tabKey); };

    const del = document.createElement("button");
    del.className="smallbtn"; del.textContent="삭제";
    del.onclick=()=>{ sec.rows.splice(idx,1); saveState(); go(tabKey); };

    act.appendChild(dup);
    act.appendChild(del);
    tdAct.appendChild(act);

    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);

  const sumBox = document.createElement("div");
  sumBox.style.marginTop="10px";
  const sumVal = (mode==="steel")
    ? sec.rows.reduce((a,b)=> a + roundUp3(b.finalQty), 0)
    : sec.rows.reduce((a,b)=> a + num(b.value), 0);
  sumBox.innerHTML = `<span class="badge">현재 구분 합계: ${roundUp3(sumVal)}</span>`;
  wrap.appendChild(sumBox);

  $view.appendChild(wrap);

  wireCells();
  wireFocusTracking();
  wireMouseFocus();

  const last = lastFocusCell[tabKey] ?? {row:0,col:0};
  setTimeout(()=>focusGrid(tabKey, last.row, last.col), 0);
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

function renderSteelTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(철골+부자재의 모든 구분의 할증후수량 합산)");
  recalcAll();

  const allRows = [];
  for(const sec of state.tabs.steel.sections) allRows.push(...sec.rows);
  for(const sec of state.tabs.steelSub.sections) allRows.push(...sec.rows);

  const grouped = groupSum(allRows, r => roundUp3(r.finalQty));

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
  const {wrap} = panel("동바리_집계(Support_Total quantity)", "코드별 합계(동바리의 모든 구분의 물량(Value) 합계)");
  recalcAll();

  const allRows = [];
  for(const sec of state.tabs.support.sections) allRows.push(...sec.rows);

  const grouped = groupSum(allRows, r => num(r.value));

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
  else if(id==="steel") renderCalcSheet("steel", "철골(Steel)", "steel");
  else if(id==="steelSub") renderCalcSheet("steelSub", "철골_부자재(Processing and assembly)", "steel");
  else if(id==="support") renderCalcSheet("support", "동바리(support)", "support");
  else if(id==="steelTotal") renderSteelTotal();
  else if(id==="supportTotal") renderSupportTotal();
}

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

function insertCodesBelow(tabKey, focusRow, codeList){
  const sec = getActiveSection(tabKey);
  if(!sec) return;

  const rows = sec.rows;
  const idx = Math.min(Math.max(0, Number(focusRow) || 0), rows.length);
  const insertAt = Math.min(idx + 1, rows.length);

  const newRows = codeList.map(code=>{
    const r = makeEmptyCalcRow();
    r.code = code;
    return r;
  });

  rows.splice(insertAt, 0, ...newRows);
  recalcTab(tabKey);
  saveState();
  go(tabKey);
  setTimeout(()=>focusGrid(tabKey, insertAt, 0), 0);
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
   HOTKEYS + 방향키 이동 (calc/codes만)
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
  // ✅ 산출표/코드표의 .cell 만 grid로 간주
  return !!(el && el.classList && el.classList.contains("cell") && el.getAttribute && el.getAttribute("data-grid") === "1");
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

  const sec = getActiveSection(activeTabId);
  if(!sec) return;
  if(r >= sec.rows.length) return;

  sec.rows.splice(r, 1);
  saveState();
  go(activeTabId);

  const newRow = Math.min(r, sec.rows.length - 1);
  if(newRow >= 0) setTimeout(()=>focusGrid(activeTabId, newRow, 0), 0);
}

document.addEventListener("keydown", (e)=>{
  // Ctrl+. (picker)
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "." || e.code === "Period")){
    e.preventDefault();
    openPickerWindow();
    return;
  }

  // Ctrl+Delete (row delete)
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "Delete" || e.code === "Delete")){
    // ✅ 변수표/구분 입력에서는 동작하지 않게
    const ae = document.activeElement;
    if(ae && (ae.classList?.contains("varcell") || ae.id?.startsWith?.("sec"))) return;
    e.preventDefault();
    deleteRowAtActiveFocus();
    return;
  }

  // Ctrl+F3 = 구분 추가(산출탭) / 코드탭에서는 행추가
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3"){
    // ✅ Shift 없으면 "구분 추가" 우선
    if(!e.shiftKey){
      e.preventDefault();
      if(activeTabId === "codes"){
        const {row, col} = lastFocusCell.codes ?? {row:0,col:0};
        const insertAt = Math.min((row ?? 0) + 1, state.codes.length);
        state.codes.splice(insertAt, 0, makeEmptyCodeRow());
        saveState();
        go("codes");
        setTimeout(()=>focusGrid("codes", insertAt, col ?? 0), 0);
      }else if(["steel","steelSub","support"].includes(activeTabId)){
        addSection(activeTabId);
      }
      return;
    }

    // Shift+Ctrl+F3 = 행 추가 (현재 구분의 산출표)
    e.preventDefault();
    if(["steel","steelSub","support"].includes(activeTabId)){
      addRowToActiveSection(activeTabId, 1);
    }
    return;
  }

  // grid nav only when active element is .cell (calc/codes)
  const el = document.activeElement;
  if(!isGridEl(el)) return;

  // F2 edit
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

  // textarea caret rules
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

  // arrow move
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

/* ===== 상단 버튼들 ===== */
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
    // 보정 + 재계산
    state = loadStateFromObject(state);
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

/* ===== import 보정 ===== */
function loadStateFromObject(obj){
  try{
    const raw = JSON.stringify(obj);
    localStorage.setItem("__TMP__", raw);
    const out = loadState();
    localStorage.removeItem("__TMP__");
    // 위 loadState는 STORAGE_KEY 기반이므로, 임시를 못 씀 -> 직접 보정
  }catch{}
  // 직접 보정:
  const s = obj && typeof obj === "object" ? obj : makeState();
  if(!Array.isArray(s.codes)) s.codes = SEED_CODES;
  if(!s.tabs) s.tabs = {};
  for(const k of ["steel","steelSub","support"]){
    if(!s.tabs[k]) s.tabs[k] = makeTabState();
    if(!Array.isArray(s.tabs[k].sections) || s.tabs[k].sections.length===0){
      s.tabs[k].sections = [ makeSection("구분 1","") ];
    }
    if(typeof s.tabs[k].activeSection !== "number") s.tabs[k].activeSection = 0;
    s.tabs[k].sections.forEach(sec=>{
      if(!Array.isArray(sec.vars)) sec.vars = Array.from({length: 12}, makeEmptyVarRow);
      if(!Array.isArray(sec.rows)) sec.rows = Array.from({length: 20}, makeEmptyCalcRow);
      if(typeof sec.name !== "string") sec.name = "구분";
      if(typeof sec.count !== "string") sec.count = "";
    });
  }
  return s;
}

/* ===== init ===== */
renderTabs();
go(activeTabId);
