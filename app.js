/* =========================================================
   FIN 산출자료 (Web) - app.js (RESTORED)
   - Code master logic restored (Excel-like)
   - Arrow key navigation restored (boundary move)
   - Shortcuts restored (Ctrl+., Ctrl+F3, Shift+Ctrl+F3, Ctrl+B, Ctrl+Enter)
   ========================================================= */

const LS_KEY = "FIN_CALC_STATE_v10";

/* -----------------------------
   Tabs (requested names)
------------------------------ */
const TAB_KEYS = [
  { key: "code", label: "코드(Ctrl+.)" },
  { key: "steel", label: "철골" },
  { key: "steel_sum", label: "철골_집계" },
  { key: "steel_sub", label: "철골_부자재" },
  { key: "support", label: "구조이기/동바리" },
  { key: "support_sum", label: "구조이기/동바리_집계" },
];

/* top-split needed tabs */
const TOP_SPLIT_TABS = new Set(["steel", "steel_sub", "support"]);

/* -----------------------------
   Helpers
------------------------------ */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }

function safeNum(x){
  if (x === "" || x == null) return 0;
  const n = Number(x);
  return Number.isFinite(n) ? n : 0;
}

/* remove "<...>" comment blocks */
function stripAngleComments(s){
  if (!s) return "";
  return String(s).replace(/<[^>]*>/g, "");
}

/* validate variable name: starts with letter, up to 3 chars, alnum */
function isValidVarName(name){
  if (!name) return false;
  const s = String(name).trim();
  return /^[A-Za-z][A-Za-z0-9]{0,2}$/.test(s);
}

/* Build variable map from current section vars */
function varsToMap(varsArr){
  const map = {};
  for (const v of varsArr){
    const k = String(v.name || "").trim();
    if (!isValidVarName(k)) continue;
    map[k] = safeNum(v.value);
  }
  return map;
}

/* Expression evaluator (numbers + + - * / ( ) .) with variables */
function evalExpr(expr, varMap){
  const raw = stripAngleComments(expr || "").trim();
  if (!raw) return 0;

  // replace variables (word boundary)
  let s = raw;
  // sort keys longest first to avoid partial replacement
  const keys = Object.keys(varMap || {}).sort((a,b)=>b.length-a.length);
  for (const k of keys){
    const v = varMap[k];
    const re = new RegExp(`\\b${k}\\b`, "g");
    s = s.replace(re, String(v));
  }

  // allow only safe characters
  if (!/^[0-9+\-*/().\s]+$/.test(s)){
    throw new Error("INVALID_EXPR");
  }

  // eslint-disable-next-line no-new-func
  const fn = new Function(`"use strict"; return (${s});`);
  const out = fn();
  if (!Number.isFinite(out)) return 0;
  return out;
}

/* -----------------------------
   Default state
------------------------------ */
function makeEmptyCodeRow(){
  return {
    code: "",
    name: "",
    spec: "",
    unit: "",
    addPct: "",      // 할증(%)
    convUnit: "",    // 환산단위
    convFactor: "",  // 환산계수
    note: ""
  };
}

function makeEmptyCalcRow(){
  return {
    code: "",
    name: "",
    spec: "",
    unit: "",
    formula: "",
    value: 0,          // 물량
    addMul: "",        // 할증(배수) ex) 1.05
    convUnit: "",
    convFactor: "",
    convQty: 0,        // 환산수량
    finalQty: 0,       // 할증후수량
    note: ""
  };
}

function makeEmptyVarRow(){
  return { name: "", expr: "", value: 0, memo: "" };
}

function makeDefaultSection(){
  return {
    title: "구분 1",
    count: 1, // 개소
    vars: Array.from({length: 12}, ()=>makeEmptyVarRow()),
    rows: Array.from({length: 12}, ()=>makeEmptyCalcRow()),
  };
}

function defaultState(){
  return {
    activeTab: "steel",
    // code master
    codes: Array.from({length: 30}, ()=>makeEmptyCodeRow()),
    // calc tabs sections
    steel: { sections: [ makeDefaultSection() ], activeSection: 0 },
    steel_sub: { sections: [ makeDefaultSection() ], activeSection: 0 },
    support: { sections: [ makeDefaultSection() ], activeSection: 0 },
  };
}

/* -----------------------------
   Load / Save
------------------------------ */
function loadState(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return defaultState();
    const st = JSON.parse(raw);

    // minimal migrations / ensure shapes
    if (!st.codes) st.codes = Array.from({length: 30}, ()=>makeEmptyCodeRow());
    for (const k of ["steel","steel_sub","support"]){
      if (!st[k]) st[k] = { sections: [ makeDefaultSection() ], activeSection: 0 };
      if (!Array.isArray(st[k].sections) || st[k].sections.length === 0){
        st[k].sections = [ makeDefaultSection() ];
      }
      st[k].activeSection = clamp(safeNum(st[k].activeSection), 0, st[k].sections.length-1);
    }
    if (!st.activeTab) st.activeTab = "steel";
    return st;
  }catch(e){
    console.warn("loadState failed", e);
    return defaultState();
  }
}

function saveState(){
  localStorage.setItem(LS_KEY, JSON.stringify(state));
}

/* -----------------------------
   Global state
------------------------------ */
let state = loadState();

// used by code picker
let picker = {
  open: false,
  filter: "",
  selectedCodes: new Set(),
  targetInput: null, // input element where to insert code
  targetTab: null,
  targetRowIndex: null,
};

/* -----------------------------
   DOM roots
------------------------------ */
const tabsEl = $("#tabs");
const viewEl = $("#view");

/* -----------------------------
   Render Tabs
------------------------------ */
function renderTabs(){
  tabsEl.innerHTML = "";
  for (const t of TAB_KEYS){
    const btn = document.createElement("button");
    btn.className = "tab" + (state.activeTab === t.key ? " active" : "");
    btn.textContent = t.label;
    btn.addEventListener("click", ()=>{
      state.activeTab = t.key;
      saveState();
      render();
    });
    tabsEl.appendChild(btn);
  }
}

/* -----------------------------
   Code master lookup (Excel-like)
------------------------------ */
function findCodeRow(code){
  const c = String(code || "").trim();
  if (!c) return null;
  return state.codes.find(r => String(r.code || "").trim() === c) || null;
}

/* Excel mapping
   - addMul: if addPct <= 0 => "" else addPct/100 + 1
*/
function computeAddMul(addPct){
  const p = safeNum(addPct);
  if (p <= 0) return "";
  return (p/100 + 1);
}

/* -----------------------------
   Recompute a calc row using code master + formula + vars
------------------------------ */
function recomputeCalcRow(tabKey, secIdx, rowIdx){
  const tab = state[tabKey];
  if (!tab) return;
  const sec = tab.sections[secIdx];
  if (!sec) return;
  const row = sec.rows[rowIdx];
  if (!row) return;

  // pull code master
  const m = findCodeRow(row.code);
  if (m){
    row.name = m.name || "";
    row.spec = m.spec || "";
    row.unit = m.unit || "";
    row.convUnit = m.convUnit || "";
    row.convFactor = m.convFactor || "";
    const mul = computeAddMul(m.addPct);
    row.addMul = mul === "" ? "" : String(mul);
  }else{
    row.name = row.name || "";
    row.spec = row.spec || "";
    row.unit = row.unit || "";
    row.convUnit = row.convUnit || "";
    row.convFactor = row.convFactor || "";
    row.addMul = row.addMul || "";
  }

  // eval value from formula
  const varMap = varsToMap(sec.vars);
  let value = 0;
  try{
    value = evalExpr(row.formula, varMap);
  }catch(e){
    value = 0;
  }
  row.value = Number.isFinite(value) ? value : 0;

  // Excel formula:
  // convQty(M) = IF(K==0 or blank, E, E*K)
  // finalQty(N)= IF(I==0 or blank, M, M*I)
  const K = safeNum(row.convFactor);
  row.convQty = (K === 0) ? row.value : row.value * K;

  const I = safeNum(row.addMul);
  row.finalQty = (I === 0) ? row.convQty : row.convQty * I;
}

/* recompute all rows in section */
function recomputeSection(tabKey, secIdx){
  const tab = state[tabKey];
  if (!tab) return;
  const sec = tab.sections[secIdx];
  if (!sec) return;

  // recompute vars first
  const varMap = {};
  for (const v of sec.vars){
    const k = String(v.name || "").trim();
    if (!isValidVarName(k)) continue;
    // evaluate v.expr -> v.value
    let out = 0;
    try{
      out = evalExpr(v.expr, varMap); // allow earlier vars
    }catch(e){
      out = safeNum(v.value);
    }
    v.value = Number.isFinite(out) ? out : 0;
    varMap[k] = v.value;
  }

  for (let i=0;i<sec.rows.length;i++){
    recomputeCalcRow(tabKey, secIdx, i);
  }
}

/* -----------------------------
   Section operations
------------------------------ */
function addSection(tabKey){
  const tab = state[tabKey];
  if (!tab) return;
  tab.sections.push({
    title: `구분 ${tab.sections.length+1}`,
    count: 1,
    vars: Array.from({length: 12}, ()=>makeEmptyVarRow()),
    rows: Array.from({length: 12}, ()=>makeEmptyCalcRow()),
  });
  tab.activeSection = tab.sections.length - 1;
  saveState();
  render();
}

function deleteSection(tabKey){
  const tab = state[tabKey];
  if (!tab) return;
  if (tab.sections.length <= 1) return;
  tab.sections.splice(tab.activeSection, 1);
  tab.activeSection = clamp(tab.activeSection, 0, tab.sections.length-1);
  saveState();
  render();
}

/* -----------------------------
   Render: Code tab
------------------------------ */
function renderCodeTab(){
  const wrap = document.createElement("section");
  wrap.className = "panel";

  const head = document.createElement("div");
  head.className = "panel-header";
  head.innerHTML = `
    <div>
      <div class="panel-title">코드(Ctrl+.)</div>
      <div class="panel-desc">코드 선택 팝업/셀 연동은 기존 로직 연결 지점</div>
    </div>
  `;
  wrap.appendChild(head);

  const tableWrap = document.createElement("div");
  tableWrap.className = "table-wrap";

  const table = document.createElement("table");
  table.innerHTML = `
    <thead>
      <tr>
        <th style="min-width:140px;">코드</th>
        <th style="min-width:180px;">품명</th>
        <th style="min-width:180px;">규격</th>
        <th style="min-width:80px;">단위</th>
        <th style="min-width:90px;">할증(%)</th>
        <th style="min-width:110px;">환산단위</th>
        <th style="min-width:110px;">환산계수</th>
        <th style="min-width:160px;">비고</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  const tbody = $("tbody", table);

  state.codes.forEach((r, i)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="0" value="${escapeHtml(r.code)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="1" value="${escapeHtml(r.name)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="2" value="${escapeHtml(r.spec)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="3" value="${escapeHtml(r.unit)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="4" value="${escapeHtml(r.addPct)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="5" value="${escapeHtml(r.convUnit)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="6" value="${escapeHtml(r.convFactor)}" /></td>
      <td><input class="cell" data-grid="code" data-r="${i}" data-c="7" value="${escapeHtml(r.note)}" /></td>
    `;
    tbody.appendChild(tr);
  });

  tableWrap.appendChild(table);
  wrap.appendChild(tableWrap);

  viewEl.appendChild(wrap);
}

/* -----------------------------
   Render: top split (sections + vars)
------------------------------ */
function renderTopSplit(tabKey){
  const tab = state[tabKey];
  const secIdx = tab.activeSection;
  const sec = tab.sections[secIdx];

  const topSplit = document.createElement("div");
  topSplit.className = "top-split";

  const grid = document.createElement("div");
  grid.className = "calc-layout top-grid";

  // left: section list
  const left = document.createElement("div");
  left.className = "rail-box section-box";
  left.innerHTML = `
    <div class="rail-title">구분명 리스트 (↑/↓ 이동)</div>
    <div class="section-list" id="sectionList"></div>
    <div class="section-editor">
      <input id="secTitle" class="cell" placeholder="구분명" value="${escapeHtml(sec.title)}" />
      <input id="secCount" class="cell" placeholder="개소" value="${escapeHtml(sec.count)}" />
      <button id="btnSecSave" class="smallbtn">저장</button>
    </div>
    <div style="display:flex; gap:8px;">
      <button id="btnSecAdd" class="smallbtn" style="flex:1;">구분 추가 (Ctrl+F3)</button>
      <button id="btnSecDel" class="smallbtn" style="flex:1;">구분 삭제</button>
    </div>
  `;

  // fill list
  const list = $(".section-list", left);
  tab.sections.forEach((s, idx)=>{
    const item = document.createElement("div");
    item.className = "section-item" + (idx === secIdx ? " active" : "");
    item.tabIndex = 0;
    item.innerHTML = `
      <div class="name">${escapeHtml(s.title)}</div>
      <div class="meta-inline">개소: ${escapeHtml(s.count)} · <b>선택</b></div>
    `;
    item.addEventListener("click", ()=>{
      tab.activeSection = idx;
      saveState();
      render();
    });
    item.addEventListener("keydown", (e)=>{
      if (e.key === "Enter"){
        tab.activeSection = idx;
        saveState();
        render();
      }
    });
    list.appendChild(item);
  });

  // right: vars table
  const right = document.createElement("div");
  right.className = "rail-box var-box";
  right.innerHTML = `
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
        <tbody id="varBody"></tbody>
      </table>
    </div>
  `;

  const vbody = $("#varBody", right);
  sec.vars.forEach((v, rIdx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input class="cell" data-grid="vars" data-tab="${tabKey}" data-s="${secIdx}" data-r="${rIdx}" data-c="0" value="${escapeHtml(v.name)}" placeholder="예: A / AB / A1" /></td>
      <td><input class="cell" data-grid="vars" data-tab="${tabKey}" data-s="${secIdx}" data-r="${rIdx}" data-c="1" value="${escapeHtml(v.expr)}" placeholder="예: (A+0.5)*2 (<...> 주석)" /></td>
      <td><input class="cell readonly" tabindex="-1" value="${formatNum(v.value)}" readonly /></td>
      <td><input class="cell" data-grid="vars" data-tab="${tabKey}" data-s="${secIdx}" data-r="${rIdx}" data-c="3" value="${escapeHtml(v.memo)}" placeholder="비고" /></td>
    `;
    vbody.appendChild(tr);
  });

  grid.appendChild(left);
  grid.appendChild(right);
  topSplit.appendChild(grid);

  // wire section editor
  setTimeout(()=>{
    const btnSave = $("#btnSecSave", topSplit);
    const btnAdd = $("#btnSecAdd", topSplit);
    const btnDel = $("#btnSecDel", topSplit);
    const inpTitle = $("#secTitle", topSplit);
    const inpCount = $("#secCount", topSplit);

    btnSave.addEventListener("click", ()=>{
      sec.title = (inpTitle.value || "").trim() || sec.title;
      sec.count = safeNum(inpCount.value || 0) || 0;
      saveState();
      render();
    });
    btnAdd.addEventListener("click", ()=>addSection(tabKey));
    btnDel.addEventListener("click", ()=>deleteSection(tabKey));
  }, 0);

  return topSplit;
}

/* -----------------------------
   Render: calculation table (steel-like)
------------------------------ */
function renderCalcTab(tabKey, title){
  const tab = state[tabKey];
  const secIdx = tab.activeSection;
  const sec = tab.sections[secIdx];

  // top split only for needed tabs
  if (TOP_SPLIT_TABS.has(tabKey)){
    viewEl.appendChild(renderTopSplit(tabKey));
  }

  const wrap = document.createElement("section");
  wrap.className = "panel";

  const head = document.createElement("div");
  head.className = "panel-header";
  head.innerHTML = `
    <div>
      <div class="panel-title">${escapeHtml(title)}</div>
      <div class="panel-desc">
        구분(↑/↓) 이동 → 해당 구분의 변수/산출표로 전환 | 산출식 Enter 계산 |
        Ctrl+. 코드선택 | Ctrl+F3 구분추가 | Shift+Ctrl+F3 행추가
      </div>
    </div>
    <div style="display:flex; gap:8px; align-items:center;">
      <button class="smallbtn" id="btnAddRow">행 추가 (Shift+Ctrl+F3)</button>
      <button class="smallbtn" id="btnAdd10Row">+10행</button>
    </div>
  `;
  wrap.appendChild(head);

  const tableWrap = document.createElement("div");
  tableWrap.className = "table-wrap";

  const table = document.createElement("table");
  table.innerHTML = `
    <thead>
      <tr>
        <th style="min-width:60px;">No</th>
        <th style="min-width:140px;">코드</th>
        <th style="min-width:180px;">품명(자동)</th>
        <th style="min-width:180px;">규격(자동)</th>
        <th style="min-width:90px;">단위(자동)</th>
        <th style="min-width:260px;">산출식</th>
        <th style="min-width:120px;">물량(Value)</th>
        <th style="min-width:110px;">할증(배수)</th>
        <th style="min-width:120px;">환산단위</th>
        <th style="min-width:120px;">환산계수</th>
        <th style="min-width:130px;">환산수량</th>
        <th style="min-width:140px;">할증후수량</th>
        <th style="min-width:160px;">비고</th>
      </tr>
    </thead>
    <tbody id="calcBody"></tbody>
  `;
  const tbody = $("#calcBody", table);

  sec.rows.forEach((r, i)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i+1}</td>
      <td><input class="cell" data-grid="calc" data-tab="${tabKey}" data-s="${secIdx}" data-r="${i}" data-c="0" value="${escapeHtml(r.code)}" placeholder="코드 입력" /></td>
      <td><input class="cell readonly" tabindex="-1" value="${escapeHtml(r.name)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${escapeHtml(r.spec)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${escapeHtml(r.unit)}" readonly /></td>
      <td><input class="cell" data-grid="calc" data-tab="${tabKey}" data-s="${secIdx}" data-r="${i}" data-c="5" value="${escapeHtml(r.formula)}" placeholder="예: (A+0.5)*2 (<...>는 주석)" /></td>
      <td><input class="cell readonly" tabindex="-1" value="${formatNum(r.value)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${escapeHtml(r.addMul)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${escapeHtml(r.convUnit)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${escapeHtml(r.convFactor)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${formatNum(r.convQty)}" readonly /></td>
      <td><input class="cell readonly" tabindex="-1" value="${formatNum(r.finalQty)}" readonly /></td>
      <td><input class="cell" data-grid="calc" data-tab="${tabKey}" data-s="${secIdx}" data-r="${i}" data-c="12" value="${escapeHtml(r.note)}" placeholder="비고" /></td>
    `;
    tbody.appendChild(tr);
  });

  tableWrap.appendChild(table);
  wrap.appendChild(tableWrap);
  viewEl.appendChild(wrap);

  // add row buttons
  setTimeout(()=>{
    $("#btnAddRow")?.addEventListener("click", ()=>{
      sec.rows.push(makeEmptyCalcRow());
      saveState();
      render();
    });
    $("#btnAdd10Row")?.addEventListener("click", ()=>{
      for (let k=0;k<10;k++) sec.rows.push(makeEmptyCalcRow());
      saveState();
      render();
    });
  }, 0);
}

/* -----------------------------
   Render: summaries (simple)
------------------------------ */
function renderSummaryTab(tabKey, title, sumField){
  const wrap = document.createElement("section");
  wrap.className = "panel";

  const total = computeTotal(tabKey, sumField);

  wrap.innerHTML = `
    <div class="panel-header">
      <div>
        <div class="panel-title">${escapeHtml(title)}</div>
        <div class="panel-desc">합계 표시</div>
      </div>
      <div class="badge">합계: ${formatNum(total)}</div>
    </div>
  `;
  viewEl.appendChild(wrap);
}

function computeTotal(tabKey, sumField){
  const tab = state[tabKey];
  if (!tab) return 0;

  // for steel_sum: sum finalQty across steel sections
  if (tabKey === "steel_sum"){
    let s = 0;
    for (const sec of state.steel.sections){
      recomputeSection("steel", state.steel.sections.indexOf(sec));
      for (const r of sec.rows) s += safeNum(r.finalQty);
    }
    return s;
  }

  if (tabKey === "support_sum"){
    let s = 0;
    for (const sec of state.support.sections){
      recomputeSection("support", state.support.sections.indexOf(sec));
      for (const r of sec.rows) s += safeNum(r.value); // 동바리 집계는 물량 합계
    }
    return s;
  }

  return 0;
}

/* -----------------------------
   Modal: Code picker
------------------------------ */
function ensurePickerModal(){
  let bd = $("#codePickerBackdrop");
  if (bd) return bd;

  bd = document.createElement("div");
  bd.id = "codePickerBackdrop";
  bd.className = "modal-backdrop";
  bd.innerHTML = `
    <div class="modal" role="dialog" aria-modal="true">
      <div class="modal-head">
        <div class="modal-title">코드 선택</div>
        <div class="modal-tools">
          <input id="pickerSearch" class="modal-search" placeholder="검색(코드/품명/규격)..." />
          <button id="pickerInsert" class="smallbtn">삽입 (Ctrl+Enter)</button>
          <button id="pickerClose" class="smallbtn">닫기 (Esc)</button>
        </div>
      </div>
      <div class="modal-body">
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th style="min-width:64px;">선택</th>
                <th style="min-width:140px;">코드</th>
                <th style="min-width:180px;">품명</th>
                <th style="min-width:200px;">규격</th>
                <th style="min-width:90px;">단위</th>
                <th style="min-width:90px;">할증(%)</th>
                <th style="min-width:110px;">환산단위</th>
                <th style="min-width:110px;">환산계수</th>
                <th style="min-width:160px;">비고</th>
              </tr>
            </thead>
            <tbody id="pickerBody"></tbody>
          </table>
        </div>
        <div style="margin-top:10px; font-size:12px; color:rgba(0,0,0,.65);">
          · 다중선택: Ctrl+B (현재 선택 토글) · 삽입: Ctrl+Enter
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(bd);

  $("#pickerClose", bd).addEventListener("click", closePicker);
  bd.addEventListener("click", (e)=>{
    if (e.target === bd) closePicker();
  });
  $("#pickerInsert", bd).addEventListener("click", insertPickedCodes);
  $("#pickerSearch", bd).addEventListener("input", (e)=>{
    picker.filter = e.target.value || "";
    renderPickerRows();
  });

  return bd;
}

function openPicker(targetInput){
  const bd = ensurePickerModal();
  picker.open = true;
  picker.filter = "";
  picker.selectedCodes = new Set();
  picker.targetInput = targetInput;

  bd.classList.add("show");
  $("#pickerSearch", bd).value = "";
  $("#pickerSearch", bd).focus();
  renderPickerRows();
}

function closePicker(){
  const bd = $("#codePickerBackdrop");
  if (!bd) return;
  picker.open = false;
  picker.targetInput = null;
  bd.classList.remove("show");
}

function renderPickerRows(){
  const bd = $("#codePickerBackdrop");
  if (!bd) return;
  const body = $("#pickerBody", bd);
  body.innerHTML = "";

  const q = (picker.filter || "").trim().toLowerCase();

  const rows = state.codes
    .map(r => ({...r, _code: String(r.code||"").trim()}))
    .filter(r => r._code)
    .filter(r=>{
      if (!q) return true;
      return (
        String(r.code||"").toLowerCase().includes(q) ||
        String(r.name||"").toLowerCase().includes(q) ||
        String(r.spec||"").toLowerCase().includes(q)
      );
    });

  rows.forEach((r)=>{
    const tr = document.createElement("tr");
    const checked = picker.selectedCodes.has(r._code);
    tr.innerHTML = `
      <td>
        <input type="checkbox" ${checked ? "checked" : ""} />
      </td>
      <td>${escapeHtml(r._code)}</td>
      <td>${escapeHtml(r.name)}</td>
      <td>${escapeHtml(r.spec)}</td>
      <td>${escapeHtml(r.unit)}</td>
      <td>${escapeHtml(r.addPct)}</td>
      <td>${escapeHtml(r.convUnit)}</td>
      <td>${escapeHtml(r.convFactor)}</td>
      <td>${escapeHtml(r.note)}</td>
    `;

    const cb = $("input[type='checkbox']", tr);
    cb.addEventListener("change", ()=>{
      if (cb.checked) picker.selectedCodes.add(r._code);
      else picker.selectedCodes.delete(r._code);
    });

    tr.addEventListener("dblclick", ()=>{
      picker.selectedCodes = new Set([r._code]);
      insertPickedCodes();
    });

    body.appendChild(tr);
  });
}

function togglePickCurrent(){
  const bd = $("#codePickerBackdrop");
  if (!bd) return;
  // 토글: 현재 hover/첫번째 선택 행 기준이 애매해서
  // "검색결과 첫번째 코드" 토글로 단순 처리
  const first = $("#pickerBody tr td:nth-child(2)", bd);
  if (!first) return;
  const code = first.textContent.trim();
  if (!code) return;
  if (picker.selectedCodes.has(code)) picker.selectedCodes.delete(code);
  else picker.selectedCodes.add(code);
  renderPickerRows();
}

function insertPickedCodes(){
  if (!picker.targetInput) return;

  const codes = Array.from(picker.selectedCodes);
  if (codes.length === 0) return;

  // insert into current cell (calc tab code column),
  // if multiple: fill downward starting from current row
  const t = picker.targetInput;
  const grid = t.dataset.grid;
  if (grid !== "calc"){
    // for now only calc grid insertion
    t.value = codes[0];
    t.dispatchEvent(new Event("change", {bubbles:true}));
    closePicker();
    return;
  }

  const tabKey = t.dataset.tab;
  const s = Number(t.dataset.s);
  const r0 = Number(t.dataset.r);
  const tab = state[tabKey];
  const sec = tab.sections[s];

  for (let i=0;i<codes.length;i++){
    const r = r0 + i;
    while (sec.rows.length <= r) sec.rows.push(makeEmptyCalcRow());
    sec.rows[r].code = codes[i];
    // recompute for that row
    recomputeCalcRow(tabKey, s, r);
  }

  saveState();
  render();

  // focus next cell after insertion
  setTimeout(()=>{
    const next = document.querySelector(`input[data-grid="calc"][data-tab="${tabKey}"][data-s="${s}"][data-r="${r0}"][data-c="5"]`);
    next?.focus();
  }, 0);

  closePicker();
}

/* -----------------------------
   Render main
------------------------------ */
function render(){
  viewEl.innerHTML = "";
  renderTabs();

  // recompute visible tab section before render
  if (state.activeTab === "steel") recomputeSection("steel", state.steel.activeSection);
  if (state.activeTab === "steel_sub") recomputeSection("steel_sub", state.steel_sub.activeSection);
  if (state.activeTab === "support") recomputeSection("support", state.support.activeSection);

  switch(state.activeTab){
    case "code": renderCodeTab(); break;
    case "steel": renderCalcTab("steel", "철골"); break;
    case "steel_sub": renderCalcTab("steel_sub", "철골_부자재"); break;
    case "support": renderCalcTab("support", "구조이기/동바리"); break;
    case "steel_sum": renderSummaryTab("steel_sum", "철골_집계", "finalQty"); break;
    case "support_sum": renderSummaryTab("support_sum", "구조이기/동바리_집계", "value"); break;
    default:
      renderCodeTab();
  }

  wireInputs();
}

/* -----------------------------
   Input wiring (change handlers)
------------------------------ */
function wireInputs(){
  // code master inputs
  $$(".cell[data-grid='code']").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      const r = Number(inp.dataset.r);
      const c = Number(inp.dataset.c);
      const row = state.codes[r];
      if (!row) return;

      const v = inp.value;
      if (c===0) row.code = v;
      if (c===1) row.name = v;
      if (c===2) row.spec = v;
      if (c===3) row.unit = v;
      if (c===4) row.addPct = v;
      if (c===5) row.convUnit = v;
      if (c===6) row.convFactor = v;
      if (c===7) row.note = v;

      saveState();
    });
  });

  // vars inputs
  $$(".cell[data-grid='vars']").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      const tabKey = inp.dataset.tab;
      const s = Number(inp.dataset.s);
      const r = Number(inp.dataset.r);
      const c = Number(inp.dataset.c);

      const sec = state[tabKey].sections[s];
      const vrow = sec.vars[r];
      if (!vrow) return;

      if (c===0) vrow.name = inp.value;
      if (c===1) vrow.expr = inp.value;
      if (c===3) vrow.memo = inp.value;

      // do not auto-jump on single 'A' (fix)
      // only recompute when leaving cell (on blur) or Enter
      saveState();
    });

    inp.addEventListener("blur", ()=>{
      const tabKey = inp.dataset.tab;
      const s = Number(inp.dataset.s);
      recomputeSection(tabKey, s);
      saveState();
      render();
    });

    inp.addEventListener("keydown", (e)=>{
      if (e.key === "Enter"){
        e.preventDefault();
        const tabKey = inp.dataset.tab;
        const s = Number(inp.dataset.s);
        recomputeSection(tabKey, s);
        saveState();
        render();
      }
    });
  });

  // calc inputs
  $$(".cell[data-grid='calc']").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      const tabKey = inp.dataset.tab;
      const s = Number(inp.dataset.s);
      const r = Number(inp.dataset.r);
      const c = Number(inp.dataset.c);

      const sec = state[tabKey].sections[s];
      const row = sec.rows[r];
      if (!row) return;

      if (c===0) row.code = inp.value;
      if (c===5) row.formula = inp.value;
      if (c===12) row.note = inp.value;

      saveState();
    });

    inp.addEventListener("blur", ()=>{
      const tabKey = inp.dataset.tab;
      const s = Number(inp.dataset.s);
      const r = Number(inp.dataset.r);
      recomputeCalcRow(tabKey, s, r);
      saveState();
      render();
    });

    inp.addEventListener("keydown", (e)=>{
      if (e.key === "Enter"){
        e.preventDefault();
        const tabKey = inp.dataset.tab;
        const s = Number(inp.dataset.s);
        const r = Number(inp.dataset.r);
        recomputeCalcRow(tabKey, s, r);
        saveState();
        render();
      }
    });
  });
}

/* -----------------------------
   Arrow key navigation (Excel-like at boundaries)
   - only move cell when caret is at boundary (leftmost/rightmost)
------------------------------ */
function getCaretInfo(input){
  try{
    return { start: input.selectionStart, end: input.selectionEnd, len: input.value.length };
  }catch{
    return { start: 0, end: 0, len: 0 };
  }
}

function focusCellByData(grid, tabKey, s, r, c){
  let sel = `input[data-grid="${grid}"][data-r="${r}"][data-c="${c}"]`;
  if (grid === "vars" || grid === "calc"){
    sel = `input[data-grid="${grid}"][data-tab="${tabKey}"][data-s="${s}"][data-r="${r}"][data-c="${c}"]`;
  }
  const el = document.querySelector(sel);
  if (el){
    el.focus();
    // place caret at end for convenience
    try{
      const L = el.value.length;
      el.setSelectionRange(L, L);
    }catch{}
    return true;
  }
  return false;
}

function handleArrowNav(e){
  const t = e.target;
  if (!(t instanceof HTMLInputElement)) return;
  if (!t.classList.contains("cell")) return;

  const grid = t.dataset.grid;
  if (!grid) return;

  const key = e.key;
  if (!["ArrowLeft","ArrowRight","ArrowUp","ArrowDown"].includes(key)) return;

  // don’t hijack readonly
  if (t.readOnly) return;

  // allow normal editing unless boundary
  const { start, end, len } = getCaretInfo(t);
  const atLeft = (start === 0 && end === 0);
  const atRight = (start === len && end === len);

  let wantMove = false;
  if (key === "ArrowLeft" && atLeft) wantMove = true;
  if (key === "ArrowRight" && atRight) wantMove = true;
  if (key === "ArrowUp") wantMove = true;
  if (key === "ArrowDown") wantMove = true;

  if (!wantMove) return;

  // if user is selecting text (range), keep native behavior
  if (start !== end && (key==="ArrowLeft" || key==="ArrowRight")) return;

  // compute next cell
  if (grid === "code"){
    const r = Number(t.dataset.r);
    const c = Number(t.dataset.c);
    const maxC = 7;
    let nr=r, nc=c;
    if (key==="ArrowLeft") nc = clamp(c-1,0,maxC);
    if (key==="ArrowRight") nc = clamp(c+1,0,maxC);
    if (key==="ArrowUp") nr = Math.max(0, r-1);
    if (key==="ArrowDown") nr = r+1;

    e.preventDefault();
    focusCellByData("code", null, null, nr, nc);
    return;
  }

  if (grid === "vars"){
    const tabKey = t.dataset.tab;
    const s = Number(t.dataset.s);
    const r = Number(t.dataset.r);
    const c = Number(t.dataset.c);
    // editable columns: 0,1,3 (2 is readonly)
    const order = [0,1,3];
    const idx = order.indexOf(c);
    if (idx < 0) return;

    let nr=r, nc=c;
    if (key==="ArrowLeft") nc = order[Math.max(0, idx-1)];
    if (key==="ArrowRight") nc = order[Math.min(order.length-1, idx+1)];
    if (key==="ArrowUp") nr = Math.max(0, r-1);
    if (key==="ArrowDown") nr = r+1;

    e.preventDefault();
    focusCellByData("vars", tabKey, s, nr, nc);
    return;
  }

  if (grid === "calc"){
    const tabKey = t.dataset.tab;
    const s = Number(t.dataset.s);
    const r = Number(t.dataset.r);
    const c = Number(t.dataset.c);
    // editable columns: code(0), formula(5), note(12)
    const order = [0,5,12];
    const idx = order.indexOf(c);
    if (idx < 0) return;

    let nr=r, nc=c;
    if (key==="ArrowLeft") nc = order[Math.max(0, idx-1)];
    if (key==="ArrowRight") nc = order[Math.min(order.length-1, idx+1)];
    if (key==="ArrowUp") nr = Math.max(0, r-1);
    if (key==="ArrowDown") nr = r+1;

    e.preventDefault();
    focusCellByData("calc", tabKey, s, nr, nc);
    return;
  }
}

/* -----------------------------
   Shortcuts
------------------------------ */
function onGlobalKeyDown(e){
  // allow ctrl combinations even inside inputs
  const isCtrl = e.ctrlKey || e.metaKey;

  // open picker
  if (isCtrl && e.key === "."){
    e.preventDefault();
    // if focused on calc code cell, use it, else open and insert single later
    const active = document.activeElement;
    if (active && active instanceof HTMLInputElement && active.dataset.grid === "calc" && Number(active.dataset.c) === 0){
      openPicker(active);
    }else{
      // open picker anyway; insert into first available code cell
      const first = document.querySelector(`input[data-grid="calc"][data-c="0"]`);
      if (first) openPicker(first);
    }
    return;
  }

  // picker shortcuts
  if (picker.open){
    if (isCtrl && e.key.toLowerCase() === "b"){
      e.preventDefault();
      togglePickCurrent();
      return;
    }
    if (isCtrl && e.key === "Enter"){
      e.preventDefault();
      insertPickedCodes();
      return;
    }
    if (e.key === "Escape"){
      e.preventDefault();
      closePicker();
      return;
    }
  }

  // Ctrl+F3 add section (only for calc tabs)
  if (isCtrl && e.key === "F3"){
    if (TOP_SPLIT_TABS.has(state.activeTab)){
      e.preventDefault();
      addSection(state.activeTab);
    }
    return;
  }

  // Shift+Ctrl+F3 add row
  if (isCtrl && e.shiftKey && e.key === "F3"){
    if (state.activeTab === "steel" || state.activeTab === "steel_sub" || state.activeTab === "support"){
      e.preventDefault();
      const tab = state[state.activeTab];
      const sec = tab.sections[tab.activeSection];
      sec.rows.push(makeEmptyCalcRow());
      saveState();
      render();
    }
    return;
  }

  // section up/down (when not typing in modal)
  if (!picker.open && (e.key === "ArrowUp" || e.key === "ArrowDown") && TOP_SPLIT_TABS.has(state.activeTab)){
    // if focused in an input, do not hijack (arrow nav handler will do boundary moves)
    const ae = document.activeElement;
    if (ae && ae instanceof HTMLInputElement) return;

    const tab = state[state.activeTab];
    const dir = (e.key === "ArrowUp") ? -1 : 1;
    const next = clamp(tab.activeSection + dir, 0, tab.sections.length-1);
    if (next !== tab.activeSection){
      e.preventDefault();
      tab.activeSection = next;
      saveState();
      render();
    }
  }
}

/* -----------------------------
   Export / Import / Reset buttons
------------------------------ */
function wireTopButtons(){
  $("#btnOpenPicker")?.addEventListener("click", ()=>{
    const active = document.activeElement;
    if (active && active instanceof HTMLInputElement && active.dataset.grid === "calc" && Number(active.dataset.c) === 0){
      openPicker(active);
    }else{
      const first = document.querySelector(`input[data-grid="calc"][data-c="0"]`);
      if (first) openPicker(first);
      else openPicker(document.createElement("input"));
    }
  });

  $("#btnExport")?.addEventListener("click", ()=>{
    const blob = new Blob([JSON.stringify(state, null, 2)], {type:"application/json"});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "FIN_state.json";
    a.click();
    URL.revokeObjectURL(a.href);
  });

  $("#fileImport")?.addEventListener("change", async (e)=>{
    const f = e.target.files?.[0];
    if (!f) return;
    const txt = await f.text();
    try{
      const imported = JSON.parse(txt);
      state = imported;
      saveState();
      render();
    }catch(err){
      alert("가져오기 실패: JSON 형식 확인");
    }finally{
      e.target.value = "";
    }
  });

  $("#btnReset")?.addEventListener("click", ()=>{
    if (!confirm("초기화할까요? (로컬저장 데이터 삭제)")) return;
    state = defaultState();
    saveState();
    render();
  });
}

/* -----------------------------
   Utils: HTML escape + number format
------------------------------ */
function escapeHtml(s){
  const x = (s ?? "").toString();
  return x
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function formatNum(n){
  const x = safeNum(n);
  // keep simple: show up to 3 decimals
  const s = (Math.round(x*1000)/1000).toString();
  return s;
}

/* -----------------------------
   Boot
------------------------------ */
document.addEventListener("keydown", onGlobalKeyDown, true);
document.addEventListener("keydown", handleArrowNav, true);

wireTopButtons();
render();
