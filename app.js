/* =========================
   FIN 산출자료(Web) app.js (섹션=서브시트 방식) - FIXED
   - ✅ "구분(빨간영역)" = 섹션 목록 (여러 줄)
   - ✅ 섹션 선택 시, 산출표(노란영역)는 해당 섹션의 별도 rows를 표시
   - ✅ Ctrl+F3: 섹션(구분) 추가 + 그 섹션으로 즉시 이동 + 새 산출표 준비
   - ✅ Ctrl+Shift+F3: 산출표 행 추가 (기존 Ctrl+F3 대체)
   - ✅ 집계(철골_집계/동바리_집계)는 "모든 섹션"을 합산
   - ✅ Ctrl+. picker 그대로
   - ✅ 방향키 이동/F2 편집모드/마우스 td 클릭 포커스 유지
   ========================= */

const STORAGE_KEY = "FIN_WEB_V9";

/* ===== Seed ===== */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
];

let suppressRerenderOnce = false;

/* =========================
   ✅ Mouse click -> focus cell (delegation, once)
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

/* ===== Row factories ===== */
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

/* ===== New: Section model =====
   tab마다 sections[] 를 갖고, activeIndex로 현재 섹션 선택
*/
function makeDefaultSection(){
  return { label:"", count:"", rows: makeDefaultRows(20) };
}
function makeSectionsPackFromLegacyRows(rows){
  // 기존 steel 배열이 있으면, 그걸 1번 섹션으로 감싸서 마이그레이션
  return {
    activeIndex: 0,
    sections: [{
      label: "",
      count: "",
      rows: Array.isArray(rows) ? rows : makeDefaultRows(20)
    }]
  };
}

function makeState(){
  return {
    codes: SEED_CODES,

    // ✅ 섹션 구조(서브시트)
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

    // codes 보정
    s.codes = Array.isArray(s.codes) ? s.codes : SEED_CODES;

    // ✅ V8(legacy) -> V9 migrate
    // 기존: s.steel / s.steelSub / s.support + s.sectionBar
    // 신규: s.sheets.tab.sections[]
    if(!s.sheets){
      s.sheets = {};
      s.sheets.steel    = makeSectionsPackFromLegacyRows(s.steel);
      s.sheets.steelSub = makeSectionsPackFromLegacyRows(s.steelSub);
      s.sheets.support  = makeSectionsPackFromLegacyRows(s.support);
      delete s.steel; delete s.steelSub; delete s.support;
      delete s.sectionBar;
    }

    // sheets 보정
    for(const k of ["steel","steelSub","support"]){
      if(!s.sheets[k]) s.sheets[k] = { activeIndex:0, sections:[makeDefaultSection()] };
      if(!Array.isArray(s.sheets[k].sections) || s.sheets[k].sections.length === 0){
        s.sheets[k].sections = [makeDefaultSection()];
      }
      if(typeof s.sheets[k].activeIndex !== "number") s.sheets[k].activeIndex = 0;
      if(s.sheets[k].activeIndex < 0) s.sheets[k].activeIndex = 0;
      if(s.sheets[k].activeIndex >= s.sheets[k].sections.length) s.sheets[k].activeIndex = 0;

      // section rows 보정
      s.sheets[k].sections.forEach(sec=>{
        if(!sec || typeof sec !== "object") return;
        if(sec.label == null) sec.label = "";
        if(sec.count == null) sec.count = "";
        if(!Array.isArray(sec.rows)) sec.rows = makeDefaultRows(20);
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
function evalExpr(exprRaw){
  const withoutTags = (exprRaw ?? "").toString().replace(/<[^>]*>/g, "");
  const s = withoutTags.trim();
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
function surchargeToMul(p){
  const x = num(p);
  return x ? (1 + x/100) : "";
}
function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}

/* ===== Sheet/Section helpers ===== */
function getPack(tab){
  return state.sheets?.[tab] ?? null;
}
function getActiveSection(tab){
  const pack = getPack(tab);
  if(!pack) return null;
  const i = Math.min(Math.max(0, pack.activeIndex|0), pack.sections.length-1);
  return pack.sections[i] ?? null;
}
function getActiveRows(tab){
  return getActiveSection(tab)?.rows ?? null;
}
function setActiveSectionIndex(tab, idx){
  const pack = getPack(tab);
  if(!pack) return;
  const n = pack.sections.length;
  const next = Math.min(Math.max(0, idx|0), n-1);
  pack.activeIndex = next;
  saveState();
}

/* ===== Recalc ===== */
function recalcRow(row){
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
  row.value = evalExpr(row.formulaExpr);

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
      (sec.rows || []).forEach(recalcRow);
    }
  }
}

/* ✅ 현재 편집중인 행 readonly만 즉시 갱신 */
function refreshReadonlyInSameRow(activeEl){
  const tabId = activeEl?.getAttribute?.("data-tab");
  if(!tabId) return;
  if(tabId === "codes") return;

  const rows = getActiveRows(tabId);
  if(!rows) return;

  const rowIdx = Number(activeEl.getAttribute("data-row") || -1);
  if(rowIdx < 0) return;
  const r = rows[rowIdx];
  if(!r) return;

  const tr = activeEl.closest("tr");
  if(!tr) return;

  // readonly 순서: [name, spec, unit, value, surchargeMul, convUnit, convQty, finalQty]
  const ro = tr.querySelectorAll("input.cell.readonly");
  if(ro.length < 8) return;

  ro[0].value = r.name ?? "";
  ro[1].value = r.spec ?? "";
  ro[2].value = r.unit ?? "";
  ro[3].value = String(roundUp3(r.value));
  ro[4].value = (r.surchargeMul === "" ? "" : String(r.surchargeMul));
  ro[5].value = r.convUnit ?? "";
  ro[6].value = String(roundUp3(r.convQty));
  ro[7].value = String(roundUp3(r.finalQty));
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

/* ✅ 포커스 셀 기억 */
const lastFocusCell = {
  codes: { row: 0, col: 0 },
  steel: { row: 0, col: 0 },
  steelSub: { row: 0, col: 0 },
  support: { row: 0, col: 0 },
};

const $tabs = document.getElementById("tabs");
const $view = document.getElementById("view");

/* ===== HTML helpers ===== */
function escapeHtml(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;");
}
function escapeAttr(s){ return escapeHtml(s).replaceAll("\n","&#10;"); }

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

/* ===== Focus tracking ===== */
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

/* ===== Grid nav ===== */
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

/* ===== Rows manipulation ===== */
function insertCalcRowBelowActive(){
  if(!["steel","steelSub","support"].includes(activeTabId)) return;

  const rows = getActiveRows(activeTabId);
  if(!rows) return;

  const {row, col} = lastFocusCell[activeTabId] ?? {row:0, col:0};
  const insertAt = Math.min((row|0) + 1, rows.length);

  rows.splice(insertAt, 0, makeEmptyCalcRow());
  saveState();
  go(activeTabId);
  setTimeout(()=>focusGrid(activeTabId, insertAt, col), 0);
}

function addSection(tab){
  const pack = getPack(tab);
  if(!pack) return;

  const curIdx = pack.activeIndex|0;
  const insertAt = Math.min(curIdx + 1, pack.sections.length);

  pack.sections.splice(insertAt, 0, makeDefaultSection());
  pack.activeIndex = insertAt;

  saveState();
}

/* ===== Views ===== */
function renderCodes(){
  $view.innerHTML = "";
  const {wrap, header} = panel('코드(Ctrl+".")', "코드 마스터(수정/추가). 엑셀 업로드(.xlsx)로 한 번에 등록 가능.");

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
    setTimeout(()=>focusGrid("codes", insertAt, col), 0);
  };

  // Excel upload
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
      recalcAll();
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
          <th style="min-width:70px;">No</th>
          <th style="min-width:170px;">코드</th>
          <th style="min-width:220px;">품명</th>
          <th style="min-width:220px;">규격</th>
          <th style="min-width:90px;">단위</th>
          <th style="min-width:90px;">할증</th>
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
      <td>${idx+1}</td>
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

/* ===== Section Bar UI ===== */
function renderSectionBar(tabId){
  const pack = getPack(tabId);
  if(!pack) return null;

  const activeIdx = pack.activeIndex|0;
  const sec = getActiveSection(tabId) || pack.sections[0];

  const bar = document.createElement("div");
  bar.className = "sectionbar";

  // 섹션 칩 리스트
  const chips = pack.sections.map((s, i)=>{
    const title = (s.label || "").trim() ? s.label.trim() : `구분 ${i+1}`;
    const count = (s.count ?? "").toString().trim();
    const sub = count !== "" ? ` · 개소:${count}` : "";
    const active = i === activeIdx ? " active" : "";
    return `<button class="secchip${active}" data-sec="${i}" title="${escapeAttr(title)}">${escapeHtml(title)}${escapeHtml(sub)}</button>`;
  }).join("");

  bar.innerHTML = `
    <div class="sectionbar-inner">
      <div class="sectionbar-left">
        <span class="sectionbar-title">구분</span>

        <div class="secchips" id="secChips">${chips}</div>

        <input class="sectionbar-input" id="secLabel" placeholder="구분명(예: 2층 바닥 철골보)" value="${escapeAttr(sec?.label ?? "")}">
        <input class="sectionbar-input small" id="secCount" placeholder="개소(예: 0,1,2...)" value="${escapeAttr(sec?.count ?? "")}">

        <button class="smallbtn" id="btnSecAdd">구분 추가 (Ctrl+F3)</button>
      </div>
      <div class="sectionbar-right">
        <span class="sectionbar-hint">구분을 클릭하면 해당 구분의 산출표로 전환됩니다.</span>
      </div>
    </div>
  `;

  // 이벤트 연결
  setTimeout(()=>{
    const $chips = document.getElementById("secChips");
    const $label = document.getElementById("secLabel");
    const $count = document.getElementById("secCount");
    const $add = document.getElementById("btnSecAdd");

    const saveCurrentMeta = ()=>{
      const cur = getActiveSection(tabId);
      if(!cur) return;
      cur.label = ($label?.value ?? "").toString();
      cur.count = ($count?.value ?? "").toString();
      saveState();
      // 칩 텍스트 갱신을 위해 재렌더
      go(tabId);
    };

    // 칩 클릭: 섹션 전환
    $chips?.addEventListener("click", (e)=>{
      const btn = e.target?.closest?.("button[data-sec]");
      if(!btn) return;
      const idx = Number(btn.getAttribute("data-sec") || 0);
      setActiveSectionIndex(tabId, idx);
      go(tabId);
    });

    // 메타 저장(입력)
    $label?.addEventListener("change", saveCurrentMeta);
    $count?.addEventListener("change", saveCurrentMeta);

    // 추가 버튼
    $add?.addEventListener("click", ()=>{
      addSection(tabId);
      go(tabId);
    });
  }, 0);

  return bar;
}

function renderCalcSheet(title, tabId, mode){
  $view.innerHTML = "";

  // ✅ 섹션바 먼저
  const bar = renderSectionBar(tabId);
  if(bar) $view.appendChild(bar);

  const rows = getActiveRows(tabId) || makeDefaultRows(20);

  const desc = '산출식 입력 → 물량(Value) 자동 계산(Enter). 코드 선택 새 창: Ctrl+.  | (구분추가: Ctrl+F3 / 행추가: Ctrl+Shift+F3)';
  const {wrap, header} = panel(title, desc);

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addRowBtn = document.createElement("button");
  addRowBtn.className="smallbtn";
  addRowBtn.textContent="행 추가 (Ctrl+Shift+F3)";
  addRowBtn.onclick=()=>insertCalcRowBelowActive();

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
          <th style="min-width:80px;">물량(Value)</th>

          <th style="min-width:220px;">비고</th>
          <th style="min-width:120px;">할증(배수)</th>
          <th style="min-width:120px;">환산단위</th>
          <th style="min-width:140px;">환산계수</th>
          <th style="min-width:140px;">환산수량</th>
          <th style="min-width:160px;">할증후수량</th>
          <th style="min-width:120px;">작업</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;

  const tbody = tableWrap.querySelector("tbody");

  rows.forEach((r, idx)=>{
    recalcRow(r);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>

      <td>${inputCell(r.code, v=>{ r.code=v; }, "코드 입력", {tabId, rowIdx:idx, colIdx:0})}</td>
      <td>${readonlyCell(r.name)}</td>
      <td>${readonlyCell(r.spec)}</td>
      <td>${readonlyCell(r.unit)}</td>

      <td>${inputCell(r.formulaExpr, v=>{ r.formulaExpr=v; }, "예: (0.5+0.3)/2", {tabId, rowIdx:idx, colIdx:1})}</td>
      <td>${readonlyCell(String(roundUp3(r.value)))}</td>

      <td>${textAreaCell(r.note, v=>{ r.note=v; }, {tabId, rowIdx:idx, colIdx:2})}</td>

      <td>${readonlyCell(r.surchargeMul === "" ? "" : String(r.surchargeMul))}</td>
      <td>${readonlyCell(r.convUnit)}</td>

      <td>${inputCell(r.convFactor, v=>{ r.convFactor=v; }, "비워도 됨", {tabId, rowIdx:idx, colIdx:3})}</td>

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

  const sumBox = document.createElement("div");
  sumBox.style.marginTop="10px";
  const sumVal = (mode==="steel")
    ? rows.reduce((a,b)=> a + roundUp3(b.finalQty), 0)
    : rows.reduce((a,b)=> a + num(b.value), 0);
  sumBox.innerHTML = `<span class="badge">합계: ${roundUp3(sumVal)}</span>`;
  wrap.appendChild(sumBox);

  $view.appendChild(wrap);

  wireCells();
  wireFocusTracking();
  wireMouseFocus();

  const last = lastFocusCell[tabId] ?? {row:0,col:0};
  setTimeout(()=>focusGrid(tabId, last.row, last.col), 0);
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
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(철골+부자재 모든 구분(섹션)의 할증후수량 합산)");
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
  const {wrap} = panel("동바리_집계(Support_Total quantity)", "코드별 합계(동바리 모든 구분(섹션)의 물량(Value) 합계)");
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

/* ===== Router ===== */
function go(id, opts={silentTabRender:false}){
  activeTabId = id;
  recalcAll();
  saveState();
  if(!opts.silentTabRender) renderTabs();

  if(id==="codes") renderCodes();
  else if(id==="steel") renderCalcSheet("철골(Steel)", "steel", "steel");
  else if(id==="steelSub") renderCalcSheet("철골_부자재(Processing and assembly)", "steelSub", "steel");
  else if(id==="support") renderCalcSheet("동바리(support)", "support", "support");
  else if(id==="steelTotal") renderSteelTotal();
  else if(id==="supportTotal") renderSupportTotal();
}

renderTabs();
go(activeTabId);

/* ===== Code Picker window (Ctrl+.) ===== */
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

/* =========================
   ✅ HOTKEYS + ✅ 방향키/편집모드 (capture)
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
  // Ctrl+. : picker
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "." || e.code === "Period")){
    e.preventDefault();
    openPickerWindow();
    return;
  }

  // ✅ Ctrl+F3 : "구분(섹션)" 추가
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3" && !e.shiftKey){
    if(["steel","steelSub","support"].includes(activeTabId)){
      e.preventDefault();
      addSection(activeTabId);
      go(activeTabId);
      return;
    }
  }

  // ✅ Ctrl+Shift+F3 : 산출행 추가
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3" && e.shiftKey){
    if(["steel","steelSub","support"].includes(activeTabId)){
      e.preventDefault();
      insertCalcRowBelowActive();
      return;
    }
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

  function beginNav(){
    suppressRerenderOnce = true;
    setTimeout(()=>suppressRerenderOnce=false, 0);
  }

  if(isTextareaEl(el)){
    if(e.key === "ArrowUp"){
      if(!textareaAtTop(el)) return;
      e.preventDefault(); beginNav();
      moveGridFrom(el, -1, 0);
      return;
    }
    if(e.key === "ArrowDown"){
      if(!textareaAtBottom(el)) return;
      e.preventDefault(); beginNav();
      moveGridFrom(el, +1, 0);
      return;
    }
    if(e.key === "ArrowLeft"){
      if(!caretAtStart(el)) return;
      e.preventDefault(); beginNav();
      moveGridFrom(el, 0, -1);
      return;
    }
    if(e.key === "ArrowRight"){
      if(!caretAtEnd(el)) return;
      e.preventDefault(); beginNav();
      moveGridFrom(el, 0, +1);
      return;
    }
    return;
  }

  if(e.key === "ArrowUp"){ e.preventDefault(); beginNav(); moveGridFrom(el, -1, 0); return; }
  if(e.key === "ArrowDown"){ e.preventDefault(); beginNav(); moveGridFrom(el, +1, 0); return; }
  if(e.key === "ArrowLeft"){ e.preventDefault(); beginNav(); moveGridFrom(el, 0, -1); return; }
  if(e.key === "ArrowRight"){ e.preventDefault(); beginNav(); moveGridFrom(el, 0, +1); return; }

}, true);

/* ===== 버튼들 ===== */
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
    // 기본 보정
    if(!state.sheets) state = loadState(); // 형식 이상 시 기본으로
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
