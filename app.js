const STORAGE_KEY = "FIN_WEB_V8";

/* ===== Seed ===== */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
];

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

function makeState(){
  return {
    codes: SEED_CODES,
    steel: Array.from({length: 20}, makeEmptyCalcRow),
    steelSub: Array.from({length: 20}, makeEmptyCalcRow),
    support: Array.from({length: 20}, makeEmptyCalcRow),
  };
}

function loadState(){
  const raw = localStorage.getItem(STORAGE_KEY);
  if(!raw) return makeState();
  try{
    const s = JSON.parse(raw);
    s.codes = Array.isArray(s.codes) ? s.codes : SEED_CODES;
    s.steel = Array.isArray(s.steel) ? s.steel : [];
    s.steelSub = Array.isArray(s.steelSub) ? s.steelSub : [];
    s.support = Array.isArray(s.support) ? s.support : [];
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
  const s = (exprRaw ?? "").toString().trim();
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
function surchargeToMul(p){ const x = num(p); return x ? (1 + x/100) : ""; }

function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}

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
  state.steel.forEach(recalcRow);
  state.steelSub.forEach(recalcRow);
  state.support.forEach(recalcRow);
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

/* ✅ 현재 포커스 셀(탭/행/열) 기억 */
const lastFocusCell = {
  codes: { row: 0, col: 0 },
  steel: { row: 0, col: 0 },
  steelSub: { row: 0, col: 0 },
  support: { row: 0, col: 0 },
};

/* ===== DOM ===== */
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

/**
 * gridAttrs = {tabId, rowIdx, colIdx}
 */
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

/* ✅ 핵심: 이벤트 바인딩(Enter 계산 포함) */
function wireCells(){
  document.querySelectorAll("[data-cell]").forEach(el=>{
    const id = el.getAttribute("data-cell");
    const meta = cellRegistry.find(x=>x.id===id);
    if(!meta) return;

    const handler = (evt)=>{
      meta.onChange(el.value);
      recalcAll();
      saveState();

      // input 중엔 전체 재렌더 금지(커서 튐 방지)
      if(evt && evt.type === "input") return;

      // blur/change/enter일 때만 재렌더
      go(activeTabId, { silentTabRender:true });
    };

    el.addEventListener("input", handler);
    el.addEventListener("blur", handler);
    el.addEventListener("change", handler);

    // ✅ 산출식(col=1) Enter = 계산 + 아래행 이동
    el.addEventListener("keydown", (e)=>{
      if(e.key !== "Enter") return;
      if(el.tagName.toLowerCase() === "textarea") return; // 비고는 줄바꿈 유지

      const col = Number(el.getAttribute("data-col") || -1);
      if(col !== 1) return;

      e.preventDefault();
      handler({ type:"enter" });
      moveGridFrom(el, +1, 0);
    });
  });

  cellRegistry.length = 0;
}

/* ===== Focus tracking for Ctrl+F3 insertion ===== */
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

/* ===== Excel-like navigation helpers ===== */
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

  // col 유동 대비 좌우 탐색
  for(let offset=1; offset<=6; offset++){
    if(focusGrid(tab, targetRow, targetCol - offset)) return true;
    if(focusGrid(tab, targetRow, targetCol + offset)) return true;
  }
  return false;
}

/* ===== Row insert below focused row (Ctrl+F3) ===== */
function getRowsByTab(tab){
  if(tab==="steel") return state.steel;
  if(tab==="steelSub") return state.steelSub;
  if(tab==="support") return state.support;
  return null;
}

function insertRowBelowActive(){
  if(!["codes","steel","steelSub","support"].includes(activeTabId)) return;

  const {row, col} = lastFocusCell[activeTabId] ?? {row:0, col:0};

  if(activeTabId === "codes"){
    const insertAt = Math.min(row + 1, state.codes.length);
    state.codes.splice(insertAt, 0, makeEmptyCodeRow());
    saveState();
    go("codes");
    setTimeout(()=>focusGrid("codes", insertAt, col), 0);
    return;
  }

  const rows = getRowsByTab(activeTabId);
  if(!rows) return;

  const insertAt = Math.min(row + 1, rows.length);
  rows.splice(insertAt, 0, makeEmptyCalcRow());
  saveState();
  go(activeTabId);
  setTimeout(()=>focusGrid(activeTabId, insertAt, col), 0);
}

/* ===== Views ===== */
function renderCodes(){
  $view.innerHTML = "";
  const {wrap, header} = panel('코드(Ctrl+".")', "코드 마스터(수정/추가). 엑셀 업로드(.xlsx)로 한 번에 등록 가능. (행 추가: Ctrl+F3)");

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가 (Ctrl+F3)";
  addBtn.onclick = ()=> insertRowBelowActive();

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

  const {row, col} = lastFocusCell.codes;
  setTimeout(()=>focusGrid("codes", row, col), 0);
}

function renderCalcSheet(title, rows, tabId, mode){
  $view.innerHTML = "";
  const desc = '산출식 입력 → 물량(Value) 자동 계산(Enter). 코드 선택 새 창: Ctrl+.  | 행 추가: Ctrl+F3';
  const {wrap, header} = panel(title, desc);

  const right = document.createElement("div");
  right.style.display="flex"; right.style.gap="8px"; right.style.flexWrap="wrap";

  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가 (Ctrl+F3)";
  addBtn.onclick=()=>insertRowBelowActive();

  const add10Btn = document.createElement("button");
  add10Btn.className="smallbtn";
  add10Btn.textContent="+10행";
  add10Btn.onclick=()=>{ for(let i=0;i<10;i++) rows.push(makeEmptyCalcRow()); saveState(); go(tabId); };

  right.appendChild(addBtn);
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
          <th style="min-width:220px;">산출식</th>
          <th style="min-width:150px;">물량(Value)</th>
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

  // ✅ calc sheet에서도 반드시 wiring
  wireCells();
  wireFocusTracking();

  const last = lastFocusCell[tabId] ?? {row:0,col:0};
  setTimeout(()=>focusGrid(tabId, last.row, last.col), 0);
}

function groupSum(rows, valueSelector){
  const map = new Map();
  for(const r of rows){
    const code = (r.code ?? "").toString().trim();
    if(!code) continue;
    const m = findCode(code);
    const cur = map.get(code) ?? { code, name:m?.name ?? r.name ?? "", spec:m?.spec ?? r.spec ?? "", unit:m?.unit ?? r.unit ?? "", sum:0 };
    cur.sum += valueSelector(r);
    map.set(code, cur);
  }
  return Array.from(map.values()).sort((a,b)=>a.code.localeCompare(b.code));
}

function renderSteelTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(철골+부자재의 할증후수량 합산)");
  recalcAll();
  const grouped = groupSum([...state.steel, ...state.steelSub], r => roundUp3(r.finalQty));

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
  const grouped = groupSum(state.support, r => num(r.value));

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

function go(id, opts={silentTabRender:false}){
  activeTabId = id;
  recalcAll();
  saveState();
  if(!opts.silentTabRender) renderTabs();

  if(id==="codes") renderCodes();
  else if(id==="steel") renderCalcSheet("철골(Steel)", state.steel, "steel", "steel");
  else if(id==="steelSub") renderCalcSheet("철골_부자재(Processing and assembly)", state.steelSub, "steelSub", "steel");
  else if(id==="support") renderCalcSheet("동바리(support)", state.support, "support", "support");
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
  const rows = getRowsByTab(tab);
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
  }

  if(msg.type === "CLOSE_PICKER"){
    try{ pickerWin?.close(); }catch{}
    pickerWin = null;
  }
});

/* ===== Hotkeys ===== */
document.addEventListener("keydown", (e)=>{
  // Ctrl+. : open picker
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "." || e.code === "Period")){
    e.preventDefault();
    openPickerWindow();
    return;
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

    const rows = getRowsByTab(activeTabId);
    if(!rows || rows.length === 0) return;
    if(r >= rows.length) return;

    rows.splice(r, 1);
    saveState();
    go(activeTabId);

    const newRow = Math.min(r, rows.length - 1);
    if(newRow >= 0) setTimeout(()=>focusGrid(activeTabId, newRow, 0), 0);
  }

  // Ctrl + Delete : delete current row (with confirm)
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "Delete" || e.code === "Delete")){
    e.preventDefault();
    deleteRowAtActiveFocus();
    return;
  }

  // Ctrl+F3 : insert row below focused row
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.code === "F3"){
    e.preventDefault();
    insertRowBelowActive();
    return;
  }
}, false);

document.getElementById("btnOpenPicker")?.addEventListener("click", openPickerWindow);

/* ===== Export/Import/Reset ===== */
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
    recalcAll();
    saveState();
    go("steel");
  }catch{
    alert("JSON 파싱 실패: 파일 내용을 확인해 주세요.");
  }finally{
    e.target.value = "";
  }
});

/* ===== GLOBAL Arrow navigation (Excel + F2 Edit Mode) ===== */
function isGridEl(el){
  return el && el.getAttribute && el.getAttribute("data-grid") === "1";
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
  const el = document.activeElement;
  if(!isGridEl(el)) return;

  const tag = (el.tagName || "").toLowerCase();
  const isTextarea = tag === "textarea";

  /* ===== F2 : 편집모드 ON ===== */
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


  /* ===== Enter / Esc : 편집모드 종료 ===== */
if(editMode && (e.key === "Enter" || e.key === "Escape")){
  e.preventDefault();
  editMode = false;
  setEditingClass(false);

  if(e.key === "Enter"){
    moveGridFrom(el, +1, 0);
  }
  return;
}

  


  /* ===== textarea는 기존 동작 유지 ===== */
  if(isTextarea){
    if(e.key === "ArrowUp"){
      if(!textareaAtTop(el)) return;
      e.preventDefault();
      moveGridFrom(el, -1, 0);
    }
    if(e.key === "ArrowDown"){
      if(!textareaAtBottom(el)) return;
      e.preventDefault();
      moveGridFrom(el, +1, 0);
    }
    return;
  }

  /* ===== input 처리 ===== */
  if(editMode){
    // 편집모드면 input 안에서 커서 이동 허용
    return;
  }

  // 🔥 기본 모드 = 무조건 셀 이동
  if(e.key === "ArrowUp"){ e.preventDefault(); setEditingClass(false); moveGridFrom(el, -1, 0); }
  if(e.key === "ArrowDown"){ e.preventDefault(); setEditingClass(false); moveGridFrom(el, -1, 0); }
  if(e.key === "ArrowLeft"){ e.preventDefault(); setEditingClass(false); moveGridFrom(el, -1, 0); }
  if(e.key === "ArrowRight"){ e.preventDefault(); setEditingClass(false); moveGridFrom(el, -1, 0); }

}, false);



let editMode = false; // F2 편집모드 여부

function setEditingClass(on){
  document.querySelectorAll('.cell.editing').forEach(x=>x.classList.remove('editing'));
  if(on){
    const el = document.activeElement;
    if(el && el.classList && el.classList.contains("cell")) el.classList.add("editing");
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
