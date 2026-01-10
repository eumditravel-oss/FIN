/* =========================
   FIN 산출자료(Web) picker.js (확장본)
   - ✅ Ctrl+. 열면 전체 코드 즉시 표시
   - ✅ 코드 선택/다중선택/삽입 유지
   - ✅ 코드 마스터 편집(추가/삭제/엑셀업로드/JSON내보내기)
   - ✅ "코드저장/반영" 버튼으로 부모창(app) state.codes 업데이트
   ========================= */

let originTab = "steel";
let focusRow = 0;

// ✅ opener에서 받은 원본 codes
let codes = [];

// ✅ 편집용 working copy (picker 안에서만 수정)
let codesDraft = [];

// PICK
let results = [];
let cursorIndex = -1;
const selected = new Set();
let shiftSelecting = false;

// EDIT dirty tracking
let dirtyCount = 0;

// DOM
const $q = document.getElementById("q");
const $mode = document.getElementById("searchMode");
const $tbody = document.getElementById("tbody");
const $status = document.getElementById("status");
const $pickInfo = document.getElementById("pickInfo");
const $originInfo = document.getElementById("originInfo");

const $btnInsert = document.getElementById("btnInsert");
const $btnClose = document.getElementById("btnClose");
const $btnApplyCodes = document.getElementById("btnApplyCodes");

const $tabBtns = document.querySelectorAll(".tabbtn");
const $viewPick = document.getElementById("viewPick");
const $viewEdit = document.getElementById("viewEdit");

// edit dom
const $editBody = document.getElementById("editBody");
const $editStatus = document.getElementById("editStatus");
const $editInfo = document.getElementById("editInfo");
const $btnAddRow = document.getElementById("btnAddRow");
const $fileXlsx = document.getElementById("fileXlsx");
const $btnExportCodes = document.getElementById("btnExportCodes");

function esc(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;");
}
function normalize(s){ return (s ?? "").toString().toLowerCase(); }

function setStatus(t){ if($status) $status.textContent = t; }
function updateBadges(){
  if($pickInfo) $pickInfo.textContent = `선택 ${selected.size}개`;
  if($originInfo) $originInfo.textContent = `대상: ${originTab} · 기준행: ${Number(focusRow)+1}`;
}

function matchItem(item, mode, q){
  const qq = normalize(q);
  if(!qq) return true;
  const c  = normalize(item.code);
  const n  = normalize(item.name);
  const sp = normalize(item.spec);

  if(mode === "code") return c.includes(qq);
  if(mode === "name") return n.includes(qq);
  if(mode === "spec") return sp.includes(qq);
  return (n + " " + sp).includes(qq); // name_spec
}

function ensureVisible(){
  if(cursorIndex < 0) return;
  const row = $tbody?.children?.[cursorIndex];
  if(!row) return;
  row.scrollIntoView({block:"nearest"});
}

function renderPick(){
  if(!$tbody) return;
  $tbody.innerHTML = "";

  results.forEach((it, i)=>{
    const tr = document.createElement("tr");
    if(i === cursorIndex) tr.classList.add("cursor");
    if(selected.has(it.code)) tr.classList.add("sel");

    tr.innerHTML = `
      <td>${esc(it.code)}</td>
      <td>${esc(it.name)}</td>
      <td>${esc(it.spec)}</td>
      <td>${esc(it.unit)}</td>
      <td>${esc(it.surcharge)}</td>
      <td>${esc(it.conv_unit)}</td>
      <td>${esc(it.conv_factor)}</td>
    `;

    tr.addEventListener("click", ()=>{
      cursorIndex = i;
      renderPick();
      ensureVisible();
    });

    tr.addEventListener("dblclick", ()=>{
      cursorIndex = i;
      toggleSelectCursor();
      ensureVisible();
    });

    $tbody.appendChild(tr);
  });

  setStatus(`결과 ${results.length}건 · 커서 ${cursorIndex>=0 ? cursorIndex+1 : "-"}`);
  updateBadges();
}

function runSearch(){
  const q = $q?.value ?? "";
  const mode = $mode?.value ?? "name_spec";
  results = (codesDraft || []).filter(it => matchItem(it, mode, q));

  if(results.length === 0){
    cursorIndex = -1;
  }else{
    cursorIndex = Math.min(Math.max(cursorIndex, 0), results.length - 1);
    if(cursorIndex < 0) cursorIndex = 0;
  }

  renderPick();
  ensureVisible();
}

function moveCursorNoRender(delta){
  if(!results.length) return;
  cursorIndex = Math.min(results.length - 1, Math.max(0, cursorIndex + delta));
}
function moveCursor(delta){
  moveCursorNoRender(delta);
  renderPick();
  ensureVisible();
}

function toggleSelectCursor(){
  if(cursorIndex < 0) return;
  const it = results[cursorIndex];
  if(!it) return;

  if(selected.has(it.code)) selected.delete(it.code);
  else selected.add(it.code);

  renderPick();
}

function insertToParent(){
  let selectedCodes = Array.from(selected);

  if(selectedCodes.length === 0 && cursorIndex >= 0 && results[cursorIndex]){
    selectedCodes = [results[cursorIndex].code];
  }

  if(selectedCodes.length === 0){
    alert("삽입할 항목이 없습니다.");
    return;
  }

  window.opener?.postMessage({
    type: "INSERT_SELECTED",
    originTab,
    focusRow,
    selectedCodes
  }, window.location.origin);
}

function closeMe(){
  try{
    window.opener?.postMessage({ type:"CLOSE_PICKER" }, window.location.origin);
  }catch{}
  window.close();
}

/* ===== Tab switch ===== */
let activeView = "pick"; // pick | edit

function setActiveView(v){
  activeView = v;

  $tabBtns.forEach(btn=>{
    btn.classList.toggle("active", btn.getAttribute("data-tab") === v);
  });

  if($viewPick) $viewPick.style.display = (v==="pick" ? "" : "none");
  if($viewEdit) $viewEdit.style.display = (v==="edit" ? "" : "none");

  // pick view focus
  if(v==="pick"){
    setTimeout(()=> $q?.focus(), 0);
  }
}

/* ===== EDIT (Code master) ===== */
function bumpDirty(){
  dirtyCount++;
  if($editInfo) $editInfo.textContent = `변경사항: ${dirtyCount}`;
}

function normalizeRow(r){
  return {
    code: (r.code ?? "").toString().trim(),
    name: (r.name ?? "").toString().trim(),
    spec: (r.spec ?? "").toString().trim(),
    unit: (r.unit ?? "").toString().trim(),
    surcharge: (r.surcharge ?? "").toString().trim(),
    conv_unit: (r.conv_unit ?? "").toString().trim(),
    conv_factor: (r.conv_factor ?? "").toString().trim(),
    note: (r.note ?? "").toString().trim(),
  };
}

function makeEmptyCodeRow(){
  return {code:"", name:"", spec:"", unit:"", surcharge:"", conv_unit:"", conv_factor:"", note:""};
}

function renderEdit(){
  if(!$editBody) return;
  $editBody.innerHTML = "";

  codesDraft.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>

      <td><input class="cell mid" data-k="code" data-i="${idx}" value="${esc(r.code)}" /></td>
      <td><input class="cell wide" data-k="name" data-i="${idx}" value="${esc(r.name)}" /></td>
      <td><input class="cell wide" data-k="spec" data-i="${idx}" value="${esc(r.spec)}" /></td>
      <td><input class="cell small" data-k="unit" data-i="${idx}" value="${esc(r.unit)}" /></td>

      <td><input class="cell small" data-k="surcharge" data-i="${idx}" value="${esc(r.surcharge)}" placeholder="예: 7" /></td>
      <td><input class="cell small" data-k="conv_unit" data-i="${idx}" value="${esc(r.conv_unit)}" /></td>
      <td><input class="cell small" data-k="conv_factor" data-i="${idx}" value="${esc(r.conv_factor)}" /></td>

      <td><textarea class="cell note" data-k="note" data-i="${idx}">${esc(r.note)}</textarea></td>

      <td>
        <div class="row-actions">
          <button class="danger" data-act="del" data-i="${idx}">삭제</button>
        </div>
      </td>
    `;
    $editBody.appendChild(tr);
  });

  if($editStatus) $editStatus.textContent = `코드 ${codesDraft.length}개`;
}

function wireEditEvents(){
  // input/textarea change delegation
  $editBody?.addEventListener("input", (e)=>{
    const t = e.target;
    if(!(t instanceof HTMLElement)) return;

    const i = Number(t.getAttribute("data-i"));
    const k = t.getAttribute("data-k");
    if(!Number.isFinite(i) || i < 0) return;
    if(!k) return;
    if(!codesDraft[i]) return;

    if(t.tagName === "TEXTAREA"){
      codesDraft[i][k] = t.value;
    }else{
      codesDraft[i][k] = (t.value ?? "").toString();
    }

    bumpDirty();
    // pick 쪽 결과도 즉시 반영되도록 (검색표시가 codesDraft 기반)
    runSearch();
  });

  // delete delegation
  $editBody?.addEventListener("click", (e)=>{
    const t = e.target;
    if(!(t instanceof HTMLElement)) return;

    const act = t.getAttribute("data-act");
    if(act !== "del") return;

    const i = Number(t.getAttribute("data-i"));
    if(!Number.isFinite(i) || i < 0) return;

    const ok = confirm("이 코드를 삭제할까요?");
    if(!ok) return;

    codesDraft.splice(i, 1);
    bumpDirty();
    renderEdit();
    runSearch();
  });
}

function addRow(){
  codesDraft.push(makeEmptyCodeRow());
  bumpDirty();
  renderEdit();
  runSearch();
}

async function importXlsx(file){
  if(!file) return;
  if(!window.XLSX){
    alert("엑셀 업로드 라이브러리(XLSX)가 로드되지 않았습니다.");
    return;
  }

  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {defval:""});

    const mapRow = (r) => normalizeRow({
      code: r["코드"] ?? r["code"],
      name: r["품명"] ?? r["name"],
      spec: r["규격"] ?? r["spec"],
      unit: r["단위"] ?? r["unit"],
      surcharge: r["할증"] ?? r["surcharge"],
      conv_unit: r["환산단위"] ?? r["conv_unit"],
      conv_factor: r["환산계수"] ?? r["conv_factor"],
      note: r["비고"] ?? r["note"],
    });

    const mapped = rows.map(mapRow).filter(x => x.code);
    if(mapped.length === 0){
      alert("엑셀에서 유효한 '코드' 행을 찾지 못했습니다.\n헤더(코드/품명/규격/단위/할증/환산단위/환산계수/비고)를 확인해 주세요.");
      return;
    }

    const ok = confirm(`엑셀에서 ${mapped.length}개 코드를 불러옵니다.\n현재 편집중인 코드 마스터를 엑셀 값으로 덮어쓸까요?`);
    if(!ok) return;

    codesDraft = mapped;
    bumpDirty();
    renderEdit();
    runSearch();
  }catch(err){
    console.error(err);
    alert("엑셀 업로드 처리 중 오류가 발생했습니다.\n콘솔 로그를 확인해 주세요.");
  }
}

function exportCodesJson(){
  const blob = new Blob([JSON.stringify(codesDraft, null, 2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `FIN_CODES_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

/* ✅ 부모창(app.js)로 코드 마스터 반영 */
function applyCodesToOpener(){
  // 기본 정리(문자 trim)
  const cleaned = codesDraft.map(normalizeRow);

  // code 공란 제거
  const finalCodes = cleaned.filter(x => x.code);

  if(finalCodes.length === 0){
    alert("반영할 코드가 없습니다. (코드 컬럼이 비어있습니다)");
    return;
  }

  // code 중복 체크
  const seen = new Set();
  const dup = [];
  for(const r of finalCodes){
    const key = r.code.trim();
    if(seen.has(key)) dup.push(key);
    else seen.add(key);
  }
  if(dup.length){
    alert(`중복된 코드가 있습니다:\n${dup.slice(0,20).join(", ")}${dup.length>20 ? "..." : ""}\n중복을 제거한 후 다시 반영해 주세요.`);
    return;
  }

  // opener에 반영 요청
  try{
    window.opener?.postMessage({
      type: "UPDATE_CODES",
      codes: finalCodes
    }, window.location.origin);

    dirtyCount = 0;
    if($editInfo) $editInfo.textContent = `변경사항: ${dirtyCount}`;
    alert("부모창에 코드 마스터 반영 요청을 보냈습니다.\n(부모창에서 즉시 저장/재계산됩니다)");
  }catch(e){
    alert("부모창으로 반영 요청 실패(팝업/보안 설정 확인).");
  }
}

/* ===== Default search mode ===== */
function ensureModeDefault(){
  const want = "name_spec";
  const has = Array.from($mode?.options ?? []).some(o => o.value === want);
  if(has) $mode.value = want;
}

/* ===== INIT from opener ===== */
window.addEventListener("message", (event) => {
  if (event.origin !== window.location.origin) return;
  const msg = event.data;
  if (!msg || typeof msg !== "object") return;

  if (msg.type === "INIT") {
    originTab = msg.originTab || "steel";
    focusRow = Number(msg.focusRow || 0);
    codes = Array.isArray(msg.codes) ? msg.codes : [];

    // draft = deep copy
    codesDraft = JSON.parse(JSON.stringify(codes || []));

    ensureModeDefault();

    // pick init
    if($q) $q.value = "";
    selected.clear();
    shiftSelecting = false;
    cursorIndex = (codesDraft.length ? 0 : -1);

    // render both
    runSearch();
    renderEdit();
  }
});

/* ===== Keys ===== */
document.addEventListener("keydown", (e)=>{
  // 공통: Esc 닫기
  if(e.key === "Escape"){
    e.preventDefault();
    closeMe();
    return;
  }

  // PICK 전용 단축키
  if(activeView === "pick"){
    // Enter: 검색 실행(검색창에서만)
    if(e.key === "Enter" && !e.ctrlKey){
      if(document.activeElement === $q){
        e.preventDefault();
        runSearch();
        return;
      }
    }

    // Shift + ArrowDown/Up : 누적선택 + 이동행도 선택
    if(e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "ArrowDown" || e.key === "ArrowUp")){
      e.preventDefault();

      if(!shiftSelecting){
        shiftSelecting = true;
        if(cursorIndex >= 0 && results[cursorIndex]){
          selected.add(results[cursorIndex].code);
        }
      }

      moveCursorNoRender(e.key === "ArrowDown" ? 1 : -1);

      if(cursorIndex >= 0 && results[cursorIndex]){
        selected.add(results[cursorIndex].code);
      }

      renderPick();
      ensureVisible();
      return;
    }

    // ArrowDown/Up: 커서 이동
    if(e.key === "ArrowDown"){
      e.preventDefault();
      shiftSelecting = false;
      moveCursor(1);
      return;
    }
    if(e.key === "ArrowUp"){
      e.preventDefault();
      shiftSelecting = false;
      moveCursor(-1);
      return;
    }

    // Ctrl+B: 다중선택 토글
    if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "b" || e.key === "B")){
      e.preventDefault();
      toggleSelectCursor();
      return;
    }

    // Ctrl+Enter: 삽입
    if(e.ctrlKey && !e.altKey && !e.metaKey && e.key === "Enter"){
      e.preventDefault();
      insertToParent();
      return;
    }
  }

  // EDIT 전용 단축키
  if(activeView === "edit"){
    // Ctrl+S: 코드 반영
    if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "s" || e.key === "S")){
      e.preventDefault();
      applyCodesToOpener();
      return;
    }
  }

  // 탭 전환 (Alt+1 / Alt+2)
  if(e.altKey && !e.ctrlKey && !e.metaKey){
    if(e.key === "1"){ e.preventDefault(); setActiveView("pick"); return; }
    if(e.key === "2"){ e.preventDefault(); setActiveView("edit"); return; }
  }
});

/* Shift 떼면 선택 시작 상태 초기화 */
document.addEventListener("keyup", (e)=>{
  if(e.key === "Shift") shiftSelecting = false;
});

/* ===== UI events ===== */
$btnInsert?.addEventListener("click", insertToParent);
$btnClose?.addEventListener("click", closeMe);
$btnApplyCodes?.addEventListener("click", applyCodesToOpener);

$tabBtns.forEach(btn=>{
  btn.addEventListener("click", ()=>{
    const t = btn.getAttribute("data-tab");
    if(t === "pick" || t === "edit") setActiveView(t);
  });
});

// pick change triggers
$q?.addEventListener("change", runSearch);
$mode?.addEventListener("change", runSearch);

// edit events
$btnAddRow?.addEventListener("click", addRow);

$fileXlsx?.addEventListener("change", async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  await importXlsx(f);
  e.target.value = "";
});

$btnExportCodes?.addEventListener("click", exportCodesJson);

wireEditEvents();

/* ===== boot ===== */
(function boot(){
  ensureModeDefault();
  if($q) $q.value = "";
  results = [];
  cursorIndex = -1;
  renderPick();
  renderEdit();
  setActiveView("pick");
  setTimeout(()=> $q?.focus(), 0);
})();
