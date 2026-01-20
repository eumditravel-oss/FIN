/* =========================
   FIN picker.js (코드선택 전용)
   - ✅ 코드 검색/선택/블록선택/삽입만 유지
   - ✅ "코드편집(마스터)" 기능/탭 완전 제거
   - ✅ Ctrl+B: 커서행 선택 토글
   - ✅ Shift+↑/↓: 연속 범위 블록 선택
   - ✅ Ctrl+Enter: 선택(또는 커서1개) 삽입 + 닫기
   - ✅ Ctrl+. 로 열린 창(INIT 메시지 수신) 그대로 사용
   - 🛠 (PATCH) INIT 수신 전 boot에서 runSearch 금지(커서/스크롤 리셋으로 ↓가 원점으로 튀는 현상 방지)
   - 🛠 (PATCH) INIT 완료 플래그(__inited)로 중복/선행 키 입력 방지
   ========================= */

let originTab = "steel";
let focusRow = 0;

// opener에서 받은 codes
let codes = [];

// search results
let results = [];
let cursorIndex = -1;
const selected = new Set(); // code string set

// Shift 블록 선택용 앵커
let rangeAnchor = null;

// ✅ INIT 완료 플래그(boot에서 runSearch 막기 + INIT 전 키입력 방지)
let __inited = false;

// DOM
const $q = document.getElementById("q");
const $mode = document.getElementById("searchMode");
const $tbody = document.getElementById("tbody");
const $status = document.getElementById("status");
const $pickInfo = document.getElementById("pickInfo");
const $originInfo = document.getElementById("originInfo");

const $btnInsert = document.getElementById("btnInsert");
const $btnClose = document.getElementById("btnClose");

function esc(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;");
}
function normalize(s){ return (s ?? "").toString().toLowerCase(); }

function updateBadges(){
  if($pickInfo) $pickInfo.textContent = `선택 ${selected.size}개`;
  if($originInfo) $originInfo.textContent = `대상: ${originTab} · 기준행: ${Number(focusRow)+1}`;
}
function setStatus(t){ if($status) $status.textContent = t; }

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

function render(){
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

    // 클릭: 커서 이동 / Shift+클릭이면 블록 선택
    tr.addEventListener("click", (e)=>{
      if(!results.length) return;

      if(e.shiftKey){
        if(rangeAnchor === null) rangeAnchor = (cursorIndex >= 0 ? cursorIndex : i);
        cursorIndex = i;
        applyRangeSelection(rangeAnchor, cursorIndex);
      }else{
        cursorIndex = i;
        rangeAnchor = null;
      }

      render();
      ensureVisible();
    });

    // 더블클릭: 토글 선택(Ctrl+B 역할)
    tr.addEventListener("dblclick", ()=>{
      cursorIndex = i;
      rangeAnchor = null;
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

  results = (codes || []).filter(it => matchItem(it, mode, q));

  if(results.length === 0){
    cursorIndex = -1;
  }else{
    cursorIndex = Math.min(Math.max(cursorIndex, 0), results.length - 1);
    if(cursorIndex < 0) cursorIndex = 0;
  }

  rangeAnchor = null;
  render();
  ensureVisible();
}

/* ===== Cursor ===== */
function moveCursorNoRender(delta){
  if(!results.length) return;
  cursorIndex = Math.min(results.length - 1, Math.max(0, cursorIndex + delta));
}
function moveCursor(delta){
  moveCursorNoRender(delta);
  render();
  ensureVisible();
}

/* ===== Selection ===== */
function toggleSelectCursor(){
  if(cursorIndex < 0) return;
  const it = results[cursorIndex];
  if(!it) return;

  if(selected.has(it.code)) selected.delete(it.code);
  else selected.add(it.code);

  render();
}

function applyRangeSelection(a, b){
  if(!results.length) return;

  const start = Math.min(a, b);
  const end   = Math.max(a, b);

  selected.clear();
  for(let i = start; i <= end; i++){
    const it = results[i];
    if(it?.code) selected.add(it.code);
  }
}

/* ===== Insert ===== */
function getSelectedCodesOrdered(){
  const ordered = [];
  for(const it of results){
    if(selected.has(it.code)) ordered.push(it.code);
  }
  return ordered;
}

function insertToParent({ closeAfter=false } = {}){
  let selectedCodes = getSelectedCodesOrdered();

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

  if(closeAfter) closeMe();
}

function closeMe(){
  try{
    window.opener?.postMessage({ type:"CLOSE_PICKER" }, window.location.origin);
  }catch{}
  window.close();
}

/* ===== INIT from opener ===== */
window.addEventListener("message", (event) => {
  if (event.origin !== window.location.origin) return;
  const msg = event.data;
  if (!msg || typeof msg !== "object") return;

  if (msg.type === "INIT") {
  // ✅ 중복 INIT 차단 (메인이 여러 번 보내도 첫 1회만 처리)
  if (__inited) {
    // (선택) 메인 쪽에서 재전송 끊게 ACK 보내고 싶으면 사용
    // try { window.opener?.postMessage({ type: "PICKER_INIT_ACK" }, window.location.origin); } catch {}
    return;
  }
  __inited = true;

  originTab = msg.originTab || "steel";
  focusRow = Number(msg.focusRow || 0);

  // codes는 배열 오브젝트(코드마스터) 그대로 들어온다고 가정
  codes = Array.isArray(msg.codes) ? msg.codes : [];

  // 초기화(여기서만 1회)
  if($q) $q.value = "";
  selected.clear();
  rangeAnchor = null;

  // ✅ 커서 초기값 확정
  cursorIndex = (codes.length ? 0 : -1);

  runSearch();
  updateBadges();

  setTimeout(()=> $q?.focus(), 0);

  // (선택) 메인 쪽에서 재전송 끊게 ACK 보내고 싶으면 사용
  // try { window.opener?.postMessage({ type: "PICKER_INIT_ACK" }, window.location.origin); } catch {}
}

});

/* ===== Keys ===== */
document.addEventListener("keydown", (e)=>{
  // ✅ INIT 전에는 조작 금지(초기 흔들림/커서 리셋 체감 방지)
  if(!__inited) {
    // 단, Esc는 닫기 허용
    if(e.key === "Escape"){
      e.preventDefault();
      closeMe();
    }
    return;
  }

  // Esc 닫기
  if(e.key === "Escape"){
    e.preventDefault();
    closeMe();
    return;
  }

  // Enter: 검색 실행(검색창에서만)
  if(e.key === "Enter" && !e.ctrlKey){
    if(document.activeElement === $q){
      e.preventDefault();
      runSearch();
      return;
    }
  }

  // Shift + ArrowDown/Up : 블록 선택
  if(e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "ArrowDown" || e.key === "ArrowUp")){
    e.preventDefault();
    if(!results.length) return;

    if(rangeAnchor === null){
      rangeAnchor = (cursorIndex >= 0 ? cursorIndex : 0);
    }

    moveCursorNoRender(e.key === "ArrowDown" ? 1 : -1);
    applyRangeSelection(rangeAnchor, cursorIndex);

    render();
    ensureVisible();
    return;
  }

  // ArrowDown/Up : 커서 이동
  if(e.key === "ArrowDown"){
    e.preventDefault();
    rangeAnchor = null;
    moveCursor(1);
    return;
  }
  if(e.key === "ArrowUp"){
    e.preventDefault();
    rangeAnchor = null;
    moveCursor(-1);
    return;
  }

  // Ctrl+B : 선택 토글
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "b" || e.key === "B")){
    e.preventDefault();
    rangeAnchor = null;
    toggleSelectCursor();
    return;
  }

  // Ctrl+Enter : 삽입 + 닫기
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.key === "Enter"){
    e.preventDefault();
    insertToParent({ closeAfter:true });
    return;
  }
});

// Shift 해제 시 앵커 해제
document.addEventListener("keyup", (e)=>{
  if(e.key === "Shift") rangeAnchor = null;
});

/* ===== UI events ===== */
$btnInsert?.addEventListener("click", ()=> insertToParent({ closeAfter:true }));
$btnClose?.addEventListener("click", closeMe);

$q?.addEventListener("change", runSearch);
$mode?.addEventListener("change", runSearch);

/* ===== boot ===== */
(function boot(){
  // ✅ INIT 수신 전에는 검색/렌더를 돌리지 않는다(커서/스크롤 리셋 원인)
  setStatus("대기중… (메인 창에서 INIT 수신)");
  updateBadges();
  setTimeout(()=> $q?.focus(), 0);
})();
