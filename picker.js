/* =========================
   FIN 산출자료(Web) picker.js (수정 통합본 - 적용 안정판)
   - ✅ Ctrl+. 열면 검색 없이 전체 코드 즉시 표시 (INIT 수신 즉시)
   - ✅ 드롭다운 기본 검색모드 = "품명+규격(name_spec)"
   - ✅ Shift+↑/↓ = 시작행 포함 누적 다중선택(엑셀 느낌)
   - ✅ 렌더 중복/타이밍 꼬임 제거(Shift 이동은 NoRender 이동)
   ========================= */

let originTab = "steel";
let focusRow = 0;
let codes = [];

let results = [];
let cursorIndex = -1;
const selected = new Set();
let shiftSelecting = false; // Shift 선택 진행중 여부

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

function setStatus(t){ if($status) $status.textContent = t; }
function updateBadges(){
  if($pickInfo) $pickInfo.textContent = `선택 ${selected.size}개`;
  if($originInfo) $originInfo.textContent = `대상: ${originTab} · 기준행: ${Number(focusRow)+1}`;
}

function normalize(s){ return (s ?? "").toString().toLowerCase(); }

function matchItem(item, mode, q){
  const qq = normalize(q);
  if(!qq) return true; // ✅ 검색어 없으면 전부 표시

  const c  = normalize(item.code);
  const n  = normalize(item.name);
  const sp = normalize(item.spec);

  if(mode === "code") return c.includes(qq);
  if(mode === "name") return n.includes(qq);
  if(mode === "spec") return sp.includes(qq);

  // name_spec
  return (n + " " + sp).includes(qq);
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
      <td>${esc(it.note)}</td>
    `;

    tr.addEventListener("click", ()=>{
      cursorIndex = i;
      render();
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

  results = (codes || []).filter(it => matchItem(it, mode, q));

  // 커서 유지: 검색 후에도 기존 커서 인덱스가 범위를 벗어나면 보정
  if(results.length === 0){
    cursorIndex = -1;
  }else{
    cursorIndex = Math.min(Math.max(cursorIndex, 0), results.length - 1);
    if(cursorIndex < 0) cursorIndex = 0;
  }

  render();
  ensureVisible();
}

function moveCursorNoRender(delta){
  if(!results.length) return;
  cursorIndex = Math.min(results.length - 1, Math.max(0, cursorIndex + delta));
}

function moveCursor(delta){
  moveCursorNoRender(delta);
  render();
  ensureVisible();
}

function toggleSelectCursor(){
  if(cursorIndex < 0) return;
  const it = results[cursorIndex];
  if(!it) return;

  if(selected.has(it.code)) selected.delete(it.code);
  else selected.add(it.code);

  render();
}

function insertToParent(){
  let selectedCodes = Array.from(selected);

  // 선택 없으면 커서 1개라도 넣기
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

/* ✅ 드롭박스 기본값만 세팅 (옵션 이동 X: 꼬임 방지) */
function ensureModeDefault(){
  const want = "name_spec";
  const has = Array.from($mode?.options ?? []).some(o => o.value === want);
  if(has) $mode.value = want;
}

/* INIT from opener */
window.addEventListener("message", (event) => {
  if (event.origin !== window.location.origin) return;
  const msg = event.data;
  if (!msg || typeof msg !== "object") return;

  if (msg.type === "INIT") {
    originTab = msg.originTab || "steel";
    focusRow = Number(msg.focusRow || 0);
    codes = Array.isArray(msg.codes) ? msg.codes : [];

    ensureModeDefault();

    // ✅ 열리자마자 전체 표시
    if($q) $q.value = "";
    selected.clear();
    shiftSelecting = false;

    // ✅ 커서 초기화
    cursorIndex = (codes.length ? 0 : -1);

    runSearch();
  }
});

/* Keys */
document.addEventListener("keydown", (e)=>{
  // Enter: 검색 실행(검색창에서만)
  if(e.key === "Enter" && !e.ctrlKey){
    if(document.activeElement === $q){
      e.preventDefault();
      runSearch();
      return;
    }
  }

  // ✅ Shift + ArrowDown/Up : 시작행 포함 누적선택 + 이동행도 선택
  if(e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "ArrowDown" || e.key === "ArrowUp")){
    e.preventDefault();

    // Shift 시작 순간: 현재 행을 먼저 선택(시작행 포함)
    if(!shiftSelecting){
      shiftSelecting = true;
      if(cursorIndex >= 0 && results[cursorIndex]){
        selected.add(results[cursorIndex].code);
      }
    }

    // 이동(렌더 없이)
    moveCursorNoRender(e.key === "ArrowDown" ? 1 : -1);

    // 이동한 행도 선택(누적)
    if(cursorIndex >= 0 && results[cursorIndex]){
      selected.add(results[cursorIndex].code);
    }

    render();
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

  // Esc: 닫기
  if(e.key === "Escape"){
    e.preventDefault();
    closeMe();
    return;
  }
});

/* Shift 떼면 선택 시작 상태 초기화 */
document.addEventListener("keyup", (e)=>{
  if(e.key === "Shift") shiftSelecting = false;
});

$btnInsert?.addEventListener("click", insertToParent);
$btnClose?.addEventListener("click", closeMe);

/* 최초 로드 시: UI만 초기화(코드는 INIT에서 들어옴) */
(function boot(){
  ensureModeDefault();
  if($q) $q.value = "";
  results = [];
  cursorIndex = -1;
  render();
  setTimeout(()=> $q?.focus(), 0);
})();
