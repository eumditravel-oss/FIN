/* =========================
   FIN 산출자료(Web) picker.js (수정 통합본)
   - ✅ Ctrl+. 열면 검색 없이 전체 코드 즉시 표시
   - ✅ 드롭다운 기본 검색모드 = "품명+규격(name_spec)"
   - ✅ INIT 변수 중복/꼬임 제거 (originTab/focusRow/codes를 실제로 사용)
   ========================= */

let originTab = "steel";
let focusRow = 0;
let codes = [];

let results = [];
let cursorIndex = -1;
const selected = new Set(); // code string
let shiftSelecting = false; // Shift로 범위선택 시작 여부


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

function setStatus(t){ $status.textContent = t; }
function updateBadges(){
  $pickInfo.textContent = `선택 ${selected.size}개`;
  $originInfo.textContent = `대상: ${originTab} · 기준행: ${Number(focusRow)+1}`;
}

function normalize(s){ return (s ?? "").toString().toLowerCase(); }

function matchItem(item, mode, q){
  const qq = normalize(q);
  if(!qq) return true; // ✅ 검색어 없으면 전부 매칭(= 전체 표시)

  const c = normalize(item.code);
  const n = normalize(item.name);
  const sp = normalize(item.spec);

  if(mode === "code") return c.includes(qq);
  if(mode === "name") return n.includes(qq);
  if(mode === "spec") return sp.includes(qq);

  // name_spec (기본)
  return (n + " " + sp).includes(qq);
}

function render(){
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

function ensureVisible(){
  if(cursorIndex < 0) return;
  const row = $tbody.children[cursorIndex];
  if(!row) return;
  row.scrollIntoView({block:"nearest"});
}

function runSearch(){
  const q = $q.value ?? "";
  const mode = $mode.value;
  results = codes.filter(it => matchItem(it, mode, q));
  cursorIndex = results.length ? 0 : -1;
  render();
  ensureVisible();
}

function moveCursor(delta){
  if(!results.length) return;
  cursorIndex = Math.min(results.length-1, Math.max(0, cursorIndex + delta));
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
  // 선택이 없으면 “커서 항목 1개”라도 넣어주기(사용성)
  let selectedCodes = Array.from(selected);
  if(selectedCodes.length === 0 && cursorIndex >= 0){
    selectedCodes = [results[cursorIndex].code];
  }
  if(selectedCodes.length === 0){
    alert("삽입할 항목이 없습니다.");
    return;
  }

  // 메인창으로 전송
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

/* ✅ 드롭박스 기본값/순서 보정 */
function ensureModeDefault(){
  // 기본값은 name_spec(품명+규격)
  const want = "name_spec";

  // 옵션이 존재하면 value만 설정
  const has = Array.from($mode?.options ?? []).some(o => o.value === want);
  if(has){
    $mode.value = want;

    // ✅ 가능하면 name_spec 옵션을 맨 앞으로 이동(“드롭박스 첫번째” 요구 대응)
    // (HTML을 안 바꿔도 되게 JS에서 재배치)
    const opts = Array.from($mode.options);
    const idx = opts.findIndex(o => o.value === want);
    if(idx > 0){
      const opt = opts[idx];
      $mode.remove(idx);
      $mode.insertBefore(opt, $mode.firstChild);
    }
  }
}

/* INIT from opener */
window.addEventListener("message", (event) => {
  if (event.origin !== window.location.origin) return;
  const msg = event.data;
  if (!msg || typeof msg !== "object") return;

  if (msg.type === "INIT") {
    // ✅ 실제로 쓰는 변수에 세팅
    originTab = msg.originTab || "steel";
    focusRow = Number(msg.focusRow || 0);
    codes = Array.isArray(msg.codes) ? msg.codes : [];

    // 기본 검색모드 설정
    ensureModeDefault();

    // ✅ 검색어 없어도 전체 바로 표시
    $q.value = "";
    selected.clear();
    runSearch();
  }
});

/* Keys */
document.addEventListener("keydown", (e)=>{
  // Enter: 검색 실행(입력창 focus일 때만)
  if(e.key === "Enter" && !e.ctrlKey){
    if(document.activeElement === $q){
      e.preventDefault();
      runSearch();
    }
  }

       // ✅ Shift + ArrowDown/Up : 시작행부터 선택 + 이동한 행도 선택
  if(e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "ArrowDown" || e.key === "ArrowUp")){
    e.preventDefault();

    // ✅ Shift 선택이 "처음 시작"될 때: 현재 커서 행을 먼저 선택(시작행 포함)
    if(!shiftSelecting){
      shiftSelecting = true;
      if(cursorIndex >= 0){
        const cur = results[cursorIndex];
        if(cur && !selected.has(cur.code)){
          selected.add(cur.code);
        }
      }
    }

    // 커서 이동
    moveCursor(e.key === "ArrowDown" ? 1 : -1);

    // 이동한 커서 행도 선택(토글 X, 무조건 add)
    if(cursorIndex >= 0){
      const it = results[cursorIndex];
      if(it) selected.add(it.code);
    }

    render();
    ensureVisible();
    return;
  }


  // 방향키: 커서 이동(리스트 있을 때)
  if(e.key === "ArrowDown"){
    e.preventDefault();
    moveCursor(1);
  }

   
  if(e.key === "ArrowUp"){
    e.preventDefault();
    moveCursor(-1);
  }

  // Ctrl+B: 다중선택 토글
  if(e.ctrlKey && !e.altKey && !e.metaKey && (e.key === "b" || e.key === "B")){
    e.preventDefault();
    toggleSelectCursor();
  }

  // Ctrl+Enter: 삽입
  if(e.ctrlKey && !e.altKey && !e.metaKey && e.key === "Enter"){
    e.preventDefault();
    insertToParent();
  }

  // Esc: 닫기
  if(e.key === "Escape"){
    e.preventDefault();
    closeMe();
  }
});

$btnInsert.addEventListener("click", insertToParent);
$btnClose.addEventListener("click", closeMe);

// UX: 열리면 검색창 포커스
setTimeout(()=> $q.focus(), 0);
