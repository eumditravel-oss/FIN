/* app.js (v15) - restore: code VLOOKUP + mult/conv + summary */
(() => {
  "use strict";

  /* =========================
     Utils
  ========================= */
  const $ = (sel, el = document) => el.querySelector(sel);
  const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

  function h(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === "class") node.className = v;
      else if (k === "dataset") Object.assign(node.dataset, v);
      else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
      else if (v === true) node.setAttribute(k, "");
      else if (v !== false && v != null) node.setAttribute(k, String(v));
    }
    for (const ch of children.flat()) {
      if (ch == null) continue;
      node.appendChild(typeof ch === "string" ? document.createTextNode(ch) : ch);
    }
    return node;
  }

  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));
  function rid() {
    try { return crypto.getRandomValues(new Uint32Array(2)).join("-"); }
    catch { return String(Date.now()) + "-" + Math.floor(Math.random() * 1e9); }
  }
  const toNum = (x, fallback = 0) => {
    const n = Number(String(x ?? "").trim());
    return Number.isFinite(n) ? n : fallback;
  };

  function caretInfo(input) {
    try {
      return { s: input.selectionStart ?? 0, e: input.selectionEnd ?? 0, len: (input.value ?? "").length };
    } catch {
      return { s: 0, e: 0, len: (input.value ?? "").length };
    }
  }

  /* =========================
     Tabs
  ========================= */
  const TAB_DEFS = [
    { id: "code", label: "코드(Ctrl+.)", type: "code" },
    { id: "steel", label: "철골", type: "calc_steel" },
    { id: "steel_sum", label: "철골_집계", type: "summary_steel" },
    { id: "steel_aux", label: "철골_부자재", type: "calc_aux" },
    { id: "support", label: "구조이기/동바리", type: "calc_support" },
    { id: "support_sum", label: "구조이기/동바리_집계", type: "summary_support" },
  ];
  const isCalcTab = (t) => t?.type?.startsWith("calc_");
  const isSummaryTab = (t) => t?.type?.startsWith("summary_");

  /* =========================
     Storage
  ========================= */
  const LS_KEY = "FIN_WEB_V15_STATE";
  const loadState = () => {
    try {
      const r = localStorage.getItem(LS_KEY);
      return r ? JSON.parse(r) : null;
    } catch { return null; }
  };
  const saveState = () => {
    try { localStorage.setItem(LS_KEY, JSON.stringify(state)); } catch {}
  };

  /* =========================
     Model
  ========================= */
  const makeDefaultVars = () =>
    Array.from({ length: 12 }).map(() => ({ key: "", expr: "", value: 0, note: "" }));

  const makeDefaultRows = (calcType) => {
    const rows = Array.from({ length: 12 }).map(() => ({
      code: "",
      name: "",
      spec: "",
      unit: "",

      expr: "",        // 산출식
      value: 0,        // 물량(Value)

      mult: "",        // 할증(배수) - 숫자/변수/수식 가능
      conv: "",        // 환산단위 - 숫자/변수/수식 가능

      adj: 0,          // 할증후수량 = value * mult * conv
      memo: "",        // (예비) 메모
    }));

    // 초기 탭별 기본 단위 예시(없으면 VLOOKUP 단위가 우선)
    if (calcType === "calc_steel") rows.forEach((r) => (r.unit = "M"));
    if (calcType === "calc_aux") rows.forEach((r) => (r.unit = "M2"));
    if (calcType === "calc_support") rows.forEach((r) => (r.unit = "EA"));

    return rows;
  };

  const makeSection = (name = "구분 1", count = "") => ({
    id: rid(),
    name,
    count,
    vars: makeDefaultVars(),
  });

  function makeDefaultState() {
    const tabs = {};
    for (const t of TAB_DEFS) {
      tabs[t.id] = {
        id: t.id,
        label: t.label,
        type: t.type,

        // calc-only
        sections: [],
        activeSectionId: null,
        sectionsRows: {},

        // code DB (master)
        codeLibrary: [], // [{code,name,spec,unit}]
      };

      if (t.type.startsWith("calc_")) {
        const sec = makeSection("1층 바닥 철골보", "1");
        tabs[t.id].sections = [sec];
        tabs[t.id].activeSectionId = sec.id;
        tabs[t.id].sectionsRows[sec.id] = makeDefaultRows(t.type);
      }
    }

    // code master는 code 탭에만 저장(다른 탭은 참조만)
    tabs.code.codeLibrary = [];

    return { activeTabId: "steel", tabs, ui: { editMode: false } };
  }

  let state = loadState() || makeDefaultState();

  const getActiveTab = () => state.tabs[state.activeTabId];
  const getActiveSection = (tab) => {
    if (!tab?.sections?.length) return null;
    const id = tab.activeSectionId || tab.sections[0].id;
    return tab.sections.find((s) => s.id === id) || tab.sections[0];
  };

  /* =========================
     CODE VLOOKUP (복원 핵심)
  ========================= */
  function normalizeCode(code) {
    return String(code ?? "").trim();
  }
  function getCodeLibrary() {
    return (state.tabs.code?.codeLibrary || []);
  }
  function lookupCode(code) {
    const k = normalizeCode(code);
    if (!k) return null;
    const lib = getCodeLibrary();
    return lib.find((r) => normalizeCode(r.code) === k) || null;
  }
  function applyVlookupToRow(rowObj) {
    const found = lookupCode(rowObj.code);
    if (found) {
      rowObj.name = found.name ?? "";
      rowObj.spec = found.spec ?? "";
      rowObj.unit = found.unit ?? rowObj.unit ?? "";
    } else {
      // 최초형식에서 보통 "못 찾으면 공란" 처리
      rowObj.name = "";
      rowObj.spec = "";
      // unit은 사용자가 이미 지정했으면 유지하는 경우도 있어 애매함 → 기본은 유지
      // rowObj.unit = rowObj.unit;
    }
  }

  /* =========================
     Expression eval
     - 산출식/할증/환산 모두 "변수 수식" 허용
  ========================= */
  const stripAngleComments = (s) => (s ? String(s).replace(/<[^>]*>/g, "") : "");
  const normalizeExpr = (s) => stripAngleComments(s).replace(/\s+/g, " ").trim();
  const isValidVarName = (v) => /^[A-Z][A-Z0-9]{0,2}$/.test(v || "");

  function safeEvalMath(expr, varMap) {
    const raw = normalizeExpr(expr);
    if (!raw) return 0;

    // 숫자 단독이면 빠르게 처리
    if (/^[+-]?\d+(\.\d+)?$/.test(raw)) return toNum(raw, 0);

    const tokenized = raw.replace(/[A-Z][A-Z0-9]{0,2}/g, (m) => String(Number(varMap[m] ?? 0) || 0));
    if (!/^[0-9+\-*/().\s]+$/.test(tokenized)) return 0;
    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(`"use strict"; return (${tokenized});`);
      const out = fn();
      const num = Number(out);
      return Number.isFinite(num) ? num : 0;
    } catch { return 0; }
  }

  function buildVarMap(section) {
    const map = {};
    for (const r of section.vars) if (isValidVarName(r.key)) map[r.key] = 0;
    // 4회 반복으로 의존 변수 수렴
    for (let i = 0; i < 4; i++) {
      for (const r of section.vars) {
        if (isValidVarName(r.key)) map[r.key] = safeEvalMath(r.expr, map);
      }
    }
    return map;
  }

  const formatNumber = (n) => {
    const num = Number(n);
    if (!Number.isFinite(num)) return "0";
    const x = Math.round(num * 1000000) / 1000000;
    return String(x);
  };

  /* =========================
     Sticky sizing vars
  ========================= */
  function updateStickyVars() {
    const topbar = $(".topbar");
    const tabs = $("#tabs");
    const topbarH = topbar ? Math.ceil(topbar.getBoundingClientRect().height) : 0;
    const tabsH = tabs ? Math.ceil(tabs.getBoundingClientRect().height) : 0;
    document.documentElement.style.setProperty("--topbarH", `${topbarH}px`);
    document.documentElement.style.setProperty("--tabsH", `${tabsH}px`);
  }

  /* =========================
     DOM targets
  ========================= */
  const tabsEl = $("#tabs");
  const viewEl = $("#view");

  /* =========================
     Section actions
  ========================= */
  function addSection(tab) {
    const sec = makeSection(`구분 ${tab.sections.length + 1}`, "");
    tab.sections.push(sec);
    tab.activeSectionId = sec.id;
    tab.sectionsRows[sec.id] = makeDefaultRows(tab.type);
    saveState();
    render();
    requestAnimationFrame(() => $(`.section-item[data-sec-id="${sec.id}"]`)?.focus());
  }

  function deleteActiveSection(tab) {
    if (!tab?.sections || tab.sections.length <= 1) return;
    const cur = getActiveSection(tab);
    const idx = tab.sections.findIndex((s) => s.id === cur.id);
    tab.sections.splice(idx, 1);
    delete tab.sectionsRows[cur.id];
    const next = tab.sections[clamp(idx, 0, tab.sections.length - 1)];
    tab.activeSectionId = next.id;
    saveState();
    render();
  }

  /* =========================
     Excel-like edit mode (F2)
  ========================= */
  function setEditMode(on) {
    state.ui.editMode = !!on;
    document.body.classList.toggle("editmode", !!on);
    const f = document.activeElement;
    if (f instanceof HTMLInputElement || f instanceof HTMLTextAreaElement) {
      if (on) f.classList.add("editing");
      else f.classList.remove("editing");
    }
  }
  function toggleEditMode() { setEditMode(!state.ui.editMode); saveState(); }

  /* =========================
     Keyboard nav: VAR table
  ========================= */
  const VAR_COLS = ["key", "expr", "note"];

  function focusVarCell(row, col) {
    const box = $("#varBox");
    if (!box) return;
    const q = `input[data-scope="var"][data-row="${row}"][data-col="${col}"]`;
    const target = box.querySelector(q);
    if (target) {
      target.focus();
      if (!state.ui.editMode) target.select?.();
    }
  }
  function moveVarCell(row, col, dRow, dCol, maxRow) {
    const colIdx = VAR_COLS.indexOfs(VAR_COLS, col);
  }

  // safer helper
  function varColIndex(col) { return VAR_COLS.indexOf(col); }

  function moveVarCell2(row, col, dRow, dCol) {
    const tab = getActiveTab();
    const sec = getActiveSection(tab);
    if (!sec) return;
    const maxRow = sec.vars.length - 1;
    const colIdx = varColIndex(col);
    const nextRow = clamp(row + dRow, 0, maxRow);
    const nextCol = VAR_COLS[clamp(colIdx + dCol, 0, VAR_COLS.length - 1)];
    focusVarCell(nextRow, nextCol);
  }

  function handleVarKeydown(e) {
    const t = e.target;
    if (!(t instanceof HTMLInputElement)) return;
    if (t.dataset.scope !== "var") return;

    const row = Number(t.dataset.row);
    const col = t.dataset.col;

    // 전역 단축키가 막히지 않도록 방향키만 여기서 강제 처리
    if (["ArrowUp","ArrowDown","ArrowLeft","ArrowRight","Enter"].includes(e.key)) {
      e.stopPropagation();
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    const c = caretInfo(t);

    // 편집모드 OFF: 방향키 = 셀 이동
    if (!state.ui.editMode) {
      if (e.key === "ArrowUp") { e.preventDefault(); moveVarCell2(row, col, -1, 0); return; }
      if (e.key === "ArrowDown") { e.preventDefault(); moveVarCell2(row, col, +1, 0); return; }
      if (e.key === "ArrowLeft") { e.preventDefault(); moveVarCell2(row, col, 0, -1); return; }
      if (e.key === "ArrowRight") { e.preventDefault(); moveVarCell2(row, col, 0, +1); return; }
      if (e.key === "Enter") {
        e.preventDefault();
        if (col === "key") focusVarCell(row, "expr");
        else if (col === "expr") { saveState(); render(); requestAnimationFrame(() => focusVarCell(row, "note")); }
        else focusVarCell(clamp(row + 1, 0, 9999), "key");
        return;
      }
      return;
    }

    // 편집모드 ON: 커서 끝/시작이면 셀 이동 허용
    if (e.key === "ArrowUp") { e.preventDefault(); moveVarCell2(row, col, -1, 0); return; }
    if (e.key === "ArrowDown") { e.preventDefault(); moveVarCell2(row, col, +1, 0); return; }
    if (e.key === "ArrowLeft") {
      if (c.s === 0 && c.e === 0 && col !== "key") { e.preventDefault(); moveVarCell2(row, col, 0, -1); }
      return;
    }
    if (e.key === "ArrowRight") {
      if (c.s === c.len && c.e === c.len && col !== "note") { e.preventDefault(); moveVarCell2(row, col, 0, +1); }
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      saveState(); render();
      requestAnimationFrame(() => focusVarCell(row, col));
      return;
    }
  }

  /* =========================
     Keyboard nav: CALC table (복원 + 개선)
     - code / expr / mult / conv
  ========================= */
  const CALC_COLS = ["code","expr","mult","conv"];

  function focusCalcCell(row, col) {
    const panel = $("#calcPanel");
    if (!panel) return;
    const q = `input[data-scope="calc"][data-row="${row}"][data-col="${col}"]`;
    const target = panel.querySelector(q);
    if (target) {
      target.focus();
      if (!state.ui.editMode) target.select?.();
    }
  }

  function moveCalcCell(row, col, dRow, dCol) {
    const tab = getActiveTab();
    const sec = getActiveSection(tab);
    if (!sec) return;
    const rows = tab.sectionsRows[sec.id] || [];
    const maxRow = rows.length - 1;
    const colIdx = CALC_COLS.indexOf(col);
    const nextRow = clamp(row + dRow, 0, maxRow);
    const nextCol = CALC_COLS[clamp(colIdx + dCol, 0, CALC_COLS.length - 1)];
    focusCalcCell(nextRow, nextCol);
  }

  function handleCalcKeydown(e) {
    const t = e.target;
    if (!(t instanceof HTMLInputElement)) return;
    if (t.dataset.scope !== "calc") return;

    const row = Number(t.dataset.row);
    const col = t.dataset.col;

    if (["ArrowUp","ArrowDown","ArrowLeft","ArrowRight","Enter"].includes(e.key)) {
      e.stopPropagation();
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    const c = caretInfo(t);

    if (!state.ui.editMode) {
      if (e.key === "ArrowUp") { e.preventDefault(); moveCalcCell(row, col, -1, 0); return; }
      if (e.key === "ArrowDown") { e.preventDefault(); moveCalcCell(row, col, +1, 0); return; }
      if (e.key === "ArrowLeft") { e.preventDefault(); moveCalcCell(row, col, 0, -1); return; }
      if (e.key === "ArrowRight") { e.preventDefault(); moveCalcCell(row, col, 0, +1); return; }
      if (e.key === "Enter") {
        e.preventDefault();
        saveState(); render();
        requestAnimationFrame(() => focusCalcCell(row, col));
        return;
      }
      return;
    }

    // 편집모드 ON
    if (e.key === "ArrowUp") { e.preventDefault(); moveCalcCell(row, col, -1, 0); return; }
    if (e.key === "ArrowDown") { e.preventDefault(); moveCalcCell(row, col, +1, 0); return; }
    if (e.key === "ArrowLeft") {
      if (c.s === 0 && c.e === 0 && col !== "code") { e.preventDefault(); moveCalcCell(row, col, 0, -1); }
      return;
    }
    if (e.key === "ArrowRight") {
      if (c.s === c.len && c.e === c.len && col !== "conv") { e.preventDefault(); moveCalcCell(row, col, 0, +1); }
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      saveState(); render();
      requestAnimationFrame(() => focusCalcCell(row, col));
      return;
    }
  }

  /* =========================
     Code picker modal (Ctrl+.)
     - 실제 삽입(현재 포커스된 code셀)까지 연결
  ========================= */
  let lastFocusedCalcCode = null; // {tabId, sectionId, rowIndex}

  function rememberFocusedCalcCode(inputEl) {
    if (!(inputEl instanceof HTMLInputElement)) return;
    if (inputEl.dataset.scope !== "calc") return;
    if (inputEl.dataset.col !== "code") return;
    const tab = getActiveTab();
    const sec = getActiveSection(tab);
    if (!tab || !sec) return;
    lastFocusedCalcCode = { tabId: tab.id, sectionId: sec.id, rowIndex: Number(inputEl.dataset.row) };
  }

  function openCodePicker() {
    const lib = getCodeLibrary();

    const overlay = h("div", { class: "modal-overlay", id: "codeModal" });
    const modal = h("div", { class: "modal" });

    const head = h("div", { class: "modal-head" },
      h("div", { class: "modal-title" }, "코드 선택 (Ctrl+B 다중선택 · Ctrl+Enter 삽입)"),
      h("button", { class: "smallbtn", type: "button", onclick: () => close() }, "닫기")
    );

    const search = h("input", {
      class: "cell",
      type: "text",
      placeholder: "검색: 코드/품명/규격",
      oninput: () => renderList(search.value),
    });

    const listWrap = h("div", { class: "table-wrap", style: "max-height:60vh;" });
    const table = h("table", {});
    const thead = h("thead", {}, h("tr", {}, h("th", {}, "선택"), h("th", {}, "코드"), h("th", {}, "품명"), h("th", {}, "규격"), h("th", {}, "단위")));
    const tbody = h("tbody");
    table.append(thead, tbody);
    listWrap.appendChild(table);

    const foot = h("div", { class: "modal-foot" },
      h("div", { style: "font-size:12px; color:rgba(0,0,0,.65);" }, `코드 DB ${lib.length}개`),
      h("button", { class: "smallbtn", type: "button", onclick: () => insertSelected() }, "삽입(Ctrl+Enter)")
    );

    modal.append(head, search, listWrap, foot);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const selected = new Set(); // key = index

    function renderList(q) {
      const keyword = (q || "").trim().toLowerCase();
      tbody.innerHTML = "";
      const rows = lib.filter((r) => {
        if (!keyword) return true;
        return (
          String(r.code || "").toLowerCase().includes(keyword) ||
          String(r.name || "").toLowerCase().includes(keyword) ||
          String(r.spec || "").toLowerCase().includes(keyword)
        );
      });

      rows.slice(0, 300).forEach((r, idx) => {
        const key = String(idx);
        const chk = h("input", {
          type: "checkbox",
          checked: selected.has(key),
          onclick: (e) => {
            e.stopPropagation();
            if (chk.checked) selected.add(key);
            else selected.delete(key);
          }
        });

        const tr = h("tr", {
          onclick: () => {
            chk.checked = !chk.checked;
            if (chk.checked) selected.add(key);
            else selected.delete(key);
          }
        },
          h("td", {}, chk),
          h("td", {}, r.code || ""),
          h("td", {}, r.name || ""),
          h("td", {}, r.spec || ""),
          h("td", {}, r.unit || "")
        );

        tbody.appendChild(tr);
      });
    }

    function insertSelected() {
      // 마지막 포커스된 calc code 셀에 삽입
      if (!lastFocusedCalcCode) {
        alert("코드 삽입 위치가 없습니다. 산출표의 '코드' 셀을 먼저 클릭하고 Ctrl+.를 실행해 주세요.");
        return;
      }
      const { tabId, sectionId, rowIndex } = lastFocusedCalcCode;
      const targetTab = state.tabs[tabId];
      if (!targetTab) return;

      const rows = targetTab.sectionsRows[sectionId];
      if (!rows) return;

      const picked = Array.from(selected).map((k) => lib[Number(k)]).filter(Boolean);
      if (!picked.length) {
        alert("선택된 코드가 없습니다.");
        return;
      }

      // 다중선택이면 아래로 연속 삽입(최초형식 느낌)
      let rIdx = rowIndex;
      for (const p of picked) {
        if (!rows[rIdx]) break;
        rows[rIdx].code = p.code || "";
        applyVlookupToRow(rows[rIdx]);
        rIdx++;
      }

      saveState();
      render();
      close();

      // 삽입 후 원래 위치로 포커스
      requestAnimationFrame(() => focusCalcCell(rowIndex, "code"));
    }

    function close() {
      overlay.remove();
      document.removeEventListener("keydown", onKey, true);
    }

    function onKey(e) {
      if (e.key === "Escape") { e.preventDefault(); close(); return; }
      if (e.ctrlKey && e.key.toLowerCase() === "b") { e.preventDefault(); return; }
      if (e.ctrlKey && e.key === "Enter") { e.preventDefault(); insertSelected(); }
    }

    document.addEventListener("keydown", onKey, true);
    renderList("");
    setTimeout(() => search.focus(), 0);
  }

  /* =========================
     Render: Tabs
  ========================= */
  function renderTabs() {
    tabsEl.innerHTML = "";
    for (const t of TAB_DEFS) {
      tabsEl.appendChild(
        h("button", {
          class: "tab" + (state.activeTabId === t.id ? " active" : ""),
          type: "button",
          onclick: () => { state.activeTabId = t.id; saveState(); render(); }
        }, t.label)
      );
    }
  }

  /* =========================
     Render: Top Split (calc only)
  ========================= */
  function renderTopSplit(tab) {
    const wrapper = h("div", { class: "top-split" });
    const layout = h("div", { class: "calc-layout top-grid" });
    layout.appendChild(renderSectionBox(tab));
    layout.appendChild(renderVarBox(tab));
    wrapper.appendChild(layout);
    return wrapper;
  }

  function renderSectionBox(tab) {
    const active = getActiveSection(tab);

    const box = h("div", { class: "rail-box section-box", id: "sectionBox" },
      h("div", { class: "rail-title" }, "구분명 리스트 (↑/↓ 이동)")
    );

    const list = h("div", { class: "section-list", id: "sectionList", tabindex: "0" });

    tab.sections.forEach((s, idx) => {
      const item = h("div", {
        class: "section-item" + (active && s.id === active.id ? " active" : ""),
        tabindex: "0",
        dataset: { secId: s.id, idx: String(idx) },
        onclick: () => {
          tab.activeSectionId = s.id;
          if (!tab.sectionsRows[s.id]) tab.sectionsRows[s.id] = makeDefaultRows(tab.type);
          saveState(); render();
          requestAnimationFrame(() => $(`.section-item[data-sec-id="${s.id}"]`)?.focus());
        },
        onkeydown: (e) => {
          if (e.key === "ArrowUp" || e.key === "ArrowDown") {
            e.preventDefault();
            const items = $$(".section-item", list);
            const cur = items.indexOf(e.currentTarget);
            const next = clamp(cur + (e.key === "ArrowDown" ? 1 : -1), 0, items.length - 1);
            items[next]?.focus();
          }
          if (e.key === "Enter") { e.preventDefault(); e.currentTarget.click(); }
        },
      },
        h("div", { class: "name" }, s.name || `구분 ${idx + 1}`),
        h("div", { class: "meta-inline" }, `개소: ${s.count ?? ""}`),
        h("div", { class: "meta" }, "선택")
      );
      list.appendChild(item);
    });

    const editor = h("div", { class: "section-editor" });
    const inpName = h("input", {
      type: "text",
      placeholder: "구분명",
      value: active?.name ?? "",
      oninput: (e) => {
        const s = getActiveSection(tab);
        if (!s) return;
        s.name = e.target.value;
        saveState();
      },
    });
    const inpCount = h("input", {
      type: "text",
      placeholder: "개소",
      value: active?.count ?? "",
      oninput: (e) => {
        const s = getActiveSection(tab);
        if (!s) return;
        s.count = e.target.value;
        saveState();
      },
    });
    const btnSave = h("button", { class: "smallbtn", type: "button", onclick: () => { saveState(); render(); } }, "저장");
    editor.append(inpName, inpCount, btnSave);

    const btnRow = h("div", { class: "row-actions" },
      h("button", { class: "smallbtn", type: "button", onclick: () => addSection(tab) }, "구분 추가 (Ctrl+F3)"),
      h("button", { class: "smallbtn", type: "button", onclick: () => deleteActiveSection(tab) }, "구분 삭제")
    );

    box.append(list, editor, btnRow);
    return box;
  }

  function renderVarBox(tab) {
    const sec = getActiveSection(tab);

    const box = h("div", { class: "rail-box var-box", id: "varBox" },
      h("div", { class: "rail-title" }, "변수표 (A, AB, A1, AB1... 최대 3자)")
    );

    const wrap = h("div", { class: "var-tablewrap", id: "varTableWrap" });
    const table = h("table", { class: "var-table" });

    const thead = h("thead", {}, h("tr", {}, h("th", {}, "변수"), h("th", {}, "산식"), h("th", {}, "값"), h("th", {}, "비고")));
    const tbody = h("tbody");
    const varMap = sec ? buildVarMap(sec) : {};

    (sec?.vars || []).forEach((r, rowIdx) => {
      if (sec && isValidVarName(r.key)) r.value = Number(varMap[r.key] ?? 0);
      else r.value = 0;

      const tdKey = h("td", {},
        h("input", {
          class: "cell",
          type: "text",
          placeholder: "예: A / AB / A1",
          value: r.key ?? "",
          dataset: { scope: "var", row: String(rowIdx), col: "key" },
          oninput: (e) => {
            if (!sec) return;
            let v = String(e.target.value || "").toUpperCase();
            v = v.replace(/[^A-Z0-9]/g, "").slice(0, 3);
            e.target.value = v;
            r.key = v;
            saveState();
          },
          onblur: () => saveState(),
        })
      );

      const tdExpr = h("td", {},
        h("input", {
          class: "cell",
          type: "text",
          placeholder: "예: (A+0.5)*2  (<...> 주석)",
          value: r.expr ?? "",
          dataset: { scope: "var", row: String(rowIdx), col: "expr" },
          oninput: (e) => { if (!sec) return; r.expr = e.target.value; saveState(); },
          onblur: () => { saveState(); render(); },
        })
      );

      const tdVal = h("td", {},
        h("input", { class: "cell readonly", type: "text", readonly: true, value: formatNumber(r.value) })
      );

      const tdNote = h("td", {},
        h("input", {
          class: "cell",
          type: "text",
          placeholder: "비고",
          value: r.note ?? "",
          dataset: { scope: "var", row: String(rowIdx), col: "note" },
          oninput: (e) => { if (!sec) return; r.note = e.target.value; saveState(); },
          onblur: () => saveState(),
        })
      );

      tbody.appendChild(h("tr", {}, tdKey, tdExpr, tdVal, tdNote));
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    box.appendChild(wrap);
    return box;
  }

  /* =========================
     Calc compute (복원 핵심)
     - value = eval(expr)
     - mult = eval(mult) or 1
     - conv = eval(conv) or 1
     - adj = value * mult * conv
  ========================= */
  function computeRow(row, varMap) {
    // VLOOKUP 적용 (코드 변경 여부와 무관하게 매 렌더 시 동기화)
    applyVlookupToRow(row);

    const value = safeEvalMath(row.expr, varMap);
    row.value = value;

    const mult = row.mult ? safeEvalMath(row.mult, varMap) : 1;
    const conv = row.conv ? safeEvalMath(row.conv, varMap) : 1;

    // NaN 방지
    const m = Number.isFinite(mult) && mult !== 0 ? mult : (row.mult ? 0 : 1);
    const c = Number.isFinite(conv) && conv !== 0 ? conv : (row.conv ? 0 : 1);

    row.adj = value * m * c;
  }

  /* =========================
     Render: Calc tab
  ========================= */
  function renderCalcTab(tab) {
    const sec = getActiveSection(tab);
    if (!sec) return h("div", { class: "panel" }, "구분을 먼저 생성해 주세요.");

    const rows = tab.sectionsRows[sec.id] || (tab.sectionsRows[sec.id] = makeDefaultRows(tab.type));
    const varMap = buildVarMap(sec);

    // 계산 + VLOOKUP
    rows.forEach((r) => computeRow(r, varMap));

    const panel = h("div", { class: "panel", id: "calcPanel" });

    panel.appendChild(
      h("div", { class: "panel-header" },
        h("div", {},
          h("div", { class: "panel-title" }, tab.label),
          h("div", { class: "panel-desc" }, "F2 편집모드 | 방향키: 셀 이동 | Ctrl+. 코드선택 | Ctrl+F3 구분추가 | Shift+Ctrl+F3 +10행")
        ),
        h("div", { class: "row-actions" },
          h("button", {
            class: "smallbtn",
            type: "button",
            onclick: () => {
              tab.sectionsRows[sec.id] = rows.concat(makeDefaultRows(tab.type).slice(0, 10));
              saveState(); render();
            }
          }, "+10행")
        )
      )
    );

    const wrap = h("div", { class: "table-wrap" });
    const table = h("table", {});
    const thead = h("thead", {},
      h("tr", {},
        h("th", {}, "No"),
        h("th", {}, "코드"),
        h("th", {}, "품명(자동)"),
        h("th", {}, "규격(자동)"),
        h("th", {}, "단위(자동)"),
        h("th", {}, "산출식"),
        h("th", {}, "물량(Value)"),
        h("th", {}, "할증(배수)"),
        h("th", {}, "환산단위"),
        h("th", {}, "할증후수량")
      )
    );

    const tbody = h("tbody");

    rows.forEach((r, i) => {
      const inpCode = h("input", {
        class: "cell",
        type: "text",
        value: r.code ?? "",
        placeholder: "코드",
        dataset: { scope: "calc", row: String(i), col: "code" },
        onfocus: (e) => rememberFocusedCalcCode(e.target),
        oninput: (e) => {
          r.code = e.target.value;
          // 코드 입력 즉시 VLOOKUP 반영
          applyVlookupToRow(r);
          saveState();
        },
        onblur: () => { saveState(); render(); }
      });

      const inpExpr = h("input", {
        class: "cell",
        type: "text",
        value: r.expr ?? "",
        placeholder: "예: (A+0.5)*2  (<...>는 주석)",
        dataset: { scope: "calc", row: String(i), col: "expr" },
        oninput: (e) => { r.expr = e.target.value; saveState(); },
        onblur: () => { saveState(); render(); }
      });

      const inpMult = h("input", {
        class: "cell",
        type: "text",
        value: r.mult ?? "",
        placeholder: "예: 1.05 or A",
        dataset: { scope: "calc", row: String(i), col: "mult" },
        oninput: (e) => { r.mult = e.target.value; saveState(); },
        onblur: () => { saveState(); render(); }
      });

      const inpConv = h("input", {
        class: "cell",
        type: "text",
        value: r.conv ?? "",
        placeholder: "예: 0.001 or AB",
        dataset: { scope: "calc", row: String(i), col: "conv" },
        oninput: (e) => { r.conv = e.target.value; saveState(); },
        onblur: () => { saveState(); render(); }
      });

      const tr = h("tr", { dataset: { row: String(i) } },
        h("td", {}, String(i + 1)),
        h("td", {}, inpCode),
        h("td", {}, h("input", { class: "cell readonly", type: "text", readonly: true, value: r.name ?? "" })),
        h("td", {}, h("input", { class: "cell readonly", type: "text", readonly: true, value: r.spec ?? "" })),
        h("td", {}, h("input", { class: "cell readonly", type: "text", readonly: true, value: r.unit ?? "" })),
        h("td", {}, inpExpr),
        h("td", {}, h("input", { class: "cell readonly", type: "text", readonly: true, value: formatNumber(r.value) })),
        h("td", {}, inpMult),
        h("td", {}, inpConv),
        h("td", {}, h("input", { class: "cell readonly", type: "text", readonly: true, value: formatNumber(r.adj) })),
      );

      tbody.appendChild(tr);
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    panel.appendChild(wrap);

    return panel;
  }

  /* =========================
     Summary (집계 탭 복원)
     - 철골_집계 / 구조이기/동바리_집계
     - 원칙:
       * 철골/부자재 집계: 할증후수량(adj) 합계
       * 동바리 집계: 물량(value) 합계  (요구사항)
  ========================= */
  function aggregateForTab(calcTabId, mode) {
    // mode: "adj" | "value"
    const calcTab = state.tabs[calcTabId];
    if (!calcTab || !isCalcTab(calcTab)) return [];

    const out = new Map(); // code -> {code,name,spec,unit,sum}
    const libLookup = (code) => lookupCode(code);

    for (const sec of calcTab.sections) {
      const varMap = buildVarMap(sec);
      const rows = calcTab.sectionsRows[sec.id] || [];
      for (const r of rows) {
        if (!normalizeCode(r.code)) continue;

        // 계산 최신화
        computeRow(r, varMap);

        const key = normalizeCode(r.code);
        const found = libLookup(key);
        const name = r.name || found?.name || "";
        const spec = r.spec || found?.spec || "";
        const unit = r.unit || found?.unit || "";

        const add = mode === "adj" ? (Number(r.adj) || 0) : (Number(r.value) || 0);
        if (!out.has(key)) out.set(key, { code: key, name, spec, unit, sum: 0 });
        out.get(key).sum += add;
      }
    }

    return Array.from(out.values()).sort((a, b) => a.code.localeCompare(b.code));
  }

  function renderSummaryTab(tab) {
    // 어떤 집계를 보여줄지 매핑
    let rows = [];
    let title = tab.label;
    let note = "";

    if (tab.type === "summary_steel") {
      // 철골_집계 = 철골(calc_steel) + 철골_부자재(calc_aux) "할증후수량" 합
      const a = aggregateForTab("steel", "adj");
      const b = aggregateForTab("steel_aux", "adj");
      const map = new Map();
      for (const r of [...a, ...b]) {
        if (!map.has(r.code)) map.set(r.code, { ...r });
        else map.get(r.code).sum += r.sum;
      }
      rows = Array.from(map.values()).sort((x, y) => x.code.localeCompare(y.code));
      note = "※ 철골/부자재 집계: “할증후수량” 합계";
    } else if (tab.type === "summary_support") {
      // 동바리_집계 = support(calc_support) "물량(Value)" 합
      rows = aggregateForTab("support", "value");
      note = "※ 동바리 집계: “물량(Value)” 합계";
    } else {
      note = "집계 탭";
    }

    const panel = h("div", { class: "panel" },
      h("div", { class: "panel-header" },
        h("div", {},
          h("div", { class: "panel-title" }, title),
          h("div", { class: "panel-desc" }, note)
        )
      )
    );

    const wrap = h("div", { class: "table-wrap" });
    const table = h("table", {});
    const thead = h("thead", {},
      h("tr", {},
        h("th", {}, "No"),
        h("th", {}, "코드"),
        h("th", {}, "품명"),
        h("th", {}, "규격"),
        h("th", {}, "단위"),
        h("th", {}, "합계")
      )
    );
    const tbody = h("tbody");

    if (!rows.length) {
      tbody.appendChild(
        h("tr", {}, h("td", { colspan: "6", style: "padding:14px; color:rgba(0,0,0,.55);" }, "집계할 데이터가 없습니다."))
      );
    } else {
      rows.forEach((r, i) => {
        tbody.appendChild(
          h("tr", {},
            h("td", {}, String(i + 1)),
            h("td", {}, r.code),
            h("td", {}, r.name || ""),
            h("td", {}, r.spec || ""),
            h("td", {}, r.unit || ""),
            h("td", {}, formatNumber(r.sum))
          )
        );
      });
    }

    table.append(thead, tbody);
    wrap.appendChild(table);
    panel.appendChild(wrap);

    return panel;
  }

  /* =========================
     Code tab (코드 DB 마스터)
     - XLSX 업로드 + 편집 + 로컬 저장
  ========================= */
  function renderCodeTab() {
    const panel = h("div", { class: "panel" });

    panel.appendChild(
      h("div", { class: "panel-header" },
        h("div", {},
          h("div", { class: "panel-title" }, "코드(Ctrl+.)"),
          h("div", { class: "panel-desc" }, "코드 입력 → 산출탭에서 품명/규격/단위 자동채움(VLOOKUP) + 환산/할증 계산")
        ),
        h("div", { class: "row-actions" },
          h("label", { class: "smallbtn" },
            "코드DB 업로드(XLSX)",
            h("input", {
              type: "file",
              accept: ".xlsx,.xls",
              hidden: true,
              onchange: async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try {
                  if (!window.XLSX) throw new Error("XLSX 라이브러리를 찾을 수 없습니다.");
                  const data = await f.arrayBuffer();
                  const wb = XLSX.read(data, { type: "array" });
                  const ws = wb.Sheets[wb.SheetNames[0]];
                  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

                  const mapped = json.map((r) => {
                    const code = r.code ?? r.CODE ?? r["코드"] ?? r["Code"] ?? "";
                    const name = r.name ?? r.NAME ?? r["품명"] ?? r["품명(자동)"] ?? "";
                    const spec = r.spec ?? r.SPEC ?? r["규격"] ?? r["규격(자동)"] ?? "";
                    const unit = r.unit ?? r.UNIT ?? r["단위"] ?? r["단위(자동)"] ?? "";
                    return {
                      code: String(code).trim(),
                      name: String(name).trim(),
                      spec: String(spec).trim(),
                      unit: String(unit).trim(),
                    };
                  }).filter((r) => r.code);

                  // 중복 코드 정리(마지막 값 우선)
                  const m = new Map();
                  for (const r of mapped) m.set(normalizeCode(r.code), r);
                  state.tabs.code.codeLibrary = Array.from(m.values());

                  saveState();
                  render();
                  alert(`코드 DB 업로드 완료: ${state.tabs.code.codeLibrary.length}개`);
                } catch (err) {
                  console.error(err);
                  alert("XLSX 업로드 실패: 파일 형식/시트 구조를 확인해줘.");
                } finally {
                  e.target.value = "";
                }
              }
            })
          ),
          h("button", { class: "smallbtn", type: "button", onclick: () => openCodePicker() }, "코드 선택(Ctrl+.)"),
          h("button", {
            class: "smallbtn",
            type: "button",
            onclick: () => {
              // 빈 줄 10개 추가
              const lib = getCodeLibrary();
              for (let i = 0; i < 10; i++) lib.push({ code: "", name: "", spec: "", unit: "" });
              state.tabs.code.codeLibrary = lib;
              saveState();
              render();
            }
          }, "+10행")
        )
      )
    );

    const lib = getCodeLibrary();

    const wrap = h("div", { class: "table-wrap", style: "margin-top:10px; max-height: 70vh;" });
    const table = h("table", {});
    const thead = h("thead", {}, h("tr", {}, h("th", {}, "No"), h("th", {}, "코드"), h("th", {}, "품명"), h("th", {}, "규격"), h("th", {}, "단위")));
    const tbody = h("tbody");

    const show = lib.length ? lib : Array.from({ length: 20 }).map(() => ({ code: "", name: "", spec: "", unit: "" }));

    show.forEach((r, i) => {
      tbody.appendChild(
        h("tr", {},
          h("td", {}, String(i + 1)),
          h("td", {}, h("input", {
            class: "cell",
            type: "text",
            value: r.code ?? "",
            oninput: (e) => { r.code = e.target.value; saveState(); },
            onblur: () => { r.code = normalizeCode(r.code); saveState(); }
          })),
          h("td", {}, h("input", { class: "cell", type: "text", value: r.name ?? "", oninput: (e) => { r.name = e.target.value; saveState(); } })),
          h("td", {}, h("input", { class: "cell", type: "text", value: r.spec ?? "", oninput: (e) => { r.spec = e.target.value; saveState(); } })),
          h("td", {}, h("input", { class: "cell", type: "text", value: r.unit ?? "", oninput: (e) => { r.unit = e.target.value; saveState(); } })),
        )
      );
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    panel.appendChild(wrap);

    return panel;
  }

  /* =========================
     Render
  ========================= */
  function render() {
    renderTabs();
    const tab = getActiveTab();
    viewEl.innerHTML = "";

    if (isCalcTab(tab)) {
      viewEl.appendChild(renderTopSplit(tab));
      viewEl.appendChild(renderCalcTab(tab));
    } else if (tab.type === "code") {
      viewEl.appendChild(renderCodeTab());
    } else if (isSummaryTab(tab)) {
      viewEl.appendChild(renderSummaryTab(tab));
    } else {
      viewEl.appendChild(h("div", { class: "panel" }, "탭을 선택해 주세요."));
    }

    requestAnimationFrame(updateStickyVars);
  }

  /* =========================
     Global key bindings (capture)
  ========================= */
  function handleGlobalKeydown(e) {
    // 방향키 이동 우선
    handleVarKeydown(e);
    handleCalcKeydown(e);

    // F2 편집모드
    if (e.key === "F2") {
      e.preventDefault();
      toggleEditMode();
      return;
    }

    // Ctrl+. 코드 선택
    if (e.ctrlKey && e.key === ".") {
      e.preventDefault();
      openCodePicker();
      return;
    }

    // Ctrl+F3: 구분 추가
    if (e.key === "F3" && e.ctrlKey && !e.shiftKey) {
      const tab = getActiveTab();
      if (isCalcTab(tab)) {
        e.preventDefault();
        addSection(tab);
      }
      return;
    }

    // Shift+Ctrl+F3: +10행
    if (e.key === "F3" && e.ctrlKey && e.shiftKey) {
      const tab = getActiveTab();
      if (isCalcTab(tab)) {
        e.preventDefault();
        const sec = getActiveSection(tab);
        if (!sec) return;
        const rows = tab.sectionsRows[sec.id] || (tab.sectionsRows[sec.id] = makeDefaultRows(tab.type));
        tab.sectionsRows[sec.id] = rows.concat(makeDefaultRows(tab.type).slice(0, 10));
        saveState();
        render();
      }
      return;
    }
  }

  /* =========================
     Top buttons
  ========================= */
  function wireTopButtons() {
    $("#btnReset")?.addEventListener("click", () => {
      state = makeDefaultState();
      saveState();
      render();
    });

    $("#btnExport")?.addEventListener("click", () => {
      const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "FIN_state.json";
      a.click();
      URL.revokeObjectURL(a.href);
    });

    $("#fileImport")?.addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      try {
        const text = await f.text();
        state = JSON.parse(text);
        saveState();
        render();
      } catch {
        alert("가져오기 실패: JSON 형식이 올바르지 않습니다.");
      } finally {
        e.target.value = "";
      }
    });

    $("#btnOpenPicker")?.addEventListener("click", () => openCodePicker());
  }

  /* =========================
     Init
  ========================= */
  document.addEventListener("keydown", handleGlobalKeydown, true);
  window.addEventListener("resize", updateStickyVars);

  wireTopButtons();
  render();
})();
