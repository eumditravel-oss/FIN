/* app.js (v14) */
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
  const LS_KEY = "FIN_WEB_V14_STATE";
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
    Array.from({ length: 10 }).map(() => ({ key: "", expr: "", value: 0, note: "" }));

  const makeDefaultRows = (calcType) => {
    const rows = Array.from({ length: 12 }).map(() => ({
      code: "",
      name: "",
      spec: "",
      unit: "",
      expr: "",
      value: 0,
      mult: "",
      conv: "",
    }));
    if (calcType === "calc_steel") rows.forEach((r) => (r.unit = "M"));
    if (calcType === "calc_aux") rows.forEach((r) => (r.unit = "M2"));
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

        // code DB (global shared 느낌)
        codeLibrary: [], // [{code,name,spec,unit}]
      };

      if (t.type.startsWith("calc_")) {
        const sec = makeSection("1층 바닥 철골보", "1");
        tabs[t.id].sections = [sec];
        tabs[t.id].activeSectionId = sec.id;
        tabs[t.id].sectionsRows[sec.id] = makeDefaultRows(t.type);
      }
    }
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
     Expression eval
  ========================= */
  const stripAngleComments = (s) => (s ? String(s).replace(/<[^>]*>/g, "") : "");
  const normalizeExpr = (s) => stripAngleComments(s).replace(/\s+/g, " ").trim();
  const isValidVarName = (v) => /^[A-Z][A-Z0-9]{0,2}$/.test(v || "");

  function safeEvalMath(expr, varMap) {
    const raw = normalizeExpr(expr);
    if (!raw) return 0;
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
    // 포커스된 input에 .editing 표시(스타일에서 사용)
    const f = document.activeElement;
    if (f instanceof HTMLInputElement || f instanceof HTMLTextAreaElement) {
      if (on) f.classList.add("editing");
      else f.classList.remove("editing");
    }
  }

  function toggleEditMode() {
    setEditMode(!state.ui.editMode);
    saveState();
  }

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

  function moveVarCell(row, col, dRow, dCol) {
    const tab = getActiveTab();
    const sec = getActiveSection(tab);
    if (!sec) return;
    const maxRow = sec.vars.length - 1;
    const colIdx = VAR_COLS.indexOf(col);
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

    // 단축키는 전역에서 처리하되, 방향키는 여기서 확실히 먹는다.
    if (["ArrowUp","ArrowDown","ArrowLeft","ArrowRight","Enter"].includes(e.key)) {
      e.stopPropagation();
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    const c = caretInfo(t);

    // 편집모드 OFF면: 방향키는 "셀 이동"이 우선
    if (!state.ui.editMode) {
      if (e.key === "ArrowUp") { e.preventDefault(); moveVarCell(row, col, -1, 0); return; }
      if (e.key === "ArrowDown") { e.preventDefault(); moveVarCell(row, col, +1, 0); return; }
      if (e.key === "ArrowLeft") { e.preventDefault(); moveVarCell(row, col, 0, -1); return; }
      if (e.key === "ArrowRight") { e.preventDefault(); moveVarCell(row, col, 0, +1); return; }
      if (e.key === "Enter") {
        e.preventDefault();
        if (col === "key") focusVarCell(row, "expr");
        else if (col === "expr") { saveState(); render(); requestAnimationFrame(() => focusVarCell(row, "note")); }
        else focusVarCell(clamp(row + 1, 0, 9999), "key");
        return;
      }
      return;
    }

    // 편집모드 ON이면: 커서가 끝에 있을 때만 셀 이동
    if (e.key === "ArrowUp") { e.preventDefault(); moveVarCell(row, col, -1, 0); return; }
    if (e.key === "ArrowDown") { e.preventDefault(); moveVarCell(row, col, +1, 0); return; }
    if (e.key === "ArrowLeft") {
      if (c.s === 0 && c.e === 0 && col !== "key") { e.preventDefault(); moveVarCell(row, col, 0, -1); }
      return;
    }
    if (e.key === "ArrowRight") {
      if (c.s === c.len && c.e === c.len && col !== "note") { e.preventDefault(); moveVarCell(row, col, 0, +1); }
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
     Keyboard nav: CALC table
  ========================= */
  const CALC_COLS = ["code","expr","mult","conv"]; // 이동 대상(주요 입력칸만)

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
        // Enter = 계산 트리거(특히 expr에서)
        saveState(); render();
        requestAnimationFrame(() => focusCalcCell(row, col));
        return;
      }
      return;
    }

    // 편집모드 ON: 끝에서만 이동
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
  ========================= */
  function openCodePicker() {
    const tab = getActiveTab();
    // 코드 탭이 아니어도 선택 가능(활성 탭의 코드칸에 넣는 구조)
    // 지금은 "표시 + 선택" 기반만 구현. 실제 삽입 대상은 다음 단계에서 정확히 연결.
    const lib = (state.tabs.code?.codeLibrary || []);

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
      h("button", {
        class: "smallbtn",
        type: "button",
        onclick: () => {
          // TODO: 실제 삽입(포커스된 calc code 칸)에 연결
          alert("선택 삽입 연결은 다음 단계에서 ‘현재 포커스된 코드칸’에 넣도록 정확히 붙여줄게.");
        }
      }, "삽입(Ctrl+Enter)")
    );

    modal.append(head, search, listWrap, foot);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const selected = new Set();

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

      rows.slice(0, 200).forEach((r, i) => {
        const key = `${r.code}__${i}`;
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

    function close() {
      overlay.remove();
      document.removeEventListener("keydown", onKey, true);
    }

    function onKey(e) {
      // modal 내부 단축키
      if (e.key === "Escape") { e.preventDefault(); close(); return; }
      if (e.ctrlKey && e.key.toLowerCase() === "b") {
        // 다중선택 토글은 체크박스로 이미 가능
        e.preventDefault();
        return;
      }
      if (e.ctrlKey && e.key === "Enter") {
        e.preventDefault();
        alert("선택 삽입 연결은 다음 단계에서 ‘현재 포커스된 코드칸’에 넣도록 정확히 붙여줄게.");
      }
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
     Render: Calc table
  ========================= */
  function renderCalcTab(tab) {
    const sec = getActiveSection(tab);
    if (!sec) return h("div", { class: "panel" }, "구분을 먼저 생성해 주세요.");

    const rows = tab.sectionsRows[sec.id] || (tab.sectionsRows[sec.id] = makeDefaultRows(tab.type));
    const varMap = buildVarMap(sec);

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
        h("th", {}, "환산단위")
      )
    );

    const tbody = h("tbody");

    rows.forEach((r, i) => {
      r.value = safeEvalMath(r.expr, varMap);

      const inpCode = h("input", {
        class: "cell",
        type: "text",
        value: r.code ?? "",
        placeholder: "코드",
        dataset: { scope: "calc", row: String(i), col: "code" },
        oninput: (e) => { r.code = e.target.value; saveState(); },
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
        placeholder: "",
        dataset: { scope: "calc", row: String(i), col: "mult" },
        oninput: (e) => { r.mult = e.target.value; saveState(); },
      });

      const inpConv = h("input", {
        class: "cell",
        type: "text",
        value: r.conv ?? "",
        placeholder: "",
        dataset: { scope: "calc", row: String(i), col: "conv" },
        oninput: (e) => { r.conv = e.target.value; saveState(); },
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
      );

      tbody.appendChild(tr);
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    panel.appendChild(wrap);

    return panel;
  }

  /* =========================
     Render: Summary tabs
  ========================= */
  function renderSummaryTab(tab) {
    return h("div", { class: "panel" },
      h("div", { class: "panel-header" },
        h("div", {},
          h("div", { class: "panel-title" }, tab.label),
          h("div", { class: "panel-desc" }, "집계 탭은 현재 산출 데이터 합계 표시 영역입니다.")
        )
      ),
      h("div", {}, "※ 철골/부자재 집계: “할증후수량” 합계 · 동바리 집계: “물량(Value)” 합계")
    );
  }

  /* =========================
     Code tab (복구형)
  ========================= */
  function renderCodeTab(tab) {
    const panel = h("div", { class: "panel" });

    panel.appendChild(
      h("div", { class: "panel-header" },
        h("div", {},
          h("div", { class: "panel-title" }, "코드(Ctrl+.)"),
          h("div", { class: "panel-desc" }, "XLSX 업로드로 코드 DB 등록 → Ctrl+. 코드선택으로 산출표 코드칸에 삽입(연결 단계)")
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
                  // xlsx 라이브러리 필요 (index.html에 CDN 있음)
                  const data = await f.arrayBuffer();
                  const wb = XLSX.read(data, { type: "array" });
                  const ws = wb.Sheets[wb.SheetNames[0]];
                  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

                  // 기대 컬럼 후보: code, name, spec, unit / 또는 한글 컬럼도 대응
                  const mapped = json.map((r) => {
                    const code = r.code ?? r.CODE ?? r["코드"] ?? r["Code"] ?? "";
                    const name = r.name ?? r.NAME ?? r["품명"] ?? r["품명(자동)"] ?? "";
                    const spec = r.spec ?? r.SPEC ?? r["규격"] ?? r["규격(자동)"] ?? "";
                    const unit = r.unit ?? r.UNIT ?? r["단위"] ?? r["단위(자동)"] ?? "";
                    return { code: String(code).trim(), name: String(name).trim(), spec: String(spec).trim(), unit: String(unit).trim() };
                  }).filter((r) => r.code || r.name || r.spec || r.unit);

                  state.tabs.code.codeLibrary = mapped;
                  saveState();
                  render();
                  alert(`코드 DB 업로드 완료: ${mapped.length}개`);
                } catch (err) {
                  console.error(err);
                  alert("XLSX 업로드 실패: 파일 형식/시트 구조를 확인해줘.");
                } finally {
                  e.target.value = "";
                }
              }
            })
          ),
          h("button", { class: "smallbtn", type: "button", onclick: () => openCodePicker() }, "코드 선택(Ctrl+.)")
        )
      )
    );

    const lib = state.tabs.code.codeLibrary || [];

    const wrap = h("div", { class: "table-wrap", style: "margin-top:10px; max-height: 64vh;" });
    const table = h("table", {});
    const thead = h("thead", {}, h("tr", {}, h("th", {}, "No"), h("th", {}, "코드"), h("th", {}, "품명"), h("th", {}, "규격"), h("th", {}, "단위")));
    const tbody = h("tbody");

    const show = lib.length ? lib : Array.from({ length: 20 }).map(() => ({ code: "", name: "", spec: "", unit: "" }));

    show.forEach((r, i) => {
      tbody.appendChild(
        h("tr", {},
          h("td", {}, String(i + 1)),
          h("td", {}, h("input", { class: "cell", type: "text", value: r.code ?? "", oninput: (e) => { r.code = e.target.value; saveState(); } })),
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
      viewEl.appendChild(renderCodeTab(tab));
    } else if (isSummaryTab(tab)) {
      viewEl.appendChild(renderSummaryTab(tab));
    } else {
      viewEl.appendChild(h("div", { class: "panel" }, "탭을 선택해 주세요."));
    }

    requestAnimationFrame(updateStickyVars);
  }

  /* =========================
     Global key bindings (capture, 최상위)
  ========================= */
  function handleGlobalKeydown(e) {
    // ✅ 1) 셀 네비게이션 우선
    handleVarKeydown(e);
    handleCalcKeydown(e);

    // ✅ 2) F2 편집모드 토글
    if (e.key === "F2") {
      e.preventDefault();
      toggleEditMode();
      return;
    }

    // ✅ 3) Ctrl+. 코드 선택
    if (e.ctrlKey && e.key === ".") {
      e.preventDefault();
      openCodePicker();
      return;
    }

    // ✅ 4) Ctrl+F3: 구분 추가 (calc 탭에서만)
    if (e.key === "F3" && e.ctrlKey && !e.shiftKey) {
      const tab = getActiveTab();
      if (isCalcTab(tab)) {
        e.preventDefault();
        addSection(tab);
      }
      return;
    }

    // ✅ 5) Shift+Ctrl+F3: +10행 (calc 탭에서만)
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

    // 상단 코드 선택 버튼도 Ctrl+.와 동일 동작
    $("#btnOpenPicker")?.addEventListener("click", () => openCodePicker());
  }

  /* =========================
     Init
  ========================= */
  document.addEventListener("keydown", handleGlobalKeydown, true); // ✅ 캡처 단계에서 무조건 받음
  window.addEventListener("resize", updateStickyVars);

  wireTopButtons();
  render();
})();
