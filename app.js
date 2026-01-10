/* app.js (v12) */
(() => {
  "use strict";

  const $ = (sel, el = document) => el.querySelector(sel);
  const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

  function el(tag, attrs = {}, ...children) {
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

  function cssEscape(s) {
    return String(s).replace(/"/g, '\\"');
  }

  function cryptoRandomId() {
    try { return crypto.getRandomValues(new Uint32Array(2)).join("-"); }
    catch { return String(Date.now()) + "-" + Math.floor(Math.random() * 1e9); }
  }

  /* Tabs */
  const TAB_DEFS = [
    { id: "code", label: "코드(Ctrl+.)", type: "code" },
    { id: "steel", label: "철골", type: "calc_steel" },
    { id: "steel_sum", label: "철골_집계", type: "summary_steel" },
    { id: "steel_aux", label: "철골_부자재", type: "calc_aux" },
    { id: "support", label: "구조이기/동바리", type: "calc_support" },
    { id: "support_sum", label: "구조이기/동바리_집계", type: "summary_support" },
  ];

  /* Storage */
  const LS_KEY = "FIN_WEB_V12_STATE";
  const loadState = () => { try { const r = localStorage.getItem(LS_KEY); return r ? JSON.parse(r) : null; } catch { return null; } };
  const saveState = () => { try { localStorage.setItem(LS_KEY, JSON.stringify(state)); } catch {} };

  /* Model */
  const makeDefaultVars = () => Array.from({ length: 10 }).map(() => ({ key:"", expr:"", value:0, note:"" }));
  const makeDefaultRows = (calcType) => {
    const base = Array.from({ length: 12 }).map(() => ({
      code:"", name:"", spec:"", unit:"", expr:"", value:0, mult:"", conv:""
    }));
    if (calcType === "calc_steel") base.forEach(r => (r.unit = "M"));
    if (calcType === "calc_aux") base.forEach(r => (r.unit = "M2"));
    return base;
  };
  const makeSection = (name="구분 1", count="") => ({ id: cryptoRandomId(), name, count, vars: makeDefaultVars() });

  function makeDefaultState() {
    const tabs = {};
    for (const t of TAB_DEFS) {
      tabs[t.id] = { id:t.id, label:t.label, type:t.type, sections:[], activeSectionId:null, sectionsRows:{} };
      if (t.type.startsWith("calc_")) {
        const sec = makeSection("1층 바닥 철골보", "1");
        tabs[t.id].sections = [sec];
        tabs[t.id].activeSectionId = sec.id;
        tabs[t.id].sectionsRows[sec.id] = makeDefaultRows(t.type);
      }
    }
    return { activeTabId:"steel", tabs };
  }

  let state = loadState() || makeDefaultState();

  /* ✅ Sticky var 계산 (초록 공백 제거 핵심) */
  function updateStickyVars() {
    const topbar = $(".topbar");
    const tabs = $("#tabs");
    const topbarH = topbar ? Math.ceil(topbar.getBoundingClientRect().height) : 0;
    const tabsH = tabs ? Math.ceil(tabs.getBoundingClientRect().height) : 0;
    document.documentElement.style.setProperty("--topbarH", `${topbarH}px`);
    document.documentElement.style.setProperty("--tabsH", `${tabsH}px`);
  }

  /* Eval */
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
    for (let i=0;i<4;i++) for (const r of section.vars) if (isValidVarName(r.key)) map[r.key] = safeEvalMath(r.expr, map);
    return map;
  }

  const formatNumber = (n) => {
    const num = Number(n);
    if (!Number.isFinite(num)) return "0";
    const x = Math.round(num * 1000000) / 1000000;
    return String(x);
  };

  /* Active */
  const getActiveTab = () => state.tabs[state.activeTabId];
  const getActiveSection = (tab) => {
    if (!tab?.sections?.length) return null;
    const id = tab.activeSectionId || tab.sections[0].id;
    return tab.sections.find(s => s.id === id) || tab.sections[0];
  };

  /* Render targets */
  const tabsEl = $("#tabs");
  const viewEl = $("#view");

  function renderTabs() {
    tabsEl.innerHTML = "";
    for (const t of TAB_DEFS) {
      tabsEl.appendChild(
        el("button", {
          class: "tab" + (state.activeTabId === t.id ? " active" : ""),
          type: "button",
          onclick: () => { state.activeTabId = t.id; saveState(); render(); }
        }, t.label)
      );
    }
  }

  function addSection(tab) {
    const sec = makeSection(`구분 ${tab.sections.length + 1}`, "");
    tab.sections.push(sec);
    tab.activeSectionId = sec.id;
    tab.sectionsRows[sec.id] = makeDefaultRows(tab.type);
    saveState(); render();
    requestAnimationFrame(() => $(`.section-item[data-sec-id="${cssEscape(sec.id)}"]`)?.focus());
  }

  function deleteActiveSection(tab) {
    if (tab.sections.length <= 1) return;
    const cur = getActiveSection(tab);
    const idx = tab.sections.findIndex(s => s.id === cur.id);
    tab.sections.splice(idx, 1);
    delete tab.sectionsRows[cur.id];
    const next = tab.sections[clamp(idx, 0, tab.sections.length - 1)];
    tab.activeSectionId = next.id;
    saveState(); render();
  }

  /* ===== 변수표 키보드 이동 (v12 fix) ===== */
  const VAR_COLS = ["key", "expr", "note"];

  function caretInfo(input) {
    try {
      return { start: input.selectionStart ?? 0, end: input.selectionEnd ?? 0, len: (input.value ?? "").length };
    } catch {
      return { start: 0, end: 0, len: (input.value ?? "").length };
    }
  }

  function focusVarCell(row, col) {
    const box = $("#varBox");
    if (!box) return;
    const q = `input[data-scope="var"][data-row="${row}"][data-col="${col}"]`;
    const target = box.querySelector(q);
    if (target) {
      target.focus();
      if (col === "key") target.select?.();
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

  function onVarKeydownFactory(rowIdx, colName) {
    return (e) => {
      if (e.ctrlKey || e.metaKey || e.altKey) return;

      // ✅ 밖으로 새지 않게 강제
      if (["ArrowUp","ArrowDown","ArrowLeft","ArrowRight","Enter"].includes(e.key)) {
        e.stopPropagation();
      }

      const input = e.currentTarget;
      const c = caretInfo(input);

      if (e.key === "ArrowUp") { e.preventDefault(); moveVarCell(rowIdx, colName, -1, 0); return; }
      if (e.key === "ArrowDown") { e.preventDefault(); moveVarCell(rowIdx, colName, +1, 0); return; }

      // ✅ 좌/우 규칙을 명확히 고정 (표 밖 탈출 없음)
      if (e.key === "ArrowLeft") {
        if (colName === "key") return; // key에서는 좌로 이동시키지 않음(커서만)
        if (c.start === 0 && c.end === 0) { e.preventDefault(); moveVarCell(rowIdx, colName, 0, -1); }
        return;
      }

      if (e.key === "ArrowRight") {
        if (colName === "note") return; // note에서는 우로 이동시키지 않음(커서만)
        if (c.start === c.len && c.end === c.len) { e.preventDefault(); moveVarCell(rowIdx, colName, 0, +1); }
        return;
      }

      if (e.key === "Enter") {
        e.preventDefault();
        if (colName === "key") focusVarCell(rowIdx, "expr");
        else if (colName === "expr") {
          saveState();
          render();
          requestAnimationFrame(() => focusVarCell(rowIdx, "note"));
        } else {
          focusVarCell(clamp(rowIdx + 1, 0, 99999), "key");
        }
      }
    };
  }

  /* ===== Top split render ===== */
  function renderTopSplit(tab) {
    const wrapper = el("div", { class: "top-split" });
    const layout = el("div", { class: "calc-layout top-grid" });
    layout.appendChild(renderSectionBox(tab));
    layout.appendChild(renderVarBox(tab));
    wrapper.appendChild(layout);
    return wrapper;
  }

  function renderSectionBox(tab) {
    const active = getActiveSection(tab);

    const box = el("div", { class: "rail-box section-box", id: "sectionBox" },
      el("div", { class: "rail-title" }, "구분명 리스트 (↑/↓ 이동)")
    );

    const list = el("div", { class: "section-list", id: "sectionList", tabindex: "0" });

    tab.sections.forEach((s, idx) => {
      const item = el("div", {
        class: "section-item" + (active && s.id === active.id ? " active" : ""),
        tabindex: "0",
        dataset: { secId: s.id, idx: String(idx) },
        onclick: () => {
          tab.activeSectionId = s.id;
          if (!tab.sectionsRows[s.id]) tab.sectionsRows[s.id] = makeDefaultRows(tab.type);
          saveState(); render();
          requestAnimationFrame(() => $(`.section-item[data-sec-id="${cssEscape(s.id)}"]`)?.focus());
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
        }
      },
        el("div", { class: "name" }, s.name || `구분 ${idx + 1}`),
        el("div", { class: "meta-inline" }, `개소: ${s.count ?? ""}`),
        el("div", { class: "meta" }, "선택")
      );
      list.appendChild(item);
    });

    const editor = el("div", { class: "section-editor" });
    const inpName = el("input", {
      type: "text",
      placeholder: "구분명",
      value: active?.name ?? "",
      oninput: (e) => { const s = getActiveSection(tab); if (!s) return; s.name = e.target.value; saveState(); }
    });
    const inpCount = el("input", {
      type: "text",
      placeholder: "개소",
      value: active?.count ?? "",
      oninput: (e) => { const s = getActiveSection(tab); if (!s) return; s.count = e.target.value; saveState(); }
    });
    const btnSave = el("button", { class:"smallbtn", type:"button", onclick:()=>{ saveState(); render(); } }, "저장");
    editor.append(inpName, inpCount, btnSave);

    const btnRow = el("div", { class: "row-actions" },
      el("button", { class:"smallbtn", type:"button", onclick:()=>addSection(tab) }, "구분 추가 (Ctrl+F3)"),
      el("button", { class:"smallbtn", type:"button", onclick:()=>deleteActiveSection(tab) }, "구분 삭제")
    );

    box.append(list, editor, btnRow);
    return box;
  }

  function renderVarBox(tab) {
    const sec = getActiveSection(tab);

    const box = el("div", { class: "rail-box var-box", id: "varBox" },
      el("div", { class: "rail-title" }, "변수표 (A, AB, A1, AB1... 최대 3자)")
    );

    const wrap = el("div", { class: "var-tablewrap", id: "varTableWrap" });
    const table = el("table", { class: "var-table" });

    const thead = el("thead", {}, el("tr", {},
      el("th", {}, "변수"),
      el("th", {}, "산식"),
      el("th", {}, "값"),
      el("th", {}, "비고"),
    ));

    const tbody = el("tbody");
    const varMap = sec ? buildVarMap(sec) : {};

    (sec?.vars || []).forEach((r, rowIdx) => {
      if (sec && isValidVarName(r.key)) r.value = Number(varMap[r.key] ?? 0);
      else r.value = 0;

      const tdKey = el("td", {}, el("input", {
        class: "cell",
        type: "text",
        placeholder: "예: A / AB / A1",
        value: r.key ?? "",
        dataset: { scope:"var", row:String(rowIdx), col:"key" },
        oninput: (e) => {
          if (!sec) return;
          let v = String(e.target.value || "").toUpperCase();
          v = v.replace(/[^A-Z0-9]/g, "").slice(0, 3);
          e.target.value = v;
          r.key = v;
          saveState();
        },
        onkeydown: onVarKeydownFactory(rowIdx, "key"),
        onblur: () => { saveState(); }
      }));

      const tdExpr = el("td", {}, el("input", {
        class: "cell",
        type: "text",
        placeholder: "예: (A+0.5)*2  (<...> 주석)",
        value: r.expr ?? "",
        dataset: { scope:"var", row:String(rowIdx), col:"expr" },
        oninput: (e) => { if (!sec) return; r.expr = e.target.value; saveState(); },
        onkeydown: onVarKeydownFactory(rowIdx, "expr"),
        onblur: () => { saveState(); render(); }
      }));

      const tdVal = el("td", {}, el("input", {
        class: "cell readonly",
        type: "text",
        readonly: true,
        value: formatNumber(r.value)
      }));

      const tdNote = el("td", {}, el("input", {
        class: "cell",
        type: "text",
        placeholder: "비고",
        value: r.note ?? "",
        dataset: { scope:"var", row:String(rowIdx), col:"note" },
        oninput: (e) => { if (!sec) return; r.note = e.target.value; saveState(); },
        onkeydown: onVarKeydownFactory(rowIdx, "note"),
        onblur: () => { saveState(); }
      }));

      tbody.appendChild(el("tr", {}, tdKey, tdExpr, tdVal, tdNote));
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    box.appendChild(wrap);
    return box;
  }

  /* Calc table (동일) */
  function renderCalcTab(tab) {
    const sec = getActiveSection(tab);
    if (!sec) return el("div", { class: "panel" }, "구분을 먼저 생성해 주세요.");

    const rows = tab.sectionsRows[sec.id] || (tab.sectionsRows[sec.id] = makeDefaultRows(tab.type));
    const varMap = buildVarMap(sec);

    const panel = el("div", { class: "panel" });
    panel.appendChild(
      el("div", { class: "panel-header" },
        el("div", {},
          el("div", { class: "panel-title" }, tab.label),
          el("div", { class: "panel-desc" },
            "구분(↑/↓) 이동 → 해당 구분의 변수/산출표 전환 | 산출식 Enter 계산 | Ctrl+. 코드선택 | Ctrl+F3 구분추가"
          )
        ),
        el("div", { class: "row-actions" },
          el("button", { class: "smallbtn", type: "button", onclick: () => {
            const extra = makeDefaultRows(tab.type).slice(0, 10);
            tab.sectionsRows[sec.id] = rows.concat(extra);
            saveState(); render();
          } }, "+10행")
        )
      )
    );

    const wrap = el("div", { class: "table-wrap" });
    const table = el("table", {});
    const thead = el("thead", {}, el("tr", {},
      el("th", {}, "No"),
      el("th", {}, "코드"),
      el("th", {}, "품명(자동)"),
      el("th", {}, "규격(자동)"),
      el("th", {}, "단위(자동)"),
      el("th", {}, "산출식"),
      el("th", {}, "물량(Value)"),
      el("th", {}, "할증(배수)"),
      el("th", {}, "환산단위")
    ));
    const tbody = el("tbody");

    rows.forEach((r, i) => {
      r.value = safeEvalMath(r.expr, varMap);
      tbody.appendChild(el("tr", {},
        el("td", {}, String(i + 1)),
        el("td", {}, el("input", { class:"cell", type:"text", value:r.code ?? "", placeholder:"코드 입력",
          oninput:(e)=>{ r.code=e.target.value; saveState(); }
        })),
        el("td", {}, el("input", { class:"cell readonly", type:"text", value:r.name ?? "", readonly:true })),
        el("td", {}, el("input", { class:"cell readonly", type:"text", value:r.spec ?? "", readonly:true })),
        el("td", {}, el("input", { class:"cell", type:"text", value:r.unit ?? "",
          oninput:(e)=>{ r.unit=e.target.value; saveState(); }
        })),
        el("td", {}, el("input", { class:"cell", type:"text", value:r.expr ?? "", placeholder:"예: (A+0.5)*2  (<...>는 주석)",
          oninput:(e)=>{ r.expr=e.target.value; saveState(); },
          onkeydown:(e)=>{ if(e.key==="Enter"){ e.preventDefault(); saveState(); render(); requestAnimationFrame(()=>e.currentTarget.focus()); } }
        })),
        el("td", {}, el("input", { class:"cell readonly", type:"text", readonly:true, value:formatNumber(r.value) })),
        el("td", {}, el("input", { class:"cell", type:"text", value:r.mult ?? "",
          oninput:(e)=>{ r.mult=e.target.value; saveState(); }
        })),
        el("td", {}, el("input", { class:"cell", type:"text", value:r.conv ?? "",
          oninput:(e)=>{ r.conv=e.target.value; saveState(); }
        }))
      ));
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    panel.appendChild(wrap);
    return panel;
  }

  function renderCodeTab() {
    return el("div", { class:"panel" },
      el("div", { class:"panel-header" },
        el("div", {},
          el("div", { class:"panel-title" }, "코드(Ctrl+.)"),
          el("div", { class:"panel-desc" }, "코드 선택 팝업/엑셀 연동은 기존 로직 연결 지점")
        )
      )
    );
  }

  function renderSummaryTab(tab) {
    return el("div", { class:"panel" },
      el("div", { class:"panel-header" },
        el("div", {},
          el("div", { class:"panel-title" }, tab.label),
          el("div", { class:"panel-desc" }, "집계 탭(확장 예정)")
        )
      )
    );
  }

  function renderView() {
    viewEl.innerHTML = "";
    const tab = getActiveTab();

    if (tab.type.startsWith("calc_")) {
      viewEl.appendChild(renderTopSplit(tab));
      viewEl.appendChild(renderCalcTab(tab));
      return;
    }
    if (tab.type === "code") { viewEl.appendChild(renderCodeTab()); return; }
    if (tab.type.startsWith("summary_")) { viewEl.appendChild(renderSummaryTab(tab)); return; }

    viewEl.appendChild(el("div", { class:"panel" }, "준비중"));
  }

  function render() {
    renderTabs();
    renderView();
    updateStickyVars();
  }

  /* Global hotkeys */
  document.addEventListener("keydown", (e) => {
    const t = e.target;
    const tag = t?.tagName ? t.tagName.toLowerCase() : "";
    const isTyping = (tag === "input" || tag === "textarea" || t?.isContentEditable);

    if (e.ctrlKey && e.key === "F3") {
      e.preventDefault();
      const tab = getActiveTab();
      if (tab.type.startsWith("calc_")) addSection(tab);
      return;
    }

    // 입력중이면 구분 이동 가로채지 않음
    if (isTyping) return;

    if (e.key === "ArrowUp" || e.key === "ArrowDown") {
      const tab = getActiveTab();
      if (!tab.type.startsWith("calc_")) return;

      const cur = getActiveSection(tab);
      const curIdx = tab.sections.findIndex(s => s.id === cur?.id);
      const nextIdx = clamp(curIdx + (e.key === "ArrowDown" ? 1 : -1), 0, tab.sections.length - 1);
      const nextSec = tab.sections[nextIdx];
      if (!nextSec) return;

      e.preventDefault();
      tab.activeSectionId = nextSec.id;
      if (!tab.sectionsRows[nextSec.id]) tab.sectionsRows[nextSec.id] = makeDefaultRows(tab.type);
      saveState(); render();
      requestAnimationFrame(() => $(`.section-item[data-sec-id="${cssEscape(nextSec.id)}"]`)?.focus());
    }
  }, { capture:true });

  /* Buttons */
  $("#btnExport")?.addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type:"application/json" });
    const url = URL.createObjectURL(blob);
    const a = el("a", { href:url, download:"FIN_WEB_export.json" });
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  });

  $("#fileImport")?.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const obj = JSON.parse(text);
      if (!obj?.tabs) throw new Error("Invalid");
      state = obj;
      saveState(); render();
    } catch {
      alert("가져오기 실패: JSON 형식을 확인해주세요.");
    } finally {
      e.target.value = "";
    }
  });

  $("#btnReset")?.addEventListener("click", () => {
    if (!confirm("초기화 하시겠습니까? (저장된 데이터가 삭제됩니다)")) return;
    state = makeDefaultState();
    saveState(); render();
  });

  $("#btnOpenPicker")?.addEventListener("click", () => {
    alert("코드 선택 창(Ctrl+.) 로직은 기존 구현을 app.js에 연결해 주세요.");
  });

  window.addEventListener("resize", updateStickyVars);
  window.addEventListener("load", updateStickyVars);

  render();
})();
