/* app.js (FINAL) - FIN 산출자료 (Web)
   - 코드탭 복원(코드/품명/규격/단위/할증/환산단위/환산계수/비고)
   - 엑셀 기반 코드마스터 기본값 내장
   - 산출탭 방향키 이동: 산출표(빨간영역) 안에서만 동작
   - Ctrl+F3: 산출표에서만 현재 행 아래 행 추가  ✅ + 코드탭도 동일 동작
   - Ctrl+Shift+F3: 산출표 +10행               ✅ + 코드탭도 동일 동작
   - ✅ Ctrl+Del: "줄(행) 삭제" → 줄 수 감소 (calc/var/code 전부 적용)
   - Ctrl+. : 코드 선택 창
   - ✅ 상단(topbar/tabs/top-split) 실제 높이 측정 → sticky offset 자동 보정
   - ✅ 노란 영역(패널 헤더) sticky 고정
*/

(() => {
  "use strict";

  /***************
   * Storage
   ***************/
  const LS_KEY = "FIN_WEB_STATE_V11";
  const deepClone = (obj) => JSON.parse(JSON.stringify(obj));
  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  /***************
   * Sticky height auto-measure
   ***************/
  function updateStickyVars() {
    const root = document.documentElement;

    const topbar = document.querySelector(".topbar");
    const tabs = document.querySelector(".tabs");
    const topSplit = document.querySelector(".top-split"); // 산출탭에서만 존재

    const topbarH = topbar ? topbar.getBoundingClientRect().height : 0;
    const tabsH = tabs ? tabs.getBoundingClientRect().height : 0;
    const topSplitH = topSplit ? topSplit.getBoundingClientRect().height : 0;

    root.style.setProperty("--topbarH", `${Math.ceil(topbarH)}px`);
    root.style.setProperty("--tabsH", `${Math.ceil(tabsH)}px`);
    root.style.setProperty("--topSplitActualH", `${Math.ceil(topSplitH)}px`);

    // 기본(코드탭/집계탭): topbar+tabs 아래에 패널헤더 고정
    const base = Math.ceil(topbarH + tabsH);
    root.style.setProperty("--stickyBaseTop", `${base}px`);

    // 산출탭: topbar+tabs+topSplit 아래에 패널헤더 고정
    const withTopSplit = Math.ceil(topbarH + tabsH + topSplitH + 10); // 10px 여유
    root.style.setProperty("--stickyWithTopSplitTop", `${withTopSplit}px`);
  }

  // resize/폰 회전/버튼 줄바꿈 대응
  window.addEventListener("resize", () => {
    requestAnimationFrame(updateStickyVars);
  });

  /***************
   * Code Master (엑셀 동일 구조)
   ***************/
  const DEFAULT_CODE_MASTER = [
    {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"convUnit":"TON","convFactor":0.0315,"note":""},
    {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"convUnit":"TON","convFactor":0.0213,"note":""},
    {"code":"A0SM355201","name":"RH형강 / SM355","spec":"200*200*8*12","unit":"M","surcharge":null,"convUnit":"","convFactor":null,"note":""},
    {"code":"A0SM355294","name":"RH형강 / SM355","spec":"294*200*8*12","unit":"M","surcharge":null,"convUnit":"","convFactor":null,"note":""},
    {"code":"A0SM355300","name":"RH형강 / SM355","spec":"300*300*10*15, CAMBER 35mm","unit":"M","surcharge":null,"convUnit":"","convFactor":null,"note":""},

    {"code":"B0SM355800","name":"BH형강 / SM355","spec":"800*300*25*40","unit":"M","surcharge":10,"convUnit":"TON","convFactor":0.3297,"note":""},
    {"code":"B0SM355900","name":"BH형강 / SM355","spec":"900*350*30*60","unit":"M","surcharge":10,"convUnit":"TON","convFactor":0.35796,"note":""},

    {"code":"C0SS275009","name":"강판 / SS275","spec":"9mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275010","name":"강판 / SS275","spec":"10mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275011","name":"강판 / SS275","spec":"11mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275012","name":"강판 / SS275","spec":"12mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275013","name":"강판 / SS275","spec":"13mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275014","name":"강판 / SS275","spec":"14mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SS275025","name":"강판 / SS275","spec":"25mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355009","name":"강판 / SM355","spec":"9mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355010","name":"강판 / SM355","spec":"10mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355011","name":"강판 / SM355","spec":"11mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355012","name":"강판 / SM355","spec":"12mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355013","name":"강판 / SM355","spec":"13mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355014","name":"강판 / SM355","spec":"14mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
    {"code":"C0SM355025","name":"강판 / SM355","spec":"25mm","unit":"M2","surcharge":null,"convUnit":"","convFactor":null,"note":"Plate / Đĩa"},
  ];

  /***************
   * Tabs
   ***************/
  const TABS = [
    { id: "code", title: "코드(Ctrl+.)" },
    { id: "steel", title: "철골" },
    { id: "steel_sum", title: "철골_집계" },
    { id: "steel_sub", title: "철골_부자재" },
    { id: "support", title: "구조이기/동바리" },
    { id: "support_sum", title: "구조이기/동바리_집계" },
  ];

  /***************
   * Default State
   ***************/
  const defaultCalcRow = () => ({
    code: "",
    name: "",
    spec: "",
    unit: "",
    formula: "",
    value: 0,
    surchargePct: null,
    surchargeMul: 1,
    convUnit: "",
    convFactor: null,
    converted: 0,
    note: "",
  });

  const defaultVarRow = () => ({
    key: "",
    expr: "",
    value: 0,
    note: "",
  });

  const defaultSection = (name = "구분 1", count = 1) => ({
    name,
    count,
    vars: Array.from({ length: 12 }, () => defaultVarRow()),
    rows: Array.from({ length: 12 }, () => defaultCalcRow()),
  });

  const DEFAULT_STATE = {
    activeTab: "code",
    codeMaster: deepClone(DEFAULT_CODE_MASTER),
    steel: { activeSection: 0, sections: [defaultSection("구분 1", 1)] },
    steel_sub: { activeSection: 0, sections: [defaultSection("구분 1", 1)] },
    support: { activeSection: 0, sections: [defaultSection("구분 1", 1)] },
  };

  function loadState() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return deepClone(DEFAULT_STATE);
      const parsed = JSON.parse(raw);

      const s = { ...deepClone(DEFAULT_STATE), ...parsed };
      s.codeMaster = Array.isArray(parsed?.codeMaster) ? parsed.codeMaster : deepClone(DEFAULT_CODE_MASTER);

      for (const k of ["steel", "steel_sub", "support"]) {
        if (!s[k] || !Array.isArray(s[k].sections) || s[k].sections.length === 0) {
          s[k] = deepClone(DEFAULT_STATE[k]);
        }
        s[k].activeSection = clamp(Number(s[k].activeSection || 0), 0, s[k].sections.length - 1);
      }

      if (!TABS.some(t => t.id === s.activeTab)) s.activeTab = "code";
      return s;
    } catch (e) {
      console.warn("loadState failed:", e);
      return deepClone(DEFAULT_STATE);
    }
  }

  function saveState() {
    localStorage.setItem(LS_KEY, JSON.stringify(state));
  }

  let state = loadState();

  /***************
   * DOM
   ***************/
  const $tabs = document.getElementById("tabs");
  const $view = document.getElementById("view");

  function el(tag, attrs = {}, children = []) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "dataset") Object.assign(node.dataset, v);
      else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
      else if (v === false || v == null) continue;
      else node.setAttribute(k, String(v));
    }
    for (const ch of children) {
      if (ch == null) continue;
      node.appendChild(typeof ch === "string" ? document.createTextNode(ch) : ch);
    }
    return node;
  }

  function clear(node) {
    while (node.firstChild) node.removeChild(node.firstChild);
  }

  /***************
   * Helpers: Code master lookup
   ***************/
  function codeLookup(code) {
    const c = String(code || "").trim();
    if (!c) return null;
    return state.codeMaster.find(x => String(x.code).trim().toUpperCase() === c.toUpperCase()) || null;
  }

  /***************
   * Expression evaluator
   ***************/
  function stripAngleComments(expr) {
    if (!expr) return "";
    return String(expr).replace(/<[^>]*>/g, "");
  }

  function buildVarMap(section) {
    const map = Object.create(null);

    for (const v of section.vars) {
      const key = (v.key || "").trim();
      if (!key) continue;
      map[key.toUpperCase()] = 0;
    }

    for (let pass = 0; pass < 6; pass++) {
      for (const v of section.vars) {
        const key = (v.key || "").trim();
        if (!key) continue;

        const exprRaw = stripAngleComments(v.expr || "");
        const val = safeEvalWithVars(exprRaw, map);
        if (Number.isFinite(val)) map[key.toUpperCase()] = val;
      }
    }

    for (const v of section.vars) {
      const key = (v.key || "").trim();
      if (!key) v.value = 0;
      else v.value = Number(map[key.toUpperCase()] ?? 0) || 0;
    }
    return map;
  }

  function safeEvalWithVars(expr, varMap) {
    const raw = String(expr || "").trim();
    if (!raw) return 0;

    const replaced = raw.replace(/\b([A-Za-z][A-Za-z0-9]{0,2})\b/g, (m, p1) => {
      const k = p1.toUpperCase();
      if (Object.prototype.hasOwnProperty.call(varMap, k)) return String(varMap[k] ?? 0);
      return "0";
    });

    const cleaned = replaced.replace(/\s+/g, "");
    if (!/^[0-9+\-*/().]*$/.test(cleaned)) return NaN;

    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(`return (${replaced});`);
      const v = fn();
      const n = Number(v);
      return Number.isFinite(n) ? n : NaN;
    } catch {
      return NaN;
    }
  }

  function recomputeSection(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];
    const varMap = buildVarMap(sec);

    for (const r of sec.rows) {
      const info = codeLookup(r.code);
      if (info) {
        r.name = info.name || "";
        r.spec = info.spec || "";
        r.unit = info.unit || "";
        r.note = info.note || "";
        r.convUnit = info.convUnit || "";
        r.convFactor = info.convFactor ?? null;

        const sPct = (r.surchargePct == null || r.surchargePct === "") ? (info.surcharge ?? null) : Number(r.surchargePct);
        r.surchargePct = (sPct == null || sPct === "") ? null : Number(sPct);
      } else {
        r.name = r.name || "";
        r.spec = r.spec || "";
        r.unit = r.unit || "";
        r.note = r.note || "";
        r.convUnit = r.convUnit || "";
      }

      const expr = stripAngleComments(r.formula || "");
      const base = safeEvalWithVars(expr, varMap);
      r.value = Number.isFinite(base) ? base : 0;

      const pct = (r.surchargePct == null || r.surchargePct === "") ? null : Number(r.surchargePct);
      const mul = pct == null || !Number.isFinite(pct) ? 1 : (1 + pct / 100);
      r.surchargeMul = mul;

      const after = r.value * mul;
      const cf = r.convFactor;
      if (cf != null && Number.isFinite(Number(cf)) && Number(cf) !== 0) r.converted = after * Number(cf);
      else r.converted = after;
    }
  }

  /***************
   * UI: Tabs
   ***************/
  function renderTabs() {
    clear($tabs);
    for (const t of TABS) {
      const btn = el("button", {
        class: "tab" + (state.activeTab === t.id ? " active" : ""),
        onclick: () => {
          state.activeTab = t.id;
          saveState();
          render();
        }
      }, [t.title]);
      $tabs.appendChild(btn);
    }
  }

  /***************
   * ✅ Code tab (방향키/행추가 지원)
   ***************/
  function renderCodeTab() {
    const wrap = el("div", { class: "table-wrap" }, [buildCodeMasterTable()]);
    attachGridNav(wrap);

    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [
        el("div", { class: "panel-title" }, ["코드"]),
        el("div", { class: "panel-desc" }, [
          "방향키: 코드표 셀 이동 | Ctrl+F3 행추가 | Shift+Ctrl+F3 +10행 | Ctrl+. 코드선택(산출표에서) | Ctrl+Del 줄삭제"
        ])
      ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => addCodeRows(1) }, ["행 추가 (Ctrl+F3)"]),
        el("button", { class: "smallbtn", onclick: () => addCodeRows(10) }, ["+10행"]),
      ])
    ]);

    return el("div", { class: "panel" }, [panelHeader, wrap]);
  }

  function buildCodeMasterTable() {
    const table = el("table", {}, []);
    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["코드(Code)"]),
        el("th", {}, ["품명(Product name)"]),
        el("th", {}, ["규격(Specifications)"]),
        el("th", {}, ["단위(unit)"]),
        el("th", {}, ["할증(surcharge)"]),
        el("th", {}, ["환산단위(Conversion unit)"]),
        el("th", {}, ["환산계수(Conversion factor)"]),
        el("th", {}, ["비고(Note)"]),
        el("th", {}, [""])
      ])
    ]);
    const tbody = el("tbody", {}, []);

    state.codeMaster.forEach((row, idx) => {
      const tr = el("tr", {}, [
        tdInput("codeMaster", idx, "code", row.code),
        tdInput("codeMaster", idx, "name", row.name),
        tdInput("codeMaster", idx, "spec", row.spec),
        tdInput("codeMaster", idx, "unit", row.unit),
        tdInput("codeMaster", idx, "surcharge", row.surcharge ?? ""),
        tdInput("codeMaster", idx, "convUnit", row.convUnit),
        tdInput("codeMaster", idx, "convFactor", row.convFactor ?? ""),
        tdInput("codeMaster", idx, "note", row.note),
        el("td", {}, [
          el("button", {
            class: "smallbtn",
            onclick: () => {
              state.codeMaster.splice(idx, 1);
              saveState(); render();
            }
          }, ["삭제"])
        ])
      ]);
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    return table;
  }

  const CODE_COL_INDEX = {
    code: 0, name: 1, spec: 2, unit: 3, surcharge: 4, convUnit: 5, convFactor: 6, note: 7
  };

  function tdInput(scope, rowIndex, field, value, opts = {}) {
    const ds =
      scope === "codeMaster"
        ? { grid: "code", row: String(rowIndex), col: String(CODE_COL_INDEX[field] ?? 0), field }
        : (opts.dataset || null);

    const input = el("input", {
      class: "cell" + (opts.readonly ? " readonly" : ""),
      value: value ?? "",
      readonly: opts.readonly ? "readonly" : null,
      dataset: ds,
      oninput: (e) => {
        const v = e.target.value;
        if (scope === "codeMaster") {
          const r = state.codeMaster[rowIndex];
          if (!r) return;
          if (field === "surcharge" || field === "convFactor") r[field] = v === "" ? null : Number(v);
          else r[field] = v;
          saveState();
        }
      }
    });

    return el("td", {}, [input]);
  }

  function addCodeRows(n, insertAfterRow = null) {
    const idx = insertAfterRow == null ? (state.codeMaster.length - 1) : insertAfterRow;
    const insertPos = clamp(idx + 1, 0, state.codeMaster.length);

    const empty = { code:"", name:"", spec:"", unit:"", surcharge:null, convUnit:"", convFactor:null, note:"" };
    const newRows = Array.from({ length: n }, () => deepClone(empty));

    state.codeMaster.splice(insertPos, 0, ...newRows);
    saveState();
    render();

    requestAnimationFrame(() => {
      const first = document.querySelector(`input[data-grid="code"][data-row="${insertPos}"][data-col="0"]`);
      if (first) first.focus();
    });
  }

  /***************
   * ✅ Ctrl+Del 줄(행) 삭제 helpers (calc/var/code)
   ***************/
  function deleteCalcRow(tabId, rowIdx) {
    const bucket = state[tabId];
    if (!bucket) return;
    const sec = bucket.sections[bucket.activeSection];
    if (!sec?.rows) return;
    if (rowIdx < 0 || rowIdx >= sec.rows.length) return;

    // 최소 1행 유지
    if (sec.rows.length <= 1) {
      sec.rows[0] = defaultCalcRow();
    } else {
      sec.rows.splice(rowIdx, 1);
    }

    recomputeSection(tabId);
    saveState();
    render();

    requestAnimationFrame(() => {
      const newRow = Math.min(rowIdx, sec.rows.length - 1);
      const t = document.querySelector(
        `input[data-grid="calc"][data-tab="${tabId}"][data-row="${newRow}"][data-col="0"]`
      );
      if (t) t.focus();
    });
  }

  function deleteVarRow(tabId, rowIdx) {
    const bucket = state[tabId];
    if (!bucket) return;
    const sec = bucket.sections[bucket.activeSection];
    if (!sec?.vars) return;
    if (rowIdx < 0 || rowIdx >= sec.vars.length) return;

    // 최소 1행 유지
    if (sec.vars.length <= 1) {
      sec.vars[0] = defaultVarRow();
    } else {
      sec.vars.splice(rowIdx, 1);
    }

    recomputeSection(tabId);
    saveState();
    render();

    requestAnimationFrame(() => {
      const newRow = Math.min(rowIdx, sec.vars.length - 1);
      const t = document.querySelector(
        `input[data-grid="var"][data-tab="${tabId}"][data-row="${newRow}"][data-col="0"]`
      );
      if (t) t.focus();
    });
  }

  function deleteCodeRow(rowIdx) {
    if (!Array.isArray(state.codeMaster)) return;
    if (rowIdx < 0 || rowIdx >= state.codeMaster.length) return;

    // 최소 1행 유지
    if (state.codeMaster.length <= 1) {
      state.codeMaster[0] = { code:"", name:"", spec:"", unit:"", surcharge:null, convUnit:"", convFactor:null, note:"" };
    } else {
      state.codeMaster.splice(rowIdx, 1);
    }

    saveState();
    render();

    requestAnimationFrame(() => {
      const newRow = Math.min(rowIdx, state.codeMaster.length - 1);
      const t = document.querySelector(`input[data-grid="code"][data-row="${newRow}"][data-col="0"]`);
      if (t) t.focus();
    });
  }

  /***************
   * UI: Section + Vars + Calc
   ***************/
  function renderCalcTab(tabId, title) {
    recomputeSection(tabId);

    const top = el("div", { class: "top-split" }, [
      el("div", { class: "calc-layout top-grid" }, [
        el("div", { class: "rail-box section-box", dataset: { region: "section" } }, [
          el("div", { class: "rail-title" }, ["구분명 리스트 (↑/↓ 이동)"]),
          buildSectionList(tabId),
          buildSectionEditor(tabId),
        ]),
        el("div", { class: "rail-box var-box", dataset: { region: "var" } }, [
          el("div", { class: "rail-title" }, ["변수표 (A, AB, A1, AB1... 최대 3자)"]),
          buildVarTable(tabId),
        ]),
      ])
    ]);

    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [
        el("div", { class: "panel-title" }, [title]),
        el("div", { class: "panel-desc" }, [
          "방향키: 산출표 셀 이동 | 산출식 Enter 계산 | Ctrl+. 코드선택 | Ctrl+F3 행추가 | Shift+Ctrl+F3 +10행 | Ctrl+Del 줄삭제"
        ])
      ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => addRows(tabId, 1) }, ["행 추가 (Ctrl+F3)"]),
        el("button", { class: "smallbtn", onclick: () => addRows(tabId, 10) }, ["+10행"]),
      ])
    ]);

    const panel = el("div", { class: "panel" }, [
      panelHeader,
      el("div", { class: "table-wrap" }, [buildCalcTable(tabId)])
    ]);

    return el("div", {}, [top, panel]);
  }

  function buildSectionList(tabId) {
    const bucket = state[tabId];
    const list = el("div", { class: "section-list", dataset: { nav: "sectionList" } }, []);

    bucket.sections.forEach((s, idx) => {
      const item = el("div", {
        class: "section-item" + (bucket.activeSection === idx ? " active" : ""),
        tabindex: "0",
        onclick: () => {
          bucket.activeSection = idx;
          saveState();
          render();
        },
      }, [
        el("div", { class: "name" }, [s.name || `구분 ${idx + 1}`]),
        el("div", { class: "meta-inline" }, [`개소: ${s.count ?? ""}`]),
        el("div", { class: "meta" }, ["선택"])
      ]);
      list.appendChild(item);
    });

    list.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
      e.preventDefault();

      const dir = e.key === "ArrowDown" ? 1 : -1;
      bucket.activeSection = clamp(bucket.activeSection + dir, 0, bucket.sections.length - 1);
      saveState();
      render();

      requestAnimationFrame(() => {
        const newList = document.querySelector(".section-list");
        const items = newList ? [...newList.querySelectorAll(".section-item")] : [];
        if (items[bucket.activeSection]) items[bucket.activeSection].focus();
      });
    });

    return list;
  }

  function buildSectionEditor(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const nameInput = el("input", {
      class: "cell",
      value: sec.name || "",
      placeholder: "구분명 (예: 2층 바닥 철골보)",
      oninput: (e) => {
        sec.name = e.target.value;
        saveState();
        const item = document.querySelectorAll(".section-item .name")[bucket.activeSection];
        if (item) item.textContent = sec.name || `구분 ${bucket.activeSection + 1}`;
      }
    });

    const countInput = el("input", {
      class: "cell",
      value: sec.count ?? "",
      placeholder: "개소(예: 0,1,2...)",
      oninput: (e) => {
        const v = e.target.value.trim();
        sec.count = v === "" ? "" : Number(v);
        saveState();
        const meta = document.querySelectorAll(".section-item .meta-inline")[bucket.activeSection];
        if (meta) meta.textContent = `개소: ${sec.count ?? ""}`;
      }
    });

    const saveBtn = el("button", { class: "smallbtn", onclick: () => { saveState(); render(); } }, ["저장"]);
    const addBtn = el("button", {
      class: "smallbtn",
      onclick: () => {
        bucket.sections.push(defaultSection(`구분 ${bucket.sections.length + 1}`, 1));
        bucket.activeSection = bucket.sections.length - 1;
        saveState(); render();
      }
    }, ["구분 추가"]);
    const delBtn = el("button", {
      class: "smallbtn",
      onclick: () => {
        if (bucket.sections.length <= 1) return alert("구분은 최소 1개가 필요합니다.");
        bucket.sections.splice(bucket.activeSection, 1);
        bucket.activeSection = clamp(bucket.activeSection, 0, bucket.sections.length - 1);
        saveState(); render();
      }
    }, ["구분 삭제"]);

    return el("div", { class: "section-editor" }, [nameInput, countInput, saveBtn, addBtn, delBtn]);
  }

  function buildVarTable(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const wrap = el("div", { class: "var-tablewrap" }, []);
    const table = el("table", { class: "var-table" }, []);
    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["변수"]),
        el("th", {}, ["산식"]),
        el("th", {}, ["값"]),
        el("th", {}, ["비고"])
      ])
    ]);
    const tbody = el("tbody", {}, []);

    sec.vars.forEach((v, r) => {
      const tr = el("tr", {}, [
        tdNavInputVar(tabId, r, 0, "key", v.key, { placeholder: "예: A / AB / A1" }),
        tdNavInputVar(tabId, r, 1, "expr", v.expr, { placeholder: "예: (A+0.5)*2  (<...> 주석)" }),
        tdNavInputVar(tabId, r, 2, "value", String(v.value ?? 0), { readonly: true }),
        tdNavInputVar(tabId, r, 3, "note", v.note, { placeholder: "비고" }),
      ]);
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    wrap.appendChild(table);

    wrap.addEventListener("input", () => {
      recomputeSection(tabId);
      saveState();
      const valueInputs = wrap.querySelectorAll('input[data-grid="var"][data-col="2"]');
      sec.vars.forEach((vv, i) => {
        if (valueInputs[i]) valueInputs[i].value = String(vv.value ?? 0);
      });
      refreshCalcComputed(tabId);
    });

    attachGridNav(wrap);
    return wrap;
  }

  function tdNavInputVar(tabId, row, col, field, value, opts = {}) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const input = el("input", {
      class: "cell" + (opts.readonly ? " readonly" : ""),
      value: value ?? "",
      placeholder: opts.placeholder || "",
      readonly: opts.readonly ? "readonly" : null,
      dataset: { grid: "var", tab: tabId, row: String(row), col: String(col), field },
      oninput: (e) => {
        if (opts.readonly) return;
        const rr = sec.vars[row];
        if (!rr) return;

        if (field === "key") {
          let val = e.target.value.toUpperCase();
          val = val.replace(/[^A-Z0-9]/g, "");
          if (val.length > 3) val = val.slice(0, 3);
          if (val && !/^[A-Z]/.test(val)) val = val.replace(/^[^A-Z]+/, "");
          e.target.value = val;
          rr.key = val;
        } else {
          rr[field] = e.target.value;
        }
      }
    });
    return el("td", {}, [input]);
  }

  function buildCalcTable(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const table = el("table", {}, []);
    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["No"]),
        el("th", {}, ["코드"]),
        el("th", {}, ["품명(자동)"]),
        el("th", {}, ["규격(자동)"]),
        el("th", {}, ["단위(자동)"]),
        el("th", {}, ["산출식"]),
        el("th", {}, ["물량(Value)"]),
        el("th", {}, ["할증(%)"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["환산후수량"]),
        el("th", {}, ["비고"]),
      ])
    ]);

    const tbody = el("tbody", {}, []);
    sec.rows.forEach((r, i) => {
      const tr = el("tr", {}, [
        el("td", {}, [String(i + 1)]),
        tdNavInputCalc(tabId, i, 0, "code", r.code, { placeholder: "코드 입력" }),
        tdNavInputCalc(tabId, i, 1, "name", r.name, { readonly: true }),
        tdNavInputCalc(tabId, i, 2, "spec", r.spec, { readonly: true }),
        tdNavInputCalc(tabId, i, 3, "unit", r.unit, { readonly: true }),
        tdNavInputCalc(tabId, i, 4, "formula", r.formula, { placeholder: "예: (A+0.5)*2  (<...> 주석)" }),
        tdNavInputCalc(tabId, i, 5, "value", String(r.value ?? 0), { readonly: true }),
        tdNavInputCalc(tabId, i, 6, "surchargePct", r.surchargePct ?? "", { placeholder: "자동/직접입력" }),
        tdNavInputCalc(tabId, i, 7, "convUnit", r.convUnit || "", { readonly: true }),
        tdNavInputCalc(tabId, i, 8, "convFactor", r.convFactor ?? "", { readonly: true }),
        tdNavInputCalc(tabId, i, 9, "converted", String(r.converted ?? 0), { readonly: true }),
        tdNavInputCalc(tabId, i, 10, "note", r.note || "", { readonly: true }),
      ]);
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    const wrap = el("div", {}, [table]);
    attachGridNav(wrap);

    wrap.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (t.dataset.grid !== "calc") return;
      if (e.key === "Enter") {
        e.preventDefault();
        recomputeSection(tabId);
        saveState();
        refreshCalcComputed(tabId);
      }
    });

    return wrap;
  }

  function tdNavInputCalc(tabId, row, col, field, value, opts = {}) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const input = el("input", {
      class: "cell" + (opts.readonly ? " readonly" : ""),
      value: value ?? "",
      placeholder: opts.placeholder || "",
      readonly: opts.readonly ? "readonly" : null,
      dataset: { grid: "calc", tab: tabId, row: String(row), col: String(col), field },
      oninput: (e) => {
        if (opts.readonly) return;

        const rr = sec.rows[row];
        if (!rr) return;

        if (field === "code") {
          rr.code = e.target.value.toUpperCase().trim();
          recomputeSection(tabId);
          saveState();
          refreshCalcComputed(tabId);
        } else if (field === "surchargePct") {
          const v = e.target.value.trim();
          rr.surchargePct = v === "" ? null : Number(v);
          recomputeSection(tabId);
          saveState();
          refreshCalcComputed(tabId);
        } else {
          rr[field] = e.target.value;
        }
      }
    });

    return el("td", {}, [input]);
  }

  function refreshCalcComputed(tabId) {
    const wrap = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"]`)?.closest(".table-wrap") || document.body;
    const inputs = wrap.querySelectorAll(`input[data-grid="calc"][data-tab="${tabId}"]`);
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    inputs.forEach((inp) => {
      const r = Number(inp.dataset.row);
      const f = inp.dataset.field;
      const row = sec.rows[r];
      if (!row) return;

      if (["name", "spec", "unit", "value", "convUnit", "convFactor", "converted", "note"].includes(f)) {
        inp.value = (row[f] ?? "") + "";
      }
    });
  }

  /***************
   * ✅ Grid navigation (방향키: 표 내부에서만)
   ***************/
  function attachGridNav(container) {
    container.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;

      const grid = t.dataset.grid;
      if (grid !== "calc" && grid !== "var" && grid !== "code") return;

      const key = e.key;
      if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(key)) return;

      e.preventDefault();

      const row = Number(t.dataset.row);
      const col = Number(t.dataset.col);
      let nr = row, nc = col;

      if (key === "ArrowUp") nr = row - 1;
      if (key === "ArrowDown") nr = row + 1;
      if (key === "ArrowLeft") nc = col - 1;
      if (key === "ArrowRight") nc = col + 1;

      const selector = `input[data-grid="${grid}"][data-row="${nr}"][data-col="${nc}"]`;
      const next = container.querySelector(selector);

      if (next) next.focus();
    });
  }

  /***************
   * Row add (Ctrl+F3 / +10)
   ***************/
  function addRows(tabId, n, insertAfterRow = null) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const idx = (insertAfterRow == null) ? (sec.rows.length - 1) : insertAfterRow;
    const insertPos = clamp(idx + 1, 0, sec.rows.length);

    const newRows = Array.from({ length: n }, () => defaultCalcRow());
    sec.rows.splice(insertPos, 0, ...newRows);

    saveState();
    render();

    requestAnimationFrame(() => {
      const first = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="0"]`);
      if (first) first.focus();
    });
  }

  /***************
   * Shortcuts
   ***************/
  window.addEventListener("keydown", (e) => {
    // Ctrl+.
    if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key === ".") {
      e.preventDefault();
      e.stopPropagation();
      openCodePicker();
      return;
    }

    // ✅ Ctrl+Del : 다양한 브라우저/키보드 대응 (key + code + keyCode)
    const isCtrlDel =
      e.ctrlKey &&
      !e.shiftKey &&
      !e.altKey &&
      (
        e.key === "Delete" || e.key === "Del" || e.key === "Backspace" ||
        e.code === "Delete" || e.code === "Backspace" ||
        e.keyCode === 46 || e.keyCode === 8 // 46: Delete, 8: Backspace
      );

    // ✅ Ctrl+Del = 줄(행) 삭제 (줄 수 감소)
    if (isCtrlDel) {
      const a = document.activeElement;

      const isEditableEl =
        (a instanceof HTMLInputElement) || (a instanceof HTMLTextAreaElement);
      if (!isEditableEl) return;

      const grid = a.dataset?.grid;
      if (grid !== "calc" && grid !== "var" && grid !== "code") return;

      e.preventDefault();
      e.stopPropagation();

      if (grid === "calc") {
        const tabId = a.dataset.tab;
        const row = Number(a.dataset.row);
        deleteCalcRow(tabId, row);
        return;
      }

      if (grid === "var") {
        const tabId = a.dataset.tab;
        const row = Number(a.dataset.row);
        deleteVarRow(tabId, row);
        return;
      }

      if (grid === "code") {
        const row = Number(a.dataset.row);
        deleteCodeRow(row);
        return;
      }
    }

    // Ctrl+F3 / Ctrl+Shift+F3
    if (e.ctrlKey && (e.key === "F3")) {
      const a = document.activeElement;
      if (a instanceof HTMLInputElement) {
        const grid = a.dataset.grid;

        if (grid === "calc") {
          e.preventDefault();
          e.stopPropagation();
          const tabId = a.dataset.tab;
          const row = Number(a.dataset.row);
          if (e.shiftKey) addRows(tabId, 10, row);
          else addRows(tabId, 1, row);
          return;
        }

        if (grid === "code") {
          e.preventDefault();
          e.stopPropagation();
          const row = Number(a.dataset.row);
          if (e.shiftKey) addCodeRows(10, row);
          else addCodeRows(1, row);
          return;
        }
      }
    }
  }, { capture: true }); // ✅ 이 옵션까지 포함해서 교체!

  /***************
   * Code Picker Popup
   ***************/
  let __pickerWin = null;

  function openCodePicker() {
    let originTab = state.activeTab || "steel";
    let focusRow = 0;

    const a = document.activeElement;
    if (a instanceof HTMLInputElement && a.dataset.grid === "calc") {
      originTab = a.dataset.tab || originTab;
      focusRow = Number(a.dataset.row || 0);
    }

    const codesForPicker = (state.codeMaster || []).map(r => ({
      code: (r.code ?? "").toString(),
      name: (r.name ?? "").toString(),
      spec: (r.spec ?? "").toString(),
      unit: (r.unit ?? "").toString(),
      surcharge: (r.surcharge ?? "").toString(),
      conv_unit: (r.convUnit ?? "").toString(),
      conv_factor: (r.convFactor ?? "").toString(),
      note: (r.note ?? "").toString(),
    }));

    const url = "picker.html";

    __pickerWin = window.open(url, "FIN_CODE_PICKER", "width=1100,height=760");
    if (!__pickerWin) {
      alert("팝업이 차단되었습니다. 브라우저에서 팝업 허용 후 다시 시도해 주세요.");
      return;
    }

    let tries = 0;
    const timer = setInterval(() => {
      tries++;
      try {
        __pickerWin.postMessage(
          { type: "INIT", originTab, focusRow, codes: codesForPicker },
          window.location.origin
        );
      } catch {}
      if (tries >= 12) clearInterval(timer);
    }, 120);
  }

  window.addEventListener("message", (event) => {
    if (event.origin !== window.location.origin) return;
    const msg = event.data;
    if (!msg || typeof msg !== "object") return;

    if (msg.type === "INSERT_SELECTED") {
      const originTab = msg.originTab || state.activeTab;
      const focusRow = Number(msg.focusRow || 0);
      const selectedCodes = Array.isArray(msg.selectedCodes) ? msg.selectedCodes : [];
      if (!selectedCodes.length) return;

      state.activeTab = originTab;
      saveState();
      render();

      requestAnimationFrame(() => {
        const target = document.querySelector(
          `input[data-grid="calc"][data-tab="${originTab}"][data-row="${focusRow}"][data-col="0"]`
        );
        if (target) target.focus();

        if (selectedCodes.length > 1) window.__FIN_INSERT_CODES__?.(selectedCodes);
        else window.__FIN_INSERT_CODE__?.(selectedCodes[0]);
      });
      return;
    }

    if (msg.type === "UPDATE_CODES") {
      const incoming = Array.isArray(msg.codes) ? msg.codes : [];

      state.codeMaster = incoming
        .map(r => ({
          code: (r.code ?? "").toString().trim(),
          name: (r.name ?? "").toString(),
          spec: (r.spec ?? "").toString(),
          unit: (r.unit ?? "").toString(),
          surcharge: (r.surcharge === "" || r.surcharge == null) ? null : Number(r.surcharge),
          convUnit: (r.conv_unit ?? "").toString(),
          convFactor: (r.conv_factor === "" || r.conv_factor == null) ? null : Number(r.conv_factor),
          note: (r.note ?? "").toString(),
        }))
        .filter(x => x.code);

      saveState();
      render();
      return;
    }

    if (msg.type === "CLOSE_PICKER") {
      try { __pickerWin?.close(); } catch {}
      __pickerWin = null;
    }
  });

  window.__FIN_GET_CODEMASTER__ = () => state.codeMaster || [];
  window.__FIN_INSERT_CODE__ = (code) => { insertCodeToActiveCell(code); };
  window.__FIN_INSERT_CODES__ = (codes) => {
    insertCodeToActiveCell(codes[0] || "");

    const a = document.activeElement;
    if (!(a instanceof HTMLInputElement) || a.dataset.grid !== "calc") return;
    const tabId = a.dataset.tab;
    const startRow = Number(a.dataset.row);
    const col = Number(a.dataset.col);

    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    for (let i = 0; i < codes.length; i++) {
      const rIdx = startRow + i;
      if (rIdx >= sec.rows.length) sec.rows.push(defaultCalcRow());
      sec.rows[rIdx].code = String(codes[i] || "").toUpperCase();
    }
    recomputeSection(tabId);
    saveState();
    render();

    requestAnimationFrame(() => {
      const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${startRow}"][data-col="${col}"]`);
      if (target) target.focus();
    });
  };

  function insertCodeToActiveCell(code) {
    const a = document.activeElement;
    if (!(a instanceof HTMLInputElement) || a.dataset.grid !== "calc") return;

    const tabId = a.dataset.tab;
    const row = Number(a.dataset.row);

    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];
    if (!sec.rows[row]) return;

    sec.rows[row].code = String(code || "").toUpperCase().trim();
    recomputeSection(tabId);
    saveState();
    render();

    requestAnimationFrame(() => {
      const next = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${row}"][data-col="4"]`);
      if (next) next.focus();
    });
  }

  /***************
   * Export/Import/Reset
   ***************/
  function bindTopButtons() {
    const btnOpen = document.getElementById("btnOpenPicker");
    const btnExport = document.getElementById("btnExport");
    const btnReset = document.getElementById("btnReset");
    const fileImport = document.getElementById("fileImport");

    if (btnOpen) btnOpen.onclick = openCodePicker;

    if (btnExport) btnExport.onclick = () => {
      const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "FIN_state.json";
      a.click();
      URL.revokeObjectURL(url);
    };

    if (fileImport) fileImport.onchange = async (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      const txt = await f.text();
      try {
        const obj = JSON.parse(txt);
        state = { ...deepClone(DEFAULT_STATE), ...obj };
        saveState();
        render();
      } catch {
        alert("가져오기(JSON) 실패: 파일 형식이 올바르지 않습니다.");
      }
      e.target.value = "";
    };

    if (btnReset) btnReset.onclick = () => {
      if (!confirm("정말 초기화할까요? (로컬 저장 데이터가 삭제됩니다)")) return;
      localStorage.removeItem(LS_KEY);
      state = loadState();
      render();
    };
  }

  /***************
   * Render
   ***************/
  function applyPanelStickyTop() {
    // 산출탭(상단 top-split 존재)은 withTopSplit, 나머지는 base
    const root = document.documentElement;
    const isCalcTab = (state.activeTab === "steel" || state.activeTab === "steel_sub" || state.activeTab === "support");
    root.style.setProperty("--panelStickyTop", isCalcTab ? "var(--stickyWithTopSplitTop)" : "var(--stickyBaseTop)");
  }

  function render() {
    renderTabs();
    clear($view);

    let content = null;

    if (state.activeTab === "code") content = renderCodeTab();
    else if (state.activeTab === "steel") content = renderCalcTab("steel", "철골");
    else if (state.activeTab === "steel_sub") content = renderCalcTab("steel_sub", "철골_부자재");
    else if (state.activeTab === "support") content = renderCalcTab("support", "구조이기/동바리");
    else if (state.activeTab === "steel_sum") content = renderSummaryTabByCodeOrder("steel", "철골_집계");
    else if (state.activeTab === "support_sum") content = renderSummaryTabByCodeOrder("support", "구조이기/동바리_집계");

    $view.appendChild(content);
    bindTopButtons();

    // ✅ 렌더 후 실제 높이 측정 → sticky offset 보정
    requestAnimationFrame(() => {
      updateStickyVars();
      applyPanelStickyTop();
      // 한 번 더(폰/브라우저에서 레이아웃 확정 후)
      requestAnimationFrame(updateStickyVars);
    });
  }

  function renderSummaryTabByCodeOrder(srcTabId, title) {
    const bucket = state[srcTabId];

    // codeMaster 순서를 위한 인덱스 맵
    const orderMap = new Map();
    (state.codeMaster || []).forEach((cm, idx) => {
      const c = String(cm.code || "").trim().toUpperCase();
      if (c) orderMap.set(c, idx);
    });

    // 코드별 누적: code -> { code,name,spec,unit, pre, post, pctSet }
    const map = new Map();

    const prev = bucket.activeSection;

    for (let sIdx = 0; sIdx < bucket.sections.length; sIdx++) {
      bucket.activeSection = sIdx;
      recomputeSection(srcTabId);

      const sec = bucket.sections[sIdx];
      for (const r of sec.rows) {
        const code = String(r.code || "").trim().toUpperCase();
        if (!code) continue;

        const name = r.name || "";
        const spec = r.spec || "";
        const unit = r.unit || "";

        const pre = Number(r.value) || 0; // 할증전수량(= 물량)
        const mul = Number(r.surchargeMul) || 1;
        const post = pre * mul; // 할증후수량(= 물량 * (1+할증))

        const pct =
          (r.surchargePct == null || r.surchargePct === "" || !Number.isFinite(Number(r.surchargePct)))
            ? null
            : Number(r.surchargePct);

        if (!map.has(code)) {
          map.set(code, {
            code, name, spec, unit,
            pre: 0,
            post: 0,
            pctSet: new Set()
          });
        }

        const agg = map.get(code);
        agg.pre += pre;
        agg.post += post;

        // 할증 표시용(동일/혼합)
        if (pct == null) agg.pctSet.add("__NULL__");
        else agg.pctSet.add(String(pct));
      }
    }

    bucket.activeSection = prev;
    saveState();

    // ✅ 정렬 규칙:
    // 1) 코드탭(codeMaster)에 존재하면 그 순서(orderMap 인덱스)
    // 2) 코드탭에 없으면 맨 뒤로, 그 내부는 코드 문자열 오름차순
    const items = [...map.values()].sort((a, b) => {
      const ai = orderMap.has(a.code) ? orderMap.get(a.code) : Number.POSITIVE_INFINITY;
      const bi = orderMap.has(b.code) ? orderMap.get(b.code) : Number.POSITIVE_INFINITY;
      if (ai !== bi) return ai - bi;
      return a.code.localeCompare(b.code);
    });

    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [
        el("div", { class: "panel-title" }, [title]),
        el("div", { class: "panel-desc" }, [
          "코드별 집계: 할증전수량(물량 합) / 할증후수량(물량×(1+할증)) · 정렬: 코드탭 순서"
        ])
      ])
    ]);

    return el("div", { class: "panel" }, [
      panelHeader,
      el("div", { class: "table-wrap" }, [
        el("table", {}, [
          el("thead", {}, [
            el("tr", {}, [
              el("th", {}, ["코드"]),
              el("th", {}, ["품명"]),
              el("th", {}, ["규격"]),
              el("th", {}, ["단위"]),
              el("th", {}, ["할증전수량"]),
              el("th", {}, ["할증(%)"]),
              el("th", {}, ["할증후수량"]),
            ])
          ]),
          el("tbody", {}, [
            ...items.map(x => {
              const pctText = (() => {
                const s = x.pctSet;
                if (s.size === 0) return "";
                if (s.size === 1) {
                  const only = [...s][0];
                  if (only === "__NULL__") return "";
                  return only;
                }
                return "혼합";
              })();

              return el("tr", {}, [
                el("td", {}, [x.code]),
                el("td", {}, [x.name]),
                el("td", {}, [x.spec]),
                el("td", {}, [x.unit]),
                el("td", {}, [String(round4(x.pre))]),
                el("td", {}, [pctText]),
                el("td", {}, [String(round4(x.post))]),
              ]);
            }),
          ])
        ])
      ])
    ]);
  }

  function round4(n) {
    const v = Number(n) || 0;
    return Math.round(v * 10000) / 10000;
  }

  /***************
   * Init
   ***************/
  render();
})();
