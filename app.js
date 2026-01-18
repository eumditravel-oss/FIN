/* app.js (FINAL FIX v13.2b+) - FIN 산출자료 (Web)
   - ✅ (v13.0) 내보내기/가져오기: JSON → Excel(.xlsx) 기반으로 변경
   - ✅ (v13.0) 내보내기 클릭 시 탭 선택 팝업(모달) 제공 (코드/철골/철골_부자재/구조이기-동바리)
   - ✅ (v13.0) 가져오기(Excel): Codes 시트 기반으로 codeMaster 갱신 (임시 양식)
   - ✅ (v12.4) 산출표(계산표)에서 "비고" 컬럼만 숨김(렌더링 제거)
   - ✅ (v12.3) 변수표 영역에서도 Ctrl+F3/Shift+Ctrl+F3 행추가 지원 (변수표 셀 선택 시)
   - ✅ (v12.3) 집계 탭: 구분 개소(count) 반영하여 코드별 수량 합산
   - ✅ (v12.3) 집계 탭: 환산단위/환산계수 있으면 환산후수량 기준으로 단위/할증전/후 집계
   - ✅ (v12.3) 산출표 헤더 "물량(Value)" -> "물량"
   - ✅ (v12.3) 산출표 컬럼폭: 단위/물량(및 코드) 가로폭 증가 (CALC_COL_WEIGHTS 조정)
   - ✅ (v13.1) 도움말 버튼 추가: 화면 안내문구 제거 + help.html로 이동
   - ✅ (v13.2) 구분명 리스트: 클릭 후에도 ↑/↓ 키로 이동 가능(렌더 후 포커스 복원)
   - ✅ (v13.2a) 내보내기 모달 '전체선택' 버튼이 실제 체크박스에 반영되도록 수정(모달 재오픈 제거)
   - ✅ (v13.2b) top-split(구분/변수) ↔ panel 사이 리사이저(split-resizer) 적용 + 높이 상태 저장(ui.topSplitH)
   - ✅ (v13.2b) section-editor(구분 편집) CSS(3컬럼)와 맞게 버튼들을 한 칸으로 묶음

   - 🛠 (Patch) LS_KEY 버전 분리 + 구버전(V11) 데이터 자동 마이그레이션 + 초기화 시 구키도 함께 삭제
   - 🛠 (Patch) 프로젝트 모달 show/hide: hidden + aria-hidden 동시 지원(접근성/표준)
   - 🛠 (Patch) Init/Render 중복 호출 제거, bindTopButtons 1회만 바인딩
*/

(() => {
  "use strict";

  /***************
   * Storage (✅ Project-ready)
   ***************/
  const PROJECT_INDEX_KEY = "FIN_PROJECT_INDEX_V1";
  const PROJECT_ACTIVE_KEY = "FIN_PROJECT_ACTIVE_V1";
  const PROJECT_STATE_PREFIX = "FIN_PROJECT_STATE_V1::";

  // (기존 단일 저장키 마이그레이션용)
  const LS_KEY_OLD_SINGLE_V13 = "FIN_WEB_STATE_V13_2A";
  const LS_KEY_OLD_SINGLE_V11 = "FIN_WEB_STATE_V11";

  const deepClone = (obj) => JSON.parse(JSON.stringify(obj));
  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  /***************
   * ✅ focus jump 방지 헬퍼
   ***************/
  function safeFocus(target) {
    if (!target) return;
    try {
      target.focus({ preventScroll: true });
    } catch {
      try { target.focus(); } catch {}
    }
  }

  function raf2(fn) {
    requestAnimationFrame(() => requestAnimationFrame(fn));
  }

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

    const base = Math.ceil(topbarH + tabsH);
    root.style.setProperty("--stickyBaseTop", `${base}px`);

    const withTopSplit = Math.ceil(topbarH + tabsH + topSplitH + 10);
    root.style.setProperty("--stickyWithTopSplitTop", `${withTopSplit}px`);
  }

  window.addEventListener("resize", () => {
    requestAnimationFrame(() => {
      updateStickyVars();
      applyPanelStickyTop();
      updateScrollHeights();
      updateViewFillHeight();
    });
  });

  /***************
   * ✅ 내부 스크롤 높이 자동 보정 (PATCH: 하단 공백 제거)
   ***************/
  function updateScrollHeights() {
    const scrolls = document.querySelectorAll(".calc-scroll");
    if (!scrolls.length) return;

    scrolls.forEach((sc) => {
      if (!(sc instanceof HTMLElement)) return;

      sc.style.overflow = "auto";
      sc.style.webkitOverflowScrolling = "touch";
      sc.tabIndex = -1;

      // ✅ 중요: flex 컨테이너 안에서 스크롤 영역이 제대로 줄어들도록
      sc.style.minHeight = "0";

      const scRect = sc.getBoundingClientRect();
      const viewportH = window.innerHeight || document.documentElement.clientHeight || 800;

      const bottomPad = 12;

      const panel = sc.closest(".panel");
      let h = 0;

      if (panel instanceof HTMLElement) {
        const panelRect = panel.getBoundingClientRect();
        h = Math.floor(panelRect.bottom - scRect.top - bottomPad);
      } else {
        h = Math.floor(viewportH - scRect.top - bottomPad);
      }

      h = clamp(h, 160, 20000);

      sc.style.maxHeight = "";
      sc.style.height = `${h}px`;
    });
  }

  /***************
   * Code Master
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

    ui: {
      topSplitH: 190,
    }
  };

  /***************
   * ✅ Project Store Adapter
   ***************/
  const ProjectStore = (() => {
    const local = {
      loadIndex() {
        try {
          const raw = localStorage.getItem(PROJECT_INDEX_KEY);
          const parsed = raw ? JSON.parse(raw) : null;
          if (!parsed || !Array.isArray(parsed.projects)) return { projects: [] };
          return parsed;
        } catch { return { projects: [] }; }
      },
      saveIndex(index) {
        localStorage.setItem(PROJECT_INDEX_KEY, JSON.stringify(index));
      },
      loadActiveId() {
        return localStorage.getItem(PROJECT_ACTIVE_KEY) || "";
      },
      saveActiveId(id) {
        if (!id) localStorage.removeItem(PROJECT_ACTIVE_KEY);
        else localStorage.setItem(PROJECT_ACTIVE_KEY, id);
      },
      loadProjectState(id) {
        try {
          const k = PROJECT_STATE_PREFIX + id;
          const raw = localStorage.getItem(k);
          return raw ? JSON.parse(raw) : null;
        } catch {
          return null;
        }
      },
      saveProjectState(id, projectState) {
        const k = PROJECT_STATE_PREFIX + id;
        localStorage.setItem(k, JSON.stringify(projectState));
      },
      deleteProject(id) {
        localStorage.removeItem(PROJECT_STATE_PREFIX + id);
      }
    };
    return local;
  })();

  function genId() {
    return "p_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  }

  function normalizeProjectMeta(p) {
    return {
      id: String(p?.id || genId()),
      name: String(p?.name || "새 프로젝트"),
      code: String(p?.code || ""),
      updatedAt: Number(p?.updatedAt || Date.now()),
      createdAt: Number(p?.createdAt || Date.now()),
    };
  }

  function loadProjectIndex() {
    const idx = ProjectStore.loadIndex();
    return { projects: Array.isArray(idx.projects) ? idx.projects.map(normalizeProjectMeta) : [] };
  }

  function saveProjectIndex(index) {
    ProjectStore.saveIndex(index);
  }

  function loadProjectState(projectId) {
    try {
      const parsed = ProjectStore.loadProjectState(projectId);
      if (!parsed) return deepClone(DEFAULT_STATE);

      const s = { ...deepClone(DEFAULT_STATE), ...parsed };
      s.codeMaster = Array.isArray(parsed?.codeMaster) ? parsed.codeMaster : deepClone(DEFAULT_CODE_MASTER);

      for (const k of ["steel", "steel_sub", "support"]) {
        if (!s[k] || !Array.isArray(s[k].sections) || s[k].sections.length === 0) {
          s[k] = deepClone(DEFAULT_STATE[k]);
        }
        s[k].activeSection = clamp(Number(s[k].activeSection || 0), 0, s[k].sections.length - 1);
      }

      if (!s.ui || typeof s.ui !== "object") s.ui = deepClone(DEFAULT_STATE.ui);
      s.ui.topSplitH = clamp(Number(s.ui.topSplitH ?? 190), 120, 520);

      if (!TABS.some(t => t.id === s.activeTab)) s.activeTab = "code";
      return s;
    } catch (e) {
      console.warn("loadProjectState failed:", e);
      return deepClone(DEFAULT_STATE);
    }
  }

  function saveProjectState(projectId) {
    if (!projectId) return;
    ProjectStore.saveProjectState(projectId, deepClone(state));
  }

  // ✅ activeProjectId가 준비되기 전 호출 방지 포함
  function saveState() {
    if (!activeProjectId) return;
    saveProjectState(activeProjectId);
  }

  let projectIndex = loadProjectIndex();
  let activeProjectId = ProjectStore.loadActiveId();

  /***************
   * ✅ Legacy migration(단일키 -> 프로젝트 1회 이관)
   ***************/
  (function migrateLegacySingleToProjectOnce() {
    const legacy = localStorage.getItem(LS_KEY_OLD_SINGLE_V13) || localStorage.getItem(LS_KEY_OLD_SINGLE_V11);
    if (!legacy) return;
    if (projectIndex.projects.length > 0) return;

    try {
      const parsed = JSON.parse(legacy);
      const pid = genId();
      const meta = normalizeProjectMeta({ id: pid, name: "마이그레이션 프로젝트", code: "LEGACY" });
      projectIndex.projects.push(meta);
      saveProjectIndex(projectIndex);
      ProjectStore.saveActiveId(pid);
      activeProjectId = pid;

      ProjectStore.saveProjectState(pid, { ...deepClone(DEFAULT_STATE), ...parsed });
    } catch {}
  })();

  (function cleanupLegacyKeys() {
    if (projectIndex.projects.length <= 0) return;
    try { localStorage.removeItem(LS_KEY_OLD_SINGLE_V13); } catch {}
    try { localStorage.removeItem(LS_KEY_OLD_SINGLE_V11); } catch {}
  })();

  (function ensureAtLeastOneProject() {
    if (projectIndex.projects.length > 0) {
      if (!activeProjectId || !projectIndex.projects.some(p => p.id === activeProjectId)) {
        activeProjectId = projectIndex.projects[0].id;
        ProjectStore.saveActiveId(activeProjectId);
      }
      return;
    }

    const pid = genId();
    const meta = normalizeProjectMeta({ id: pid, name: "프로젝트 1", code: "" });
    projectIndex.projects.push(meta);
    saveProjectIndex(projectIndex);

    activeProjectId = pid;
    ProjectStore.saveActiveId(activeProjectId);
    ProjectStore.saveProjectState(pid, deepClone(DEFAULT_STATE));
  })();

  let state = activeProjectId ? loadProjectState(activeProjectId) : deepClone(DEFAULT_STATE);

  // ✅ (v13.2) 구분명 리스트 클릭/↑↓ 후 렌더링되면 포커스 복원
  let __pendingSectionFocus = null;

  /***************
   * ✅ Calc(산출표) 멀티선택 상태 (비저장/런타임)
   ***************/
  const __calcMulti = {
    active: false,
    tabId: null,
    sectionIndex: null,
    anchorRow: null,
    rows: new Set(),
  };

  function __calcMultiClear() {
    __calcMulti.active = false;
    __calcMulti.tabId = null;
    __calcMulti.sectionIndex = null;
    __calcMulti.anchorRow = null;
    __calcMulti.rows.clear();
  }

  function __calcMultiIsSameContext(tabId) {
    const bucket = state?.[tabId];
    const secIdx = bucket?.activeSection ?? 0;
    return __calcMulti.active && __calcMulti.tabId === tabId && __calcMulti.sectionIndex === secIdx;
  }

  function __calcMultiBegin(tabId, anchorRow) {
    const bucket = state?.[tabId];
    const secIdx = bucket?.activeSection ?? 0;

    __calcMulti.active = true;
    __calcMulti.tabId = tabId;
    __calcMulti.sectionIndex = secIdx;
    __calcMulti.anchorRow = clamp(
      Number(anchorRow || 0),
      0,
      (bucket?.sections?.[secIdx]?.rows?.length ?? 1) - 1
    );

    __calcMulti.rows.clear();
    __calcMulti.rows.add(__calcMulti.anchorRow);
  }

  function __calcMultiSetRange(tabId, fromRow, toRow) {
    if (!__calcMultiIsSameContext(tabId)) {
      __calcMultiBegin(tabId, fromRow);
    }
    const a = __calcMulti.anchorRow ?? fromRow;
    const lo = Math.min(a, toRow);
    const hi = Math.max(a, toRow);

    __calcMulti.rows.clear();
    for (let r = lo; r <= hi; r++) __calcMulti.rows.add(r);
  }

  function __applyCalcRowSelectionStyles(tabId) {
    const table = document
      .querySelector(`table.calc-table input[data-grid="calc"][data-tab="${tabId}"]`)
      ?.closest("table.calc-table");
    if (!table) return;

    const should = __calcMultiIsSameContext(tabId);
    const trs = table.querySelectorAll("tbody tr");
    trs.forEach((tr, i) => {
      if (should && __calcMulti.rows.has(i)) tr.classList.add("row-selected");
      else tr.classList.remove("row-selected");
    });
  }

  function __getSelectedCalcRows(tabId) {
    if (!__calcMultiIsSameContext(tabId)) return [];
    return [...__calcMulti.rows].sort((a, b) => a - b);
  }

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
   * ✅ (v13.2b) topSplit height 적용
   ***************/
  function applyTopSplitH() {
    const root = document.documentElement;
    const h = clamp(Number(state?.ui?.topSplitH ?? 190), 120, 520);
    root.style.setProperty("--topSplitH", `${Math.round(h)}px`);
  }

  /***************
   * ✅ zoom(--uiScale) 대응: view 높이 보정
   ***************/
  function getUiScale() {
    const v = getComputedStyle(document.documentElement).getPropertyValue("--uiScale").trim();
    const n = Number(v);
    return (Number.isFinite(n) && n > 0.2 && n < 2.5) ? n : 1;
  }

  function updateViewFillHeight() {
    const view = document.getElementById("view");
    if (!view) return;

    const scale = getUiScale();
    const vh = window.innerHeight || document.documentElement.clientHeight || 800;
    const target = Math.ceil(vh / scale);

    view.style.height = `${target}px`;
    view.style.minHeight = `${target}px`;
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
   * Column width helpers
   ***************/
  function buildColGroupFromWeights(weights) {
    const sum = weights.reduce((a, b) => a + b, 0);
    const cg = el("colgroup", {}, []);
    weights.forEach((w) => {
      const pct = (w / sum) * 100;
      cg.appendChild(el("col", { style: `width:${pct.toFixed(3)}%` }, []));
    });
    return cg;
  }

  const CALC_COL_WEIGHTS = [
    0.35,  // No
    0.75,  // 코드
    2.5,   // 품명(자동)
    2.5,   // 규격(자동)
    0.50,  // 단위(자동)
    2.5,   // 산출식
    0.50,  // 물량
    0.25,  // 할증(%)
    0.25,  // 환산단위
    0.25,  // 환산계수
    0.25,  // 환산후수량
  ];

  const CODE_COL_WEIGHTS = [0.6, 2.2, 2.2, 0.6, 0.6, 0.7, 0.7, 1.2, 0.6];

  /***************
   * ✅ Help
   ***************/
  function buildHelpPayload() {
    return {
      title: "FIN 산출자료 도움말",
      sections: [
        { title: "코드 선택(팝업)", items: [
          "Ctrl+. : 코드 선택 창 열기",
          "코드 선택 창에서 Ctrl+B : 다중선택",
          "코드 선택 창에서 Ctrl+Enter : 삽입",
        ]},
        { title: "표 이동/편집(공통)", items: [
          "방향키: 셀 이동",
          "F2: 편집 모드(읽기전용 셀 제외)",
          "편집 모드에서 Enter: 편집 종료",
          "PageUp / PageDown: 한 페이지 단위로 위/아래 이동(현재 열 유지)",
          "Ctrl+Home / Ctrl+End: 최상단/최하단으로 이동(현재 열 유지)"
        ]},
        { title: "행 추가/삭제", items: [
          "Ctrl+F3: 현재 행 아래 행 추가",
          "Shift+Ctrl+F3: +10행 추가",
          "Ctrl+Del: 삭제(확인창) - 산출표/코드표는 현재 '행' 삭제, 변수표는 현재 '셀' 비움",
          "ESC: (산출표 다중선택 중) 다중선택 취소"
        ]},
        { title: "산출 탭", items: [
          "구분 리스트: ↑/↓ 로 이동 및 선택",
          "구분/변수 영역 높이 조절: 중간 점선 바(리사이저)를 드래그"
        ]},
        { title: "산출표 다중선택", items: [
          "Shift+B: 다중선택 모드 토글",
          "Shift+↑ / Shift+↓: 다중선택 범위 확장",
          "Ctrl+Del: (다중선택 중) 선택된 행들을 한 번에 삭제",
          "Ctrl+G: (다중선택 중) 선택된 행들을 현재 행 아래로 복사/삽입"
        ]},
        { title: "엑셀 내보내기/가져오기", items: [
          "내보내기(EXCEL): 선택 모달에서 탭 선택 후 .xlsx 다운로드",
          "가져오기(EXCEL): 'Codes(또는 코드)' 시트 기반으로 codeMaster 갱신"
        ]},
      ]
    };
  }

  function openHelpWindow() {
    const w = window.open("help.html", "FIN_HELP", "width=980,height=820");
    if (!w) {
      alert("팝업이 차단되었습니다. 브라우저에서 팝업 허용 후 다시 시도해 주세요.");
      return;
    }

    const payload = buildHelpPayload();
    let tries = 0;
    const timer = setInterval(() => {
      tries++;
      try { w.postMessage({ type: "HELP_INIT", payload }, window.location.origin); } catch {}
      if (tries >= 20) clearInterval(timer);
    }, 120);
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
   * Code tab
   ***************/
  function renderCodeTab() {
    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [ el("div", { class: "panel-title" }, ["코드"]) ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => addCodeRows(1) }, ["행 추가 (Ctrl+F3)"]),
        el("button", { class: "smallbtn", onclick: () => addCodeRows(10) }, ["+10행"]),
      ])
    ]);

    const scroll = el("div", { class: "table-wrap calc-scroll", dataset: { scroll: "code" } }, [buildCodeMasterTable()]);
    forceScrollStyle(scroll);
    attachGridNav(scroll);
    attachWheelLock(scroll);

    return el("div", { class: "panel" }, [panelHeader, scroll]);
  }

  function buildCodeMasterTable() {
    const table = el("table", { class: "code-table" }, []);
    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    table.style.minWidth = "100%";

    table.appendChild(buildColGroupFromWeights(CODE_COL_WEIGHTS));

    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["코드"]),
        el("th", {}, ["품명"]),
        el("th", {}, ["규격"]),
        el("th", {}, ["단위"]),
        el("th", {}, ["할증"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["비고"]),
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

  const CODE_COL_INDEX = { code: 0, name: 1, spec: 2, unit: 3, surcharge: 4, convUnit: 5, convFactor: 6, note: 7 };

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

    input.addEventListener("blur", () => { delete input.dataset.editing; });

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

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const first = document.querySelector(`input[data-grid="code"][data-row="${insertPos}"][data-col="0"]`);
      if (first) safeFocus(first);
      ensureScrollIntoView();
    });
  }

  /***************
   * ✅ Split resizer
   ***************/
  function attachSplitResizer(resizerEl, topPaneEl) {
    if (!resizerEl || !topPaneEl) return;

    const root = document.documentElement;

    const begin = (clientY) => {
      const startH = topPaneEl.getBoundingClientRect().height;
      const startY = clientY;

      document.body.classList.add("is-resizing");

      const move = (y) => {
        const dy = y - startY;
        const next = clamp(startH + dy, 120, 520);
        state.ui.topSplitH = next;
        root.style.setProperty("--topSplitH", `${Math.round(next)}px`);
        saveState();
        updateStickyVars();
        applyPanelStickyTop();
        updateViewFillHeight();
        updateScrollHeights();
      };

      const onMove = (e) => {
        if (e.touches && e.touches[0]) move(e.touches[0].clientY);
        else move(e.clientY);
      };

      const end = () => {
        document.body.classList.remove("is-resizing");
        window.removeEventListener("mousemove", onMove, true);
        window.removeEventListener("mouseup", end, true);
        window.removeEventListener("touchmove", onMove, { capture: true });
        window.removeEventListener("touchend", end, true);
        window.removeEventListener("touchcancel", end, true);

        raf2(() => {
          updateStickyVars();
          applyPanelStickyTop();
          updateViewFillHeight();
          updateScrollHeights();
        });
      };

      window.addEventListener("mousemove", onMove, true);
      window.addEventListener("mouseup", end, true);
      window.addEventListener("touchmove", onMove, { capture: true, passive: false });
      window.addEventListener("touchend", end, true);
      window.addEventListener("touchcancel", end, true);
    };

    resizerEl.addEventListener("mousedown", (e) => {
      e.preventDefault();
      begin(e.clientY);
    });

    resizerEl.addEventListener("touchstart", (e) => {
      if (!e.touches || !e.touches[0]) return;
      e.preventDefault();
      begin(e.touches[0].clientY);
    }, { passive: false });
  }

  /***************
   * Calc tab
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
      el("div", {}, [ el("div", { class: "panel-title" }, [title]) ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => addRows(tabId, 1) }, ["행 추가 (Ctrl+F3)"]),
        el("button", { class: "smallbtn", onclick: () => addRows(tabId, 10) }, ["+10행"]),
      ])
    ]);

    const scroll = el("div", { class: "table-wrap calc-scroll", dataset: { scroll: "calc" } }, [buildCalcTable(tabId)]);
    forceScrollStyle(scroll);
    attachGridNav(scroll);
    attachWheelLock(scroll);

    const panel = el("div", { class: "panel" }, [panelHeader, scroll]);

    const topPane = el("div", { class: "pane top-pane" }, [top]);
    const resizer = el("div", { class: "split-resizer", dataset: { ui: "splitResizer" } }, []);
    const bottomPane = el("div", { class: "pane bottom-pane" }, [panel]);

    const workArea = el("div", { class: "work-area" }, [topPane, resizer, bottomPane]);

    raf2(() => {
      attachSplitResizer(resizer, topPane);
      updateViewFillHeight();
      updateScrollHeights();
    });

    return workArea;
  }

  function buildSectionList(tabId) {
    const bucket = state[tabId];

    const list = el("div", {
      class: "section-list",
      tabindex: "0",
      dataset: { nav: "sectionList", tab: tabId }
    }, []);

    bucket.sections.forEach((s, idx) => {
      const item = el("div", {
        class: "section-item" + (bucket.activeSection === idx ? " active" : ""),
        tabindex: "0",
        onclick: () => {
          bucket.activeSection = idx;
          saveState();
          __pendingSectionFocus = { tabId, index: idx };
          render();
        },
      }, [
        el("div", { class: "name" }, [s.name || `구분 ${idx + 1}`]),
        el("div", { class: "meta-inline" }, [`개소: ${s.count ?? ""}`]),
        el("div", { class: "meta" }, ["선택"])
      ]);
      list.appendChild(item);
    });

    list.addEventListener("mousedown", () => safeFocus(list));

    list.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
      const a = document.activeElement;
      if (a instanceof HTMLInputElement || a instanceof HTMLTextAreaElement) return;

      e.preventDefault();
      e.stopPropagation();

      const dir = e.key === "ArrowDown" ? 1 : -1;
      const nextIdx = clamp(bucket.activeSection + dir, 0, bucket.sections.length - 1);
      if (nextIdx === bucket.activeSection) return;

      bucket.activeSection = nextIdx;
      saveState();
      __pendingSectionFocus = { tabId, index: nextIdx };
      render();
    }, true);

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

    const btnWrap = el("div", { class: "row-actions", style: "justify-content:flex-end; gap:6px;" }, [
      el("button", { class: "smallbtn", onclick: () => { saveState(); render(); } }, ["저장"]),
      el("button", {
        class: "smallbtn",
        onclick: () => {
          bucket.sections.push(defaultSection(`구분 ${bucket.sections.length + 1}`, 1));
          bucket.activeSection = bucket.sections.length - 1;
          saveState(); render();
        }
      }, ["구분 추가"]),
      el("button", {
        class: "smallbtn",
        onclick: () => {
          if (bucket.sections.length <= 1) return alert("구분은 최소 1개가 필요합니다.");
          bucket.sections.splice(bucket.activeSection, 1);
          bucket.activeSection = clamp(bucket.activeSection, 0, bucket.sections.length - 1);
          saveState(); render();
        }
      }, ["구분 삭제"]),
    ]);

    return el("div", { class: "section-editor" }, [nameInput, countInput, btnWrap]);
  }

  function buildVarTable(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const wrap = el("div", { class: "var-tablewrap calc-scroll", dataset: { scroll: "var" } }, []);
    forceScrollStyle(wrap);
    attachWheelLock(wrap);

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

      updateViewFillHeight();
      updateScrollHeights();
    });

    attachGridNav(wrap);

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
    });

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

    input.addEventListener("blur", () => { delete input.dataset.editing; });
    return el("td", {}, [input]);
  }

  function buildCalcTable(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const table = el("table", { class: "calc-table" }, []);
    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    table.style.minWidth = "100%";

    table.appendChild(buildColGroupFromWeights(CALC_COL_WEIGHTS));

    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["No"]),
        el("th", {}, ["코드"]),
        el("th", {}, ["품명(자동)"]),
        el("th", {}, ["규격(자동)"]),
        el("th", {}, ["단위(자동)"]),
        el("th", {}, ["산출식"]),
        el("th", {}, ["물량"]),
        el("th", {}, ["할증(%)"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["환산후수량"]),
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
        tdNavInputCalc(tabId, i, 6, "surchargePct", r.surchargePct ?? "", { readonly: true }),
        tdNavInputCalc(tabId, i, 7, "convUnit", r.convUnit || "", { readonly: true }),
        tdNavInputCalc(tabId, i, 8, "convFactor", r.convFactor ?? "", { readonly: true }),
        tdNavInputCalc(tabId, i, 9, "converted", String(r.converted ?? 0), { readonly: true }),
      ]);
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    raf2(() => __applyCalcRowSelectionStyles(tabId));

    table.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (t.dataset.grid !== "calc") return;

      // =========================
// ✅ calc-table keydown (누락 복구 블록)
// =========================

// 편집 중이면 편집 종료만 허용
if (t.dataset.editing === "1") {
  if (e.key === "Enter") {
    e.preventDefault();
    delete t.dataset.editing;
    t.blur();
    raf2(() => safeFocus(t));
  }
  return;
}

const curRow = Number(t.dataset.row || 0);

// 일반 Delete → 현재 행 삭제
if ((e.key === "Delete" || e.key === "Del") && !e.ctrlKey) {
  e.preventDefault();
  if (!confirm("현재 행을 삭제할까요?")) return;
  deleteCalcRows(tabId, [curRow]);
  return;
}

       /* ============================
   ✅ CONTINUE FROM YOUR CUT
   (table.addEventListener("keydown"...)
============================ */

/* ---- Calc grid col index (calc table) ----
   No(0) / code(1) / name(2) / spec(3) / unit(4) / formula(5) / value(6) / surchargePct(7) / convUnit(8) / convFactor(9) / converted(10)
   ※ tdNavInputCalc에서는 "input"에 data-col을 아래처럼 매핑해서 넣습니다.
*/
const CALC_COL_INDEX = {
  code: 1,
  name: 2,
  spec: 3,
  unit: 4,
  formula: 5,
  value: 6,
  surchargePct: 7,
  convUnit: 8,
  convFactor: 9,
  converted: 10,
};

      // ✅ (계속) calc-table keydown
      // - Shift+B : 다중선택 토글
      // - Shift+↑/↓ : 다중선택 범위
      // - Ctrl+Del : 선택행 삭제
      // - Ctrl+G : 선택행 복사/삽입
      // - ESC : 다중선택 종료
      if (e.key === "Escape") {
        if (__calcMulti.active) {
          e.preventDefault();
          __calcMultiClear();
          __applyCalcRowSelectionStyles(tabId);
        }
        return;
      }

      if (e.key === "B" && e.shiftKey) {
        e.preventDefault();
        const r = Number(t.dataset.row || 0);
        if (!__calcMultiIsSameContext(tabId)) __calcMultiBegin(tabId, r);
        else __calcMultiClear();
        __applyCalcRowSelectionStyles(tabId);
        return;
      }

      // 다중선택 범위 확장
      if ((e.key === "ArrowUp" || e.key === "ArrowDown") && e.shiftKey) {
        e.preventDefault();
        const r = Number(t.dataset.row || 0);
        const next = clamp(r + (e.key === "ArrowDown" ? 1 : -1), 0, sec.rows.length - 1);
        if (!__calcMultiIsSameContext(tabId)) __calcMultiBegin(tabId, r);
        __calcMultiSetRange(tabId, __calcMulti.anchorRow ?? r, next);
        __applyCalcRowSelectionStyles(tabId);

        // 포커스는 이동 대상 row 같은 col로
        raf2(() => {
          const col = t.dataset.col || "0";
          const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${next}"][data-col="${col}"]`);
          safeFocus(target);
          ensureScrollIntoView(target);
        });
        return;
      }

      // 선택행 삭제
      if ((e.key === "Delete" || e.key === "Del") && e.ctrlKey) {
        e.preventDefault();
        const selected = __getSelectedCalcRows(tabId);
        if (!selected.length) {
          // 단일행 삭제로 처리(현재행)
          const row = Number(t.dataset.row || 0);
          deleteCalcRows(tabId, [row]);
          return;
        }
        if (!confirm(`선택된 ${selected.length}행을 삭제할까요?`)) return;
        deleteCalcRows(tabId, selected);
        __calcMultiClear();
        return;
      }

      // 선택행 복사/삽입
      if ((e.key === "g" || e.key === "G") && e.ctrlKey) {
        e.preventDefault();
        const selected = __getSelectedCalcRows(tabId);
        if (!selected.length) return;
        const anchor = Number(t.dataset.row || 0);
        duplicateCalcRows(tabId, selected, anchor);
        return;
      }
    }, true);

    // ✅ input 변화가 있을 때 재계산 + 저장 + 렌더 반영
    table.addEventListener("input", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (t.dataset.grid !== "calc") return;
      if (t.dataset.tab !== tabId) return;

      const row = Number(t.dataset.row || 0);
      const field = t.dataset.field;

      const bucket2 = state[tabId];
      const sec2 = bucket2.sections[bucket2.activeSection];
      const rr = sec2.rows[row];
      if (!rr) return;

      if (field === "code") {
        rr.code = (t.value || "").trim();
      } else if (field === "formula") {
        rr.formula = t.value || "";
      } else {
        // readonly는 원칙적으로 여기에 안 옴
        rr[field] = t.value;
      }

      recomputeSection(tabId);
      saveState();
      refreshCalcComputed(tabId); // 값/환산/자동필드 갱신
    });

    return table;
  }

  function tdNavInputCalc(tabId, row, colNo, field, value, opts = {}) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    // colNo는 “표에서의 순번(0=code...)”로 들어오므로,
    // 실제 data-col은 CALC_COL_INDEX로 맞춤(그리드 네비게이션 일관성)
    let dataCol = String(colNo);
    if (field && CALC_COL_INDEX[field] != null) dataCol = String(CALC_COL_INDEX[field]);

    const input = el("input", {
      class: "cell" + (opts.readonly ? " readonly" : ""),
      value: value ?? "",
      placeholder: opts.placeholder || "",
      readonly: opts.readonly ? "readonly" : null,
      dataset: { grid: "calc", tab: tabId, row: String(row), col: dataCol, field },
      onfocus: (e) => {
        // 다중선택 컨텍스트 유지용: 포커스 행 표시
        if (__calcMulti.active && __calcMultiIsSameContext(tabId)) {
          __applyCalcRowSelectionStyles(tabId);
        }
      },
      onkeydown: (e) => {
        const t = e.target;
        if (!(t instanceof HTMLInputElement)) return;

        // F2: 편집모드 플래그
        if (e.key === "F2") {
          if (t.readOnly) return;
          e.preventDefault();
          t.dataset.editing = "1";
          t.setSelectionRange?.(t.value.length, t.value.length);
          return;
        }

        // Enter: 편집모드 종료
        if (e.key === "Enter") {
          if (t.dataset.editing === "1") {
            e.preventDefault();
            delete t.dataset.editing;
            t.blur();
            raf2(() => safeFocus(t));
            return;
          }
        }
      },
      oninput: (e) => {
        if (opts.readonly) return;
        const rr = sec.rows[row];
        if (!rr) return;
        rr[field] = e.target.value;
      }
    });

    input.addEventListener("blur", () => { delete input.dataset.editing; });

    return el("td", {}, [input]);
  }

  function refreshCalcComputed(tabId) {
    // 현재 tab의 calc-table에서 readonly 셀들 업데이트
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    sec.rows.forEach((r, i) => {
      const setVal = (field, v) => {
        const col = CALC_COL_INDEX[field];
        const inp = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${i}"][data-col="${col}"]`);
        if (inp) inp.value = (v ?? "");
      };

      setVal("name", r.name || "");
      setVal("spec", r.spec || "");
      setVal("unit", r.unit || "");
      setVal("value", String(r.value ?? 0));
      setVal("surchargePct", (r.surchargePct ?? "") === null ? "" : String(r.surchargePct ?? ""));
      setVal("convUnit", r.convUnit || "");
      setVal("convFactor", (r.convFactor ?? "") === null ? "" : String(r.convFactor ?? ""));
      setVal("converted", String(r.converted ?? 0));
    });

    // 다중선택 표시 갱신
    raf2(() => __applyCalcRowSelectionStyles(tabId));
  }

  function addRows(tabId, n, insertAfterRow = null) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const idx = insertAfterRow == null ? (sec.rows.length - 1) : insertAfterRow;
    const insertPos = clamp(idx + 1, 0, sec.rows.length);

    const newRows = Array.from({ length: n }, () => defaultCalcRow());
    sec.rows.splice(insertPos, 0, ...newRows);

    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const first = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="${CALC_COL_INDEX.code}"]`);
      safeFocus(first);
      ensureScrollIntoView(first);
    });
  }

  function deleteCalcRows(tabId, rowIndices) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];
    const uniq = [...new Set(rowIndices)].sort((a, b) => b - a); // 뒤에서부터 삭제

    uniq.forEach((r) => {
      if (r >= 0 && r < sec.rows.length) sec.rows.splice(r, 1);
    });

    if (sec.rows.length === 0) sec.rows.push(defaultCalcRow());

    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const targetRow = clamp(Math.min(...rowIndices), 0, sec.rows.length - 1);
      const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${targetRow}"][data-col="${CALC_COL_INDEX.code}"]`);
      safeFocus(target);
      ensureScrollIntoView(target);
    });
  }

  function duplicateCalcRows(tabId, rowIndices, insertAfterRow) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const selected = [...new Set(rowIndices)].sort((a, b) => a - b);
    const clones = selected
      .map((r) => sec.rows[r])
      .filter(Boolean)
      .map((r) => deepClone(r));

    if (!clones.length) return;

    const insertPos = clamp((insertAfterRow ?? selected[selected.length - 1]) + 1, 0, sec.rows.length);
    sec.rows.splice(insertPos, 0, ...clones);

    saveState();
    render();

    raf2(() => {
      updateViewFillHeight();
      updateScrollHeights();
      const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="${CALC_COL_INDEX.code}"]`);
      safeFocus(target);
      ensureScrollIntoView(target);
    });
  }

  /* ============================
     ✅ Summation Tabs (steel_sum / support_sum)
     - (v12.3) 개소(count) 반영
     - (v12.3) 환산단위/환산계수 있으면 converted 기준 집계
  ============================ */
  function buildSummaryRows(tabId) {
    const bucket = state[tabId];
    const map = new Map();

    bucket.sections.forEach((sec) => {
      const count = Number(sec.count ?? 1);
      const mult = Number.isFinite(count) && count > 0 ? count : 1;

      // section별로 vars/rows 값이 계산되어 있어야 함
      // recomputeSection는 activeSection만 계산하므로, 여기선 간단히 현재 저장값(value/converted)을 사용
      sec.rows.forEach((r) => {
        const code = (r.code || "").trim();
        if (!code) return;

        const info = codeLookup(code);
        const unit = info?.unit || r.unit || "";
        const surcharge = (r.surchargePct == null ? (info?.surcharge ?? null) : r.surchargePct);

        // 환산계수 있으면 converted 기준
        const hasConv = r.convFactor != null && Number.isFinite(Number(r.convFactor)) && Number(r.convFactor) !== 0;
        const qty = hasConv ? Number(r.converted || 0) : Number((r.value || 0) * (r.surchargeMul || 1));

        const key = code.toUpperCase();
        const prev = map.get(key) || {
          code,
          name: info?.name || r.name || "",
          spec: info?.spec || r.spec || "",
          unit,
          convUnit: info?.convUnit || r.convUnit || "",
          convFactor: info?.convFactor ?? r.convFactor ?? null,
          surchargePct: surcharge,
          qty: 0,
        };
        prev.qty += qty * mult;
        map.set(key, prev);
      });
    });

    return [...map.values()].sort((a, b) => String(a.code).localeCompare(String(b.code)));
  }

  function renderSummaryTab(srcTabId, title) {
    const rows = buildSummaryRows(srcTabId);

    const header = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [ el("div", { class: "panel-title" }, [title]) ]),
      el("div", { class: "row-actions" }, [
        el("button", { class: "smallbtn", onclick: () => { /* noop */ } }, ["집계(자동)"]),
      ]),
    ]);

    const table = el("table", { class: "code-table" }, []);
    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    table.style.minWidth = "100%";
    table.appendChild(buildColGroupFromWeights([0.9, 2.4, 2.4, 0.8, 0.8, 0.9, 0.9, 1.4, 1.2]));

    const thead = el("thead", {}, [
      el("tr", {}, [
        el("th", {}, ["코드"]),
        el("th", {}, ["품명"]),
        el("th", {}, ["규격"]),
        el("th", {}, ["단위"]),
        el("th", {}, ["할증"]),
        el("th", {}, ["환산단위"]),
        el("th", {}, ["환산계수"]),
        el("th", {}, ["수량(환산/할증 반영)"]),
        el("th", {}, ["비고"]),
      ])
    ]);

    const tbody = el("tbody", {}, []);
    rows.forEach((r) => {
      tbody.appendChild(el("tr", {}, [
        el("td", {}, [r.code]),
        el("td", {}, [r.name || ""]),
        el("td", {}, [r.spec || ""]),
        el("td", {}, [r.unit || ""]),
        el("td", {}, [r.surchargePct == null ? "" : String(r.surchargePct)]),
        el("td", {}, [r.convUnit || ""]),
        el("td", {}, [r.convFactor == null ? "" : String(r.convFactor)]),
        el("td", {}, [String(Math.round((Number(r.qty) || 0) * 1000) / 1000)]),
        el("td", {}, [""]),
      ]));
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    const scroll = el("div", { class: "table-wrap calc-scroll", dataset: { scroll: "sum" } }, [table]);
    forceScrollStyle(scroll);
    attachWheelLock(scroll);

    return el("div", { class: "panel" }, [header, scroll]);
  }

  /* ============================
     ✅ Grid Navigation (Arrow/PageUp/PageDown/Home/End)
     - (간단 구현) data-grid="code|var|calc"
  ============================ */
  function parseCellDataset(input) {
    const ds = input?.dataset || {};
    return {
      grid: ds.grid || "",
      tab: ds.tab || "",
      row: Number(ds.row || 0),
      col: Number(ds.col || 0),
    };
  }

  function queryCell(grid, tab, row, col) {
    const selector =
      grid === "code"
        ? `input[data-grid="code"][data-row="${row}"][data-col="${col}"]`
        : `input[data-grid="${grid}"][data-tab="${tab}"][data-row="${row}"][data-col="${col}"]`;
    return document.querySelector(selector);
  }

  function moveCell(fromInput, dRow, dCol, pageJump = false) {
    const { grid, tab, row, col } = parseCellDataset(fromInput);
    if (!grid) return;

    // row/col 범위 추정
    let maxRow = 0;
    let maxCol = 0;

    const all = grid === "code"
      ? document.querySelectorAll(`input[data-grid="code"]`)
      : document.querySelectorAll(`input[data-grid="${grid}"][data-tab="${tab}"]`);

    all.forEach((x) => {
      const r = Number(x.dataset.row || 0);
      const c = Number(x.dataset.col || 0);
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    });

    let nextRow = clamp(row + dRow, 0, maxRow);
    let nextCol = clamp(col + dCol, 0, maxCol);

    if (pageJump) {
      // pageJump일 때는 scroller 높이 기준으로 row를 대략 이동
      const sc = fromInput.closest(".calc-scroll");
      if (sc) {
        const rect = sc.getBoundingClientRect();
        const rowH = 34; // 대략
        const jump = Math.max(1, Math.floor(rect.height / rowH) - 1);
        nextRow = clamp(row + (dRow > 0 ? jump : -jump), 0, maxRow);
      }
    }

    const target = queryCell(grid, tab, nextRow, nextCol);
    if (target) {
      safeFocus(target);
      ensureScrollIntoView(target);
    }
  }

  function attachGridNav(container) {
    if (!container) return;
    container.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (!t.dataset.grid) return;

      // 편집중이면 방향키 이동 막음
      if (t.dataset.editing === "1") return;

      const isInput = (document.activeElement instanceof HTMLInputElement);
      if (!isInput) return;

      const key = e.key;

      if (key === "ArrowUp") { e.preventDefault(); moveCell(t, -1, 0); }
      else if (key === "ArrowDown") { e.preventDefault(); moveCell(t, 1, 0); }
      else if (key === "ArrowLeft") { e.preventDefault(); moveCell(t, 0, -1); }
      else if (key === "ArrowRight") { e.preventDefault(); moveCell(t, 0, 1); }
      else if (key === "PageUp") { e.preventDefault(); moveCell(t, -1, 0, true); }
      else if (key === "PageDown") { e.preventDefault(); moveCell(t, 1, 0, true); }
      else if (key === "Home" && e.ctrlKey) { e.preventDefault(); moveCell(t, -99999, 0); }
      else if (key === "End" && e.ctrlKey) { e.preventDefault(); moveCell(t, 99999, 0); }
      else if ((key === "Delete" || key === "Del") && e.ctrlKey) {
        // Ctrl+Del: 변수표는 현재 셀 비움 / 코드표/산출표는 “현재 행 삭제”를 상단 핸들러에서 처리
        const grid = t.dataset.grid;
        if (grid === "var") {
          if (t.readOnly) return;
          e.preventDefault();
          t.value = "";
          t.dispatchEvent(new Event("input", { bubbles: true }));
        }
      }
    }, true);
  }

  /* ============================
     ✅ wheel lock (trackpad/space bounce 방지)
  ============================ */
  function attachWheelLock(scroller) {
    if (!scroller) return;
    scroller.addEventListener("wheel", (e) => {
      // 기본 스크롤 허용(단, 바깥으로 튀는 스크롤만 차단)
      const el = scroller;
      const delta = e.deltaY;
      if (delta < 0 && el.scrollTop <= 0) e.preventDefault();
      else if (delta > 0 && el.scrollTop + el.clientHeight >= el.scrollHeight) e.preventDefault();
    }, { passive: false });
  }

  function forceScrollStyle(sc) {
    if (!sc || !(sc instanceof HTMLElement)) return;
    sc.style.overflow = "auto";
    sc.style.webkitOverflowScrolling = "touch";
    sc.style.minHeight = "0";
    sc.tabIndex = -1;
  }

  function ensureScrollIntoView(target) {
    if (!target || !(target instanceof HTMLElement)) return;
    const sc = target.closest(".calc-scroll");
    if (!sc) return;
    const tRect = target.getBoundingClientRect();
    const sRect = sc.getBoundingClientRect();

    const pad = 8;
    if (tRect.top < sRect.top + pad) {
      sc.scrollTop -= (sRect.top + pad - tRect.top);
    } else if (tRect.bottom > sRect.bottom - pad) {
      sc.scrollTop += (tRect.bottom - (sRect.bottom - pad));
    }
  }

  /* ============================
     ✅ Sticky Panel Top 적용
  ============================ */
  function applyPanelStickyTop() {
    const root = document.documentElement;
    const top = state.activeTab === "code"
      ? getComputedStyle(root).getPropertyValue("--stickyBaseTop").trim()
      : getComputedStyle(root).getPropertyValue("--stickyWithTopSplitTop").trim();

    document.querySelectorAll('[data-sticky="panel"]').forEach((h) => {
      if (!(h instanceof HTMLElement)) return;
      h.style.top = top || "0px";
    });
  }

  /* ============================
     ✅ Export / Import (placeholder-safe)
     - XLSX가 페이지에 로드되어 있으면 실제로 동작
     - 없으면 alert로 안내 (런타임 에러 방지)
  ============================ */
  function exportToExcelSelectedTabs(tabIds) {
    // XLSX가 없는 경우 대비
    if (!window.XLSX) {
      alert("XLSX 라이브러리가 로드되지 않았습니다.\nexport.js 또는 CDN이 필요합니다.");
      return;
    }
    // TODO: 기존 v13.0 export 로직이 있다면 여기로 연결
    alert("내보내기 로직 연결 지점입니다. (기존 v13.0 export 함수로 연결하세요)");
  }

  function importFromExcelFile(file) {
    if (!window.XLSX) {
      alert("XLSX 라이브러리가 로드되지 않았습니다.\nimport.js 또는 CDN이 필요합니다.");
      return;
    }
    // TODO: 기존 v13.0 import 로직이 있다면 여기로 연결
    alert("가져오기 로직 연결 지점입니다. (기존 v13.0 import 함수로 연결하세요)");
  }

  /* ============================
     ✅ Top Buttons (bind once)
  ============================ */
  let __topButtonsBound = false;
  function bindTopButtonsOnce() {
    if (__topButtonsBound) return;
    __topButtonsBound = true;

    const btnHelp = document.getElementById("btnHelp");
    if (btnHelp) btnHelp.addEventListener("click", openHelpWindow);

    // ✅ 프로젝트 열기(상단 버튼)
    const btnProjectOpen = document.getElementById("btnProjectOpen");
    if (btnProjectOpen) btnProjectOpen.addEventListener("click", openProjectModal);

    // ✅ 프로젝트 모달 버튼
    const btnProjectAdd = document.getElementById("btnProjectAdd");
    const btnProjectClose = document.getElementById("btnProjectClose");
    if (btnProjectAdd) btnProjectAdd.addEventListener("click", createProject);
    if (btnProjectClose) btnProjectClose.addEventListener("click", closeProjectModal);

    // ✅ 내보내기
    const btnExport = document.getElementById("btnExport");
    if (btnExport) btnExport.addEventListener("click", () => {
      // 기존 v13.0 모달이 있다면 그 모달을 여는 함수로 연결
      // 임시: 현재 탭만 내보내기
      exportToExcelSelectedTabs([state.activeTab]);
    });

    // ✅ 가져오기
    const fileImport = document.getElementById("fileImport");
    if (fileImport) fileImport.addEventListener("change", (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      importFromExcelFile(f);
      e.target.value = "";
    });

    // ✅ 리셋
    const btnReset = document.getElementById("btnReset");
    if (btnReset) btnReset.addEventListener("click", () => {
      if (!activeProjectId) return;
      if (!confirm("현재 프로젝트를 초기화할까요?")) return;
      state = deepClone(DEFAULT_STATE);
      saveState();
      render();
    });

    // ✅ ESC로 프로젝트 모달 닫기
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      const modal = document.getElementById("projectModal");
      if (!modal) return;
      const hidden = modal.getAttribute("aria-hidden");
      if (hidden === "false") closeProjectModal();
    });
  }

  /* ============================
     ✅ Project UI (list 기반) — 너가 전에 올린 방식 유지
     (projectName/projectCode/projectModal/projectList 기준)
  ============================ */
  function setTopButtonsEnabled(enabled) {
    const btnOpen = document.getElementById("btnOpenPicker");
    const btnExport = document.getElementById("btnExport");
    const btnReset = document.getElementById("btnReset");
    const fileImport = document.getElementById("fileImport");
    const btnImportWrap = document.getElementById("btnImportWrap");

    if (btnOpen) btnOpen.disabled = !enabled;
    if (btnExport) btnExport.disabled = !enabled;
    if (btnReset) btnReset.disabled = !enabled;
    if (fileImport) fileImport.disabled = !enabled;

    if (btnImportWrap) {
      btnImportWrap.style.opacity = enabled ? "1" : "0.45";
      btnImportWrap.style.pointerEvents = enabled ? "auto" : "none";
      btnImportWrap.setAttribute("aria-disabled", enabled ? "false" : "true");
    }

    const help = document.getElementById("btnHelp");
    if (help) help.disabled = false;
  }

  function updateProjectHeaderUI() {
    const meta = projectIndex.projects.find(p => p.id === activeProjectId);
    const $name = document.getElementById("projectName");
    const $code = document.getElementById("projectCode");

    if ($name) $name.textContent = meta ? meta.name : "프로젝트 미선택";
    if ($code) $code.textContent = meta ? (`공사코드 ${meta.code || "-"}`) : "공사코드 -";

    setTopButtonsEnabled(!!meta);
  }

  function openProjectModal() {
    const modal = document.getElementById("projectModal");
    if (!modal) return;

    // Patch: hidden + aria-hidden 동시 지원
    modal.hidden = false;
    modal.setAttribute("aria-hidden", "false");

    renderProjectList();
  }

  function closeProjectModal() {
    const modal = document.getElementById("projectModal");
    if (!modal) return;

    modal.hidden = true;
    modal.setAttribute("aria-hidden", "true");
  }

  function renderProjectList() {
    const $list = document.getElementById("projectList");
    if (!$list) return;
    $list.innerHTML = "";

    const items = projectIndex.projects
      .slice()
      .sort((a,b) => (b.updatedAt||0) - (a.updatedAt||0));

    if (!items.length) {
      const empty = document.createElement("div");
      empty.style.padding = "12px";
      empty.style.color = "rgba(90,90,97,1)";
      empty.style.fontWeight = "700";
      empty.textContent = "프로젝트가 없습니다. [프로젝트 추가]로 생성하세요.";
      $list.appendChild(empty);
      return;
    }

    items.forEach(p => {
      const row = document.createElement("div");
      row.className = "project-row" + (p.id === activeProjectId ? " active" : "");

      const name = document.createElement("input");
      name.className = "cell";
      name.value = p.name || "";
      name.placeholder = "프로젝트명";

      const code = document.createElement("input");
      code.className = "cell";
      code.value = p.code || "";
      code.placeholder = "공사코드";

      const btnSave = document.createElement("button");
      btnSave.className = "smallbtn";
      btnSave.textContent = "저장";
      btnSave.onclick = () => {
        if (p.id === activeProjectId) saveProjectState(activeProjectId);

        p.name = name.value.trim() || "새 프로젝트";
        p.code = code.value.trim();
        p.updatedAt = Date.now();

        saveProjectIndex(projectIndex);
        updateProjectHeaderUI();
        renderProjectList();
      };

      const btnOpen = document.createElement("button");
      btnOpen.className = "smallbtn";
      btnOpen.textContent = "열기";
      btnOpen.onclick = () => {
        selectProject(p.id);
        closeProjectModal();
      };

      const btnDel = document.createElement("button");
      btnDel.className = "smallbtn";
      btnDel.textContent = "삭제";
      btnDel.onclick = () => {
        if (!confirm(`프로젝트를 삭제할까요?\n${p.name} (${p.code || "-"})`)) return;
        deleteProject(p.id);
      };

      row.appendChild(name);
      row.appendChild(code);
      row.appendChild(btnSave);
      row.appendChild(btnOpen);
      row.appendChild(btnDel);

      $list.appendChild(row);
    });
  }

  function createProject() {
    const pid = genId();
    const meta = normalizeProjectMeta({ id: pid, name: "새 프로젝트", code: "" });

    projectIndex.projects.push(meta);
    saveProjectIndex(projectIndex);

    ProjectStore.saveProjectState(pid, deepClone(DEFAULT_STATE));
    selectProject(pid);

    renderProjectList();
  }

  function deleteProject(projectId) {
    if (projectId === activeProjectId) {
      try { saveProjectState(activeProjectId); } catch {}
      activeProjectId = "";
      ProjectStore.saveActiveId("");
    }

    projectIndex.projects = projectIndex.projects.filter(p => p.id !== projectId);
    saveProjectIndex(projectIndex);
    ProjectStore.deleteProject(projectId);

    updateProjectHeaderUI();
    renderProjectList();

    if (!activeProjectId) {
      state = deepClone(DEFAULT_STATE);
      render();
    }
  }

  function selectProject(projectId) {
    const meta = projectIndex.projects.find(p => p.id === projectId);
    if (!meta) return alert("프로젝트를 찾을 수 없습니다.");

    // 현재 프로젝트 저장
    if (activeProjectId) saveProjectState(activeProjectId);

    // 새 프로젝트 로드
    state = loadProjectState(projectId);

    activeProjectId = projectId;
    ProjectStore.saveActiveId(activeProjectId);

    updateProjectHeaderUI();
    render();
  }

  /* ============================
     ✅ Render Main
  ============================ */
  function render() {
    if (!$view) return;

    applyTopSplitH();
    renderTabs();

    clear($view);

    let node = null;
    if (state.activeTab === "code") node = renderCodeTab();
    else if (state.activeTab === "steel") node = renderCalcTab("steel", "철골");
    else if (state.activeTab === "steel_sum") node = renderSummaryTab("steel", "철골_집계");
    else if (state.activeTab === "steel_sub") node = renderCalcTab("steel_sub", "철골_부자재");
    else if (state.activeTab === "support") node = renderCalcTab("support", "구조이기/동바리");
    else if (state.activeTab === "support_sum") node = renderSummaryTab("support", "구조이기/동바리_집계");
    else node = renderCodeTab();

    $view.appendChild(node);

    // sticky / heights patch
    raf2(() => {
      updateStickyVars();
      applyPanelStickyTop();
      updateViewFillHeight();
      updateScrollHeights();

      // (v13.2) 구분 포커스 복원
      if (__pendingSectionFocus && __pendingSectionFocus.tabId === state.activeTab) {
        const list = document.querySelector(`.section-list[data-tab="${__pendingSectionFocus.tabId}"]`);
        const idx = __pendingSectionFocus.index;
        const item = list?.querySelectorAll(".section-item")?.[idx];
        raf2(() => safeFocus(item));
        __pendingSectionFocus = null;
      }
    });
  }

  /* ============================
     ✅ Init (중복 호출 제거)
  ============================ */
  bindTopButtonsOnce();
  updateProjectHeaderUI();
  render();

  // 초기 높이/스티키 계산
  raf2(() => {
    updateStickyVars();
    applyPanelStickyTop();
    updateViewFillHeight();
    updateScrollHeights();
  });

})();


