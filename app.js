/* app.js (FINAL FIX v13.2b) - FIN 산출자료 (Web)
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
   - ✅ (v13.2b) ✅ top-split(구분/변수) ↔ panel 사이 리사이저(split-resizer) 적용 + 높이 상태 저장(ui.topSplitH)
   - ✅ (v13.2b) ✅ section-editor(구분 편집) CSS(3컬럼)와 맞게 버튼들을 한 칸으로 묶음

   - 🛠 (Patch) LS_KEY 버전 분리 + 구버전(V11) 데이터 자동 마이그레이션 + 초기화 시 구키도 함께 삭제
*/

(() => {
  "use strict";

  /***************
   * Storage (✅ Project-ready)
   * - 프로젝트 목록/선택/CRUD는 "index"로 관리
   * - 실제 산출 데이터는 프로젝트별 key로 저장
   * - 추후 서버 연동 시 adapter만 교체하면 됨
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
    });
  });

  /***************
   * ✅ 내부 스크롤 높이 자동 보정 (PATCH: 하단 공백 제거)
   * - viewport 기준 maxHeight → panel 기준 "height" 고정
   * - flex 레이아웃과 max-height 충돌로 생기던 하단 공백 방지
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

      // ✅ 우선순위:
      // 1) scroll이 들어있는 panel 기준으로 bottom까지 꽉 채우기
      // 2) 없으면 viewport 기준 fallback
      const panel = sc.closest(".panel");
      let h = 0;

      if (panel instanceof HTMLElement) {
        const panelRect = panel.getBoundingClientRect();
        h = Math.floor(panelRect.bottom - scRect.top - bottomPad);
      } else {
        h = Math.floor(viewportH - scRect.top - bottomPad);
      }

      h = clamp(h, 160, 20000);

      // ✅ 핵심: maxHeight 대신 height로 고정
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

    // ✅ (v13.2b) UI 상태(리사이저 높이)
    ui: {
      topSplitH: 190, // CSS :root --topSplitH 기본값과 맞춤
    }
  };

  /***************
   * ✅ Project Store Adapter (Local now, Server later)
   ***************/
  const ProjectStore = (() => {
    // ---- Local adapter ----
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

    // ---- Server adapter placeholder ----
    // 나중에 서버 붙이면 아래 객체를 server로 구현해서 return만 바꾸면 됨.
    // const server = { ...same methods... }

    return local;
  })();

  function genId() {
    // 충돌 방지용(서버 붙여도 무난)
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

  /***************
   * ✅ Project-aware load/save
   ***************/
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
    ProjectStore.saveProjectState(projectId, deepClone(state)); // ✅ 안전
  }

  // ✅ activeProjectId가 준비되기 전 호출 방지 포함
  function saveState() {
    if (!activeProjectId) return;           // ✅ 프로젝트 선택 전엔 저장 안 함
    saveProjectState(activeProjectId);
  }

  let projectIndex = loadProjectIndex();
  let activeProjectId = ProjectStore.loadActiveId();

  // ✅ (마이그레이션) 예전 단일키가 남아있으면 "마이그레이션 프로젝트"로 1회 옮김
  (function migrateLegacySingleToProjectOnce(){
    const legacy = localStorage.getItem(LS_KEY_OLD_SINGLE_V13) || localStorage.getItem(LS_KEY_OLD_SINGLE_V11);
    if (!legacy) return;

    // 이미 프로젝트가 있으면 마이그레이션 생략
    if (projectIndex.projects.length > 0) return;

    try {
      const parsed = JSON.parse(legacy);
      const pid = genId();
      const meta = normalizeProjectMeta({ id: pid, name: "마이그레이션 프로젝트", code: "LEGACY" });
      projectIndex.projects.push(meta);
      saveProjectIndex(projectIndex);
      ProjectStore.saveActiveId(pid);
      activeProjectId = pid;

      // 상태 저장
      ProjectStore.saveProjectState(pid, { ...deepClone(DEFAULT_STATE), ...parsed });
    } catch {}
  })();

  // ✅ (Patch) 초기화 시 구키도 함께 삭제 (마이그레이션 되었거나, 프로젝트가 이미 존재하면 정리)
  (function cleanupLegacyKeys(){
    if (projectIndex.projects.length <= 0) return;
    try { localStorage.removeItem(LS_KEY_OLD_SINGLE_V13); } catch {}
    try { localStorage.removeItem(LS_KEY_OLD_SINGLE_V11); } catch {}
  })();

  // ✅ (안전) 프로젝트가 하나도 없으면 기본 프로젝트 1개 생성 (저장 가능 상태 보장)
  (function ensureAtLeastOneProject(){
    if (projectIndex.projects.length > 0) {
      // activeProjectId가 유효하지 않으면 첫 프로젝트로 보정
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
    ProjectStore.saveActiveId(pid);

    // 기본 상태 저장(바로 저장/렌더 안정)
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
    rows: new Set(),   // 선택된 row index들
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
    // ✅ 현재 탭(tabId)의 calc-table만 하이라이트 처리
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
   * ✅ (PATCH) zoom(--uiScale) 대응: 하단 공백 제거 + calc-scroll 높이 보정
   * - styles.css에서 :root { --uiScale: 0.75; } 같은 값을 사용하는 경우 대응
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

    // zoom은 "시각적" 크기만 줄여서 레이아웃 하단이 남는 문제가 생김
    // → view 자체 높이를 1/scale로 키워서 보이는 화면을 꽉 채움
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

  // ✅ (v12.4) 산출표 "비고" 컬럼 제거 → weights도 1개 줄임
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
   * ✅ Help contents
   ***************/
  function buildHelpPayload() {
    return {
      title: "FIN 산출자료 도움말",
      sections: [
        {
          title: "코드 선택(팝업)",
          items: [
            "Ctrl+. : 코드 선택 창 열기",
            "코드 선택 창에서 Ctrl+B : 다중선택",
            "코드 선택 창에서 Ctrl+Enter : 삽입",
            "검색 입력 후 Enter : 필터 적용(구현된 검색 방식에 따라 동작)"
          ]
        },
        {
          title: "표 이동/편집(공통)",
          items: [
            "방향키: 셀 이동",
            "F2: 편집 모드(읽기전용 셀 제외)",
            "편집 모드에서 Enter: 편집 종료",
            "PageUp / PageDown: 한 페이지 단위로 위/아래 이동(현재 열 유지)",
            "Ctrl+Home / Ctrl+End: 최상단/최하단으로 이동(현재 열 유지)"
          ]
        },
        {
          title: "행 추가/삭제",
          items: [
            "Ctrl+F3: 현재 행 아래 행 추가",
            "Shift+Ctrl+F3: +10행 추가",
            "Ctrl+Del: 삭제(확인창) - 산출표/코드표는 현재 '행' 삭제, 변수표는 현재 '셀' 비움",
            "ESC: (산출표 다중선택 중) 다중선택 취소"
          ]
        },
        {
          title: "산출 탭(철골/철골_부자재/구조이기/동바리)",
          items: [
            "산출식 Enter: 계산(재계산)",
            "구분 리스트: ↑/↓ 로 이동 및 선택",
            "변수표: A, AB, A1, AB1... 최대 3자(첫 글자는 영문)",
            "구분/변수 영역 높이 조절: 중간 점선 바(리사이저)를 드래그"
          ]
        },
        {
          title: "산출표(하단 표) 다중선택/다중작업",
          items: [
            "Shift+B: 다중선택 모드 토글(현재 행을 기준(anchor)으로 선택 시작/해제)",
            "Shift+↑ / Shift+↓: 다중선택 범위 확장(연속 선택)",
            "Ctrl+Del: (다중선택 중) 선택된 행들을 한 번에 삭제",
            "Ctrl+G: (다중선택 중) 선택된 행들을 현재 행 아래로 복사/삽입(행추가 붙여넣기)"
          ]
        },
        {
          title: "집계 탭",
          items: [
            "코드별 집계: 구분 개소(count) 반영",
            "환산단위/환산계수 있으면 환산후수량 기준으로 할증전/할증후 합산"
          ]
        },
        {
          title: "엑셀 내보내기/가져오기",
          items: [
            "내보내기(EXCEL): 선택 모달에서 탭 선택 후 .xlsx 다운로드",
            "가져오기(EXCEL): 'Codes(또는 코드)' 시트 기반으로 codeMaster 갱신(임시 매핑)"
          ]
        },
        {
          title: "UI/레이아웃(패치)",
          items: [
            "상단(구분/변수) ↔ 하단(산출표) 사이 높이 조절: 리사이저 드래그(높이 저장: ui.topSplitH)",
            "하단 공백 최소화: 내부 스크롤 높이 자동 보정(calc-scroll 자동 높이)"
          ]
        }
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
      try {
        w.postMessage({ type: "HELP_INIT", payload }, window.location.origin);
      } catch {}
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
      el("div", {}, [
        el("div", { class: "panel-title" }, ["코드"]),
      ]),
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
      // ✅ (PATCH) zoom 공백 제거/높이 보정
      updateViewFillHeight();
      updateScrollHeights();

      const first = document.querySelector(`input[data-grid="code"][data-row="${insertPos}"][data-col="0"]`);
      if (first) safeFocus(first);
      ensureScrollIntoView();
    });
  }

  /***************
   * ✅ (v13.2b) Split resizer: top-split 높이 조절
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

        // ✅ (PATCH) zoom 공백 제거/높이 보정
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

          // ✅ (PATCH) zoom 공백 제거/높이 보정
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
      el("div", {}, [
        el("div", { class: "panel-title" }, [title]),
      ]),
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

    // ✅ (v13.2b) work-area / top-pane / split-resizer / bottom-pane 구조로 렌더
    const topPane = el("div", { class: "pane top-pane" }, [top]);
    const resizer = el("div", { class: "split-resizer", dataset: { ui: "splitResizer" } }, []);
    const bottomPane = el("div", { class: "pane bottom-pane" }, [panel]);

    const workArea = el("div", { class: "work-area" }, [topPane, resizer, bottomPane]);

    // 렌더 직후 attach
    raf2(() => {
      attachSplitResizer(resizer, topPane);

      // ✅ (PATCH) 첫 렌더에서도 zoom 공백 제거/높이 보정
      updateViewFillHeight();
      updateScrollHeights();
    });

    return workArea;
  }

  // ✅ (v13.2) 구분리스트: 클릭/↑/↓ 후 렌더링해도 포커스 유지 + ↑/↓ 이동 가능
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

    list.addEventListener("mousedown", () => {
      safeFocus(list);
    });

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

    // ✅ (v13.2b) section-editor CSS(3컬럼) 대응: 버튼을 한 칸으로 묶어서 배치
    const btnWrap = el("div", { class: "row-actions", style: "justify-content:flex-end; gap:6px;" }, [
      saveBtn, addBtn, delBtn
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

      // ✅ (PATCH) 변수 입력으로 높이/레이아웃이 변하면 공백이 다시 생길 수 있어 보정
      updateViewFillHeight();
      updateScrollHeights();
    });

    attachGridNav(wrap);

    // ✅ (PATCH) 최초 렌더 시에도 1회 보정
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

  // ✅ (v12.4) 산출표: "비고" 컬럼 제거
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

    // ✅ 멀티선택 하이라이트 반영(렌더 직후)
    raf2(() => __applyCalcRowSelectionStyles(tabId));

    table.addEventListener("keydown", (e) => {
      const t = e.target;
      if (!(t instanceof HTMLInputElement)) return;
      if (t.dataset.grid !== "calc") return;

      if (t.dataset.editing === "1" && e.key === "Enter") {
        e.preventDefault();
        delete t.dataset.editing;
        return;
      }

      if (e.key === "Enter") {
        e.preventDefault();
        recomputeSection(tabId);
        saveState();
        refreshCalcComputed(tabId);
      }
    }, true);

    return table;
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

    input.addEventListener("blur", () => { delete input.dataset.editing; });

    return el("td", {}, [input]);
  }

  function refreshCalcComputed(tabId) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const inputs = document.querySelectorAll(`input[data-grid="calc"][data-tab="${tabId}"]`);
    inputs.forEach((inp) => {
      const r = Number(inp.dataset.row);
      const f = inp.dataset.field;
      const rowObj = sec.rows[r];
      if (!rowObj) return;

      if (["name", "spec", "unit", "value", "convUnit", "convFactor", "converted"].includes(f)) {
        inp.value = (rowObj[f] ?? "") + "";
      }
    });
  }

  /***************
   * ✅ Grid navigation + F2 edit mode
   ***************/
  function attachGridNav(container) {
    container.addEventListener("keydown", (e) => {
      const t = e.target;
      const isInput = (t instanceof HTMLInputElement) || (t instanceof HTMLTextAreaElement);
      if (!isInput) return;

      const grid = t.dataset.grid;
      if (grid !== "calc" && grid !== "var" && grid !== "code") return;

       

             // ✅ (NEW) ESC : calc 멀티선택 취소
      if (grid === "calc" && e.key === "Escape") {
        const tabId = t.dataset.tab;
        if (__calcMultiIsSameContext(tabId)) {
          e.preventDefault();
          e.stopPropagation();
          __calcMultiClear();
          __applyCalcRowSelectionStyles(tabId);
        }
        return;
      }


       

       

      if (e.key === "F2") {
        if (t.hasAttribute("readonly")) return;
        e.preventDefault();
        t.dataset.editing = "1";
        try {
          const len = (t.value ?? "").length;
          t.setSelectionRange(len, len);
        } catch {}
        return;
      }
       

      if (t.dataset.editing === "1") {
        if (e.key === "Enter") {
          e.preventDefault();
          delete t.dataset.editing;
        }
        return;
      }

             // ✅ (NEW) calc-table 멀티선택: Shift + ↑/↓
      // - Shift만 누르고 이동하면 자동으로 멀티선택 시작/확장
      if (grid === "calc" && e.shiftKey && !e.ctrlKey && !e.altKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
        e.preventDefault();

        const row = Number(t.dataset.row);
        const col = Number(t.dataset.col);
        const tabId = t.dataset.tab;

        let nr = row + (e.key === "ArrowDown" ? 1 : -1);

        const bucket = state[tabId];
        const sec = bucket.sections[bucket.activeSection];
        nr = clamp(nr, 0, sec.rows.length - 1);

        // ✅ 선택 범위 갱신 (anchor ~ nr)
        __calcMultiSetRange(tabId, row, nr);

        // ✅ 포커스 이동 (같은 col 유지)
        const selector = `[data-grid="calc"][data-tab="${tabId}"][data-row="${nr}"][data-col="${col}"]`;
        const next = container.querySelector(selector);
        if (next && (next instanceof HTMLInputElement || next instanceof HTMLTextAreaElement)) {
          safeFocus(next);
          ensureScrollIntoView();
        }

        // ✅ 하이라이트 적용
        __applyCalcRowSelectionStyles(tabId);
        return;
      }


             // ✅ (NEW) PageUp / PageDown : 한 페이지 점프 (현재 열 유지)
      // ✅ (NEW) Ctrl+Home / Ctrl+End : 최상단/최하단 셀로 이동 (현재 열 유지)
      const isPage = (e.key === "PageDown" || e.key === "PageUp");
      const isHomeEnd = (e.key === "Home" || e.key === "End");

      if (isPage || isHomeEnd) {
        // 입력 중이면 브라우저 기본 동작 유지(원하면 막아도 됨)
        // 여기서는 "editing 모드가 아닐 때만" 처리
        e.preventDefault();
        e.stopPropagation();

        const tabId = t.dataset.tab || null;
        const row = Number(t.dataset.row || 0);
        const col = Number(t.dataset.col || 0);

        // grid별 rowCount 계산
        const getRowCount = () => {
          if (grid === "code") return (state.codeMaster?.length ?? 0);
          if (grid === "calc") {
            const bucket = state[tabId];
            const sec = bucket?.sections?.[bucket?.activeSection ?? 0];
            return (sec?.rows?.length ?? 0);
          }
          if (grid === "var") {
            const bucket = state[tabId];
            const sec = bucket?.sections?.[bucket?.activeSection ?? 0];
            return (sec?.vars?.length ?? 0);
          }
          return 0;
        };

        const rowCount = getRowCount();
        if (!rowCount) return;

        const clampRow = (r) => clamp(r, 0, rowCount - 1);

        // 스크롤 컨테이너(현재 셀이 들어있는 calc-scroll)를 기준으로 페이지 크기 계산
        const sc = t.closest(".calc-scroll") || container.closest(".calc-scroll") || container;

        const head = sc?.querySelector?.("thead");
        const headH = head ? Math.ceil(head.getBoundingClientRect().height) : 0;

        // 현재 행 높이(없으면 fallback)
        const tr = t.closest("tr");
        const rowH = tr ? Math.max(18, Math.ceil(tr.getBoundingClientRect().height)) : 34;

        // 한 페이지에 들어갈 row 개수(헤더 제외)
        const pageRows = Math.max(1, Math.floor((sc.clientHeight - headH - 12) / rowH));

        let targetRow = row;

        // Ctrl+Home / Ctrl+End
        if (e.ctrlKey && !e.shiftKey && !e.altKey && isHomeEnd) {
          targetRow = (e.key === "Home") ? 0 : (rowCount - 1);
        }
        // PageUp / PageDown
        else if (!e.ctrlKey && !e.shiftKey && !e.altKey && isPage) {
          targetRow = (e.key === "PageDown") ? (row + pageRows) : (row - pageRows);

          // 스크롤도 같이 페이지 단위로 이동(체감 “바로 이동”)
          const delta = (sc.clientHeight - headH - 12);
          sc.scrollTop += (e.key === "PageDown" ? delta : -delta);
        } else {
          // 다른 조합(예: Shift+PageDown)은 일단 무시
          return;
        }

        targetRow = clampRow(targetRow);

        // 다음 셀로 포커스 이동 (현재 열 유지)
        const baseSel = `[data-grid="${grid}"]` +
          (grid === "calc" || grid === "var" ? `[data-tab="${tabId}"]` : "") +
          `[data-row="${targetRow}"][data-col="${col}"]`;

        let next = container.querySelector(baseSel);

        // 혹시 해당 열이 없는 케이스 대비(안전장치): 같은 row의 col=0으로 fallback
        if (!next) {
          const fallbackSel = `[data-grid="${grid}"]` +
            (grid === "calc" || grid === "var" ? `[data-tab="${tabId}"]` : "") +
            `[data-row="${targetRow}"][data-col="0"]`;
          next = container.querySelector(fallbackSel);
        }

        if (next && (next instanceof HTMLInputElement || next instanceof HTMLTextAreaElement)) {
          safeFocus(next);
          ensureScrollIntoView();
        }

        return;
      }


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

      let selector = `[data-grid="${grid}"][data-row="${nr}"][data-col="${nc}"]`;
if (grid === "calc" || grid === "var") {
  const tabId = t.dataset.tab;
  selector = `[data-grid="${grid}"][data-tab="${tabId}"][data-row="${nr}"][data-col="${nc}"]`;
}

      const next = container.querySelector(selector);

      if (next && ((next instanceof HTMLInputElement) || (next instanceof HTMLTextAreaElement))) {
        safeFocus(next);
        ensureScrollIntoView();
      }
    }, true);
  }

  /***************
   * scroll helpers
   ***************/
  function forceScrollStyle(scrollEl) {
    if (!scrollEl) return;
    scrollEl.style.overflow = "auto";
    scrollEl.style.webkitOverflowScrolling = "touch";
    scrollEl.tabIndex = -1;
  }

  function attachWheelLock(scrollEl) {
    if (!scrollEl) return;

    scrollEl.addEventListener("wheel", (e) => {
      const canScroll = scrollEl.scrollHeight > scrollEl.clientHeight + 2;
      if (!canScroll) return;

      e.preventDefault();
      scrollEl.scrollTop += e.deltaY;
    }, { passive: false });
  }

  function ensureScrollIntoView() {
    const a = document.activeElement;
    if (!(a instanceof HTMLElement)) return;

    const scroll = a.closest(".calc-scroll");
    if (!scroll) return;

    const r = a.getBoundingClientRect();
    const s = scroll.getBoundingClientRect();

    const thead = scroll.querySelector("thead");
    const headH = thead ? Math.ceil(thead.getBoundingClientRect().height) : 0;

    const topPad = headH + 6;
    const botPad = 6;

    if (r.top < s.top + topPad) {
      scroll.scrollTop -= (s.top + topPad - r.top);
    } else if (r.bottom > s.bottom - botPad) {
      scroll.scrollTop += (r.bottom - (s.bottom - botPad));
    }
  }

  /***************
   * Row add/delete/shortcuts/picker/export/import/reset
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

    raf2(() => {
      updateScrollHeights();
      const first = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="0"]`);
      if (first) safeFocus(first);
      ensureScrollIntoView();
    });
  }

  function addVarRows(tabId, n, insertAfterRow = null) {
    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const idx = (insertAfterRow == null) ? (sec.vars.length - 1) : insertAfterRow;
    const insertPos = clamp(idx + 1, 0, sec.vars.length);

    const newRows = Array.from({ length: n }, () => defaultVarRow());
    sec.vars.splice(insertPos, 0, ...newRows);

    recomputeSection(tabId);
    saveState();
    render();

    raf2(() => {
      updateScrollHeights();
      const first = document.querySelector(`input[data-grid="var"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="0"]`);
      if (first) safeFocus(first);
      ensureScrollIntoView();
    });
  }

  function deleteCalcRowAtActiveCell(inputEl) {
    const tabId = inputEl.dataset.tab;
    const row = Number(inputEl.dataset.row);
    const col = Number(inputEl.dataset.col);

    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    if (!sec?.rows?.length) return;
    if (sec.rows.length <= 1) {
      sec.rows[0] = defaultCalcRow();
    } else {
      sec.rows.splice(row, 1);
    }

    recomputeSection(tabId);
    saveState();
    render();

    raf2(() => {
      updateScrollHeights();
      const nr = clamp(row, 0, (sec.rows.length - 1));
      const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${nr}"][data-col="${col}"]`);
      if (target) safeFocus(target);
      ensureScrollIntoView();
    });
  }

  function deleteCodeMasterRowAtActiveCell(inputEl) {
    const row = Number(inputEl.dataset.row);
    const col = Number(inputEl.dataset.col);

    if (!Array.isArray(state.codeMaster)) return;
    if (state.codeMaster.length <= 1) {
      state.codeMaster[0] = { code:"", name:"", spec:"", unit:"", surcharge:null, convUnit:"", convFactor:null, note:"" };
    } else {
      state.codeMaster.splice(row, 1);
    }

    saveState();
    render();

    raf2(() => {
      updateScrollHeights();
      const nr = clamp(row, 0, state.codeMaster.length - 1);
      const target = document.querySelector(`input[data-grid="code"][data-row="${nr}"][data-col="${col}"]`);
      if (target) safeFocus(target);
      ensureScrollIntoView();
    });
  }

  window.addEventListener("keydown", (e) => {
         // ✅ (NEW) ESC : calc 멀티선택 취소(포커스가 input이 아니어도 동작)
    if (!e.ctrlKey && !e.shiftKey && !e.altKey && e.key === "Escape") {
      if (__calcMulti.active) {
        e.preventDefault();
        e.stopPropagation();
        const tabId = __calcMulti.tabId;
        __calcMultiClear();
        if (tabId) __applyCalcRowSelectionStyles(tabId);
      }
      return;
    }

    if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key === ".") {
      e.preventDefault();
      e.stopPropagation();
      openCodePicker();
      return;
    }

         // ✅ (NEW) Shift+B : 멀티선택 토글(현재 행을 anchor로)
    if (!e.ctrlKey && e.shiftKey && !e.altKey && (e.key === "B" || e.key === "b")) {
      const a = document.activeElement;
      if (!(a instanceof HTMLInputElement)) return;
      if (a.dataset?.grid !== "calc") return;

      e.preventDefault();
      e.stopPropagation();

      const tabId = a.dataset.tab;
      const row = Number(a.dataset.row);

      if (__calcMultiIsSameContext(tabId)) {
        __calcMultiClear();
      } else {
        __calcMultiBegin(tabId, row);
      }
      __applyCalcRowSelectionStyles(tabId);
      return;
    }

    // ✅ (NEW) Ctrl+G : 멀티선택된 "행들"을 현재 행 아래로 복사(행추가)
    if (e.ctrlKey && !e.shiftKey && !e.altKey && (e.key === "g" || e.key === "G")) {
      const a = document.activeElement;
      if (!(a instanceof HTMLInputElement)) return;
      if (a.dataset?.grid !== "calc") return;

      const tabId = a.dataset.tab;
      const col = Number(a.dataset.col);
      const curRow = Number(a.dataset.row);

      const selected = __getSelectedCalcRows(tabId);
      if (!selected.length) return; // 선택 없으면 아무것도 안 함

      e.preventDefault();
      e.stopPropagation();

      const bucket = state[tabId];
      const sec = bucket.sections[bucket.activeSection];

      // ✅ 선택된 행 데이터 deep clone
      const clones = selected.map((ri) => deepClone(sec.rows[ri] || defaultCalcRow()));

      // ✅ "선택된 셀(현재 포커스 행)" 아래로 삽입
      const insertPos = clamp(curRow + 1, 0, sec.rows.length);
      sec.rows.splice(insertPos, 0, ...clones);

      recomputeSection(tabId);
      saveState();
      render();

      raf2(() => {
        updateScrollHeights();
        const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${insertPos}"][data-col="${col}"]`);
        if (target) safeFocus(target);

        // 복붙 후 선택은 해제(원하면 유지로 바꿔줄 수 있음)
        __calcMultiClear();
        __applyCalcRowSelectionStyles(tabId);

        ensureScrollIntoView();
      });
      return;
    }


    const isCtrlDel =
      e.ctrlKey &&
      !e.shiftKey &&
      !e.altKey &&
      (
        e.key === "Delete" || e.key === "Del" || e.key === "Backspace" ||
        e.code === "Delete" || e.code === "Backspace" ||
        e.keyCode === 46 || e.keyCode === 8
      );

        if (isCtrlDel) {
      const a = document.activeElement;
      const isEditableEl = (a instanceof HTMLInputElement) || (a instanceof HTMLTextAreaElement);
      if (!isEditableEl) return;

      const grid = a.dataset?.grid;
      if (grid !== "calc" && grid !== "var" && grid !== "code") return;

      // ✅ calc 멀티선택이 있으면 readonly여도 "행 삭제"는 허용
      const tabId = a.dataset?.tab;
      const hasMulti = (grid === "calc" && tabId && __getSelectedCalcRows(tabId).length > 0);

      if (!hasMulti) {
        // 기존 정책 유지: readonly 셀에서는 삭제 금지(단일)
        if (a.hasAttribute("readonly")) return;
      }

      const ok = confirm("정말로 삭제할까요?\n- 산출표/코드표: 현재 '행'이 삭제됩니다.\n- 변수표: 현재 '셀'이 비워집니다.");
      if (!ok) {
        e.preventDefault();
        e.stopPropagation();
        return;
      }

      e.preventDefault();
      e.stopPropagation();

      // ✅ (NEW) calc 멀티선택 행 삭제
      if (hasMulti) {
        const bucket = state[tabId];
        const sec = bucket.sections[bucket.activeSection];

        const col = Number(a.dataset.col);
        const selected = __getSelectedCalcRows(tabId);

        // 뒤에서부터 삭제(인덱스 꼬임 방지)
        for (let i = selected.length - 1; i >= 0; i--) {
          const rIdx = selected[i];
          if (sec.rows.length <= 1) {
            sec.rows[0] = defaultCalcRow();
            break;
          }
          sec.rows.splice(rIdx, 1);
        }

        recomputeSection(tabId);
        saveState();
        render();

        raf2(() => {
          updateScrollHeights();

          // 삭제 후 포커스: "가장 위 선택 행" 위치로 복원
          const base = selected[0] ?? 0;
          const nr = clamp(base, 0, sec.rows.length - 1);
          const target = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${nr}"][data-col="${col}"]`);
          if (target) safeFocus(target);

          __calcMultiClear();
          __applyCalcRowSelectionStyles(tabId);
          ensureScrollIntoView();
        });
        return;
      }

      // ✅ 기존 단일 삭제 로직
      if (grid === "calc") { deleteCalcRowAtActiveCell(a); return; }
      if (grid === "code") { deleteCodeMasterRowAtActiveCell(a); return; }

      // var: 셀 비움
      a.value = "";
      a.dispatchEvent(new Event("input", { bubbles: true }));
      return;
    }


    if (e.ctrlKey && (e.key === "F3")) {
      const a = document.activeElement;
      const isEditableEl = (a instanceof HTMLInputElement) || (a instanceof HTMLTextAreaElement);
      if (!isEditableEl) return;

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

      if (grid === "var") {
        e.preventDefault();
        e.stopPropagation();
        const tabId = a.dataset.tab;
        const row = Number(a.dataset.row);
        if (e.shiftKey) addVarRows(tabId, 10, row);
        else addVarRows(tabId, 1, row);
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
  }, { capture: true });

  /***************
   * Code Picker Popup (기존 그대로)
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

      raf2(() => {
        updateScrollHeights();
        const target = document.querySelector(
          `input[data-grid="calc"][data-tab="${originTab}"][data-row="${focusRow}"][data-col="0"]`
        );
        if (target) safeFocus(target);

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
    const a = document.activeElement;
    if (!(a instanceof HTMLInputElement) || a.dataset.grid !== "calc") return;

    const tabId = a.dataset.tab;
    const startRowRaw = Number(a.dataset.row);
    const col = Number(a.dataset.col);

    const bucket = state[tabId];
    const sec = bucket.sections[bucket.activeSection];

    const startRow = clamp(startRowRaw, 0, sec.rows.length);
    const insertRows = codes.map(c => {
      const r = defaultCalcRow();
      r.code = String(c || "").toUpperCase().trim();
      return r;
    });

    sec.rows.splice(startRow, 0, ...insertRows);

    recomputeSection(tabId);
    saveState();
    render();

    raf2(() => {
      updateScrollHeights();
      const target = document.querySelector(
        `input[data-grid="calc"][data-tab="${tabId}"][data-row="${startRow}"][data-col="${col}"]`
      );
      if (target) safeFocus(target);
      ensureScrollIntoView();
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

    raf2(() => {
      updateScrollHeights();
      const next = document.querySelector(`input[data-grid="calc"][data-tab="${tabId}"][data-row="${row}"][data-col="4"]`);
      if (next) safeFocus(next);
      ensureScrollIntoView();
    });
  }

  /***************
   * ✅ Excel Modal Styles (app.js에서 자동 주입)
   ***************/
  function ensureExcelModalStyles() {
    if (document.getElementById("excel-modal-style")) return;

    const css = `
      .excel-modal-backdrop{
        position:fixed; inset:0;
        background: rgba(0,0,0,.25);
        display:flex; align-items:center; justify-content:center;
        z-index: 99999;
        padding:16px;
      }
      .excel-modal{
        width:min(520px, 96vw);
        background: rgba(255,250,240,.96);
        border: 1px solid rgba(0,0,0,.10);
        border-radius: 18px;
        box-shadow: 0 24px 60px rgba(0,0,0,.18);
        overflow:hidden;
      }
      .excel-modal-head{
        padding:14px 16px;
        border-bottom:1px solid rgba(0,0,0,.08);
        display:flex; align-items:center; justify-content:space-between; gap:12px;
      }
      .excel-modal-title{ font-weight:900; }
      .excel-modal-body{ padding:14px 16px; }
      .excel-modal-foot{
        padding:14px 16px;
        border-top:1px solid rgba(0,0,0,.08);
        display:flex; justify-content:flex-end; gap:8px; flex-wrap:wrap;
      }
      .excel-modal-list{ display:flex; flex-direction:column; gap:10px; }
      .excel-modal-item{
        display:flex; align-items:center; justify-content:space-between; gap:12px;
        padding:10px 12px;
        background: rgba(255,255,255,.55);
        border: 1px solid rgba(0,0,0,.10);
        border-radius: 14px;
      }
      .excel-modal-item label{ font-weight:900; color:#1d1d1f; }
      .excel-modal-item small{ color: rgba(90,90,97,1); font-weight:700; }
      .excel-modal-item input[type="checkbox"]{ width:18px; height:18px; }
    `;
    const style = document.createElement("style");
    style.id = "excel-modal-style";
    style.textContent = css;
    document.head.appendChild(style);
  }

  /***************
   * ✅ Excel Export Modal + Export/Import 구현
   ***************/
  function openExcelExportModal() {
    ensureExcelModalStyles();
    document.querySelectorAll(".excel-modal-backdrop").forEach(n => n.remove());

    const selections = {
      code: true,
      steel: true,
      steel_sub: false,
      support: false,
    };

    const checkboxMap = Object.create(null);

    const makeItem = (key, title, desc) => {
      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.checked = !!selections[key];
      chk.addEventListener("change", () => selections[key] = chk.checked);

      checkboxMap[key] = chk;

      return el("div", { class: "excel-modal-item" }, [
        el("div", {}, [
          el("label", {}, [title]),
          el("div", {}, [el("small", {}, [desc])]),
        ]),
        chk
      ]);
    };

    const backdrop = el("div", { class: "excel-modal-backdrop" }, []);
    const modal = el("div", { class: "excel-modal" }, []);

    const head = el("div", { class: "excel-modal-head" }, [
      el("div", { class: "excel-modal-title" }, ["엑셀 내보내기"]),
      el("button", { class: "smallbtn", onclick: () => backdrop.remove() }, ["닫기"])
    ]);

    const body = el("div", { class: "excel-modal-body" }, [
      el("div", { class: "excel-modal-list" }, [
        makeItem("code", "코드", "Codes 시트로 codeMaster를 내보냅니다."),
        makeItem("steel", "철골", "Steel 시트로 산출/변수를 내보냅니다."),
        makeItem("steel_sub", "철골_부자재", "Steel_Sub 시트로 산출/변수를 내보냅니다."),
        makeItem("support", "구조이기/동바리", "Support 시트로 산출/변수를 내보냅니다."),
      ])
    ]);

    const foot = el("div", { class: "excel-modal-foot" }, [
      el("button", {
        class: "btn ghost",
        onclick: () => {
          for (const k of Object.keys(selections)) {
            selections[k] = true;
            if (checkboxMap[k]) checkboxMap[k].checked = true;
          }
        }
      }, ["전체선택"]),
      el("button", {
        class: "btn",
        onclick: () => {
          const any = Object.values(selections).some(Boolean);
          if (!any) return alert("내보낼 항목을 하나 이상 선택해 주세요.");
          try {
            exportSelectedToExcel(selections);
            backdrop.remove();
          } catch (err) {
            console.error(err);
            alert("엑셀 내보내기 실패: XLSX 라이브러리 로드 여부 / 브라우저 다운로드 권한을 확인해 주세요.");
          }
        }
      }, ["내보내기(Excel)"])
    ]);

    modal.appendChild(head);
    modal.appendChild(body);
    modal.appendChild(foot);

    backdrop.addEventListener("click", (e) => {
      if (e.target === backdrop) backdrop.remove();
    });

    backdrop.appendChild(modal);
    document.body.appendChild(backdrop);

    const onKey = (e) => {
      if (e.key === "Escape") {
        backdrop.remove();
      }
    };
    window.addEventListener("keydown", onKey, { capture: true, once: true });
  }

  function exportSelectedToExcel(sel) {
    if (typeof XLSX === "undefined" || !XLSX?.utils) {
      throw new Error("XLSX not loaded");
    }

    if (state.activeTab === "steel" || state.activeTab === "steel_sub" || state.activeTab === "support") {
      recomputeSection(state.activeTab);
    }

    const wb = XLSX.utils.book_new();

    if (sel.code) {
      const rows = (state.codeMaster || []).map(r => ({
        code: r.code ?? "",
        name: r.name ?? "",
        spec: r.spec ?? "",
        unit: r.unit ?? "",
        surcharge: r.surcharge ?? "",
        convUnit: r.convUnit ?? "",
        convFactor: r.convFactor ?? "",
        note: r.note ?? "",
      }));
      const ws = XLSX.utils.json_to_sheet(rows, { skipHeader: false });
      XLSX.utils.book_append_sheet(wb, ws, "Codes");
    }

    if (sel.steel) appendCalcTabSheet(wb, "steel", "Steel");
    if (sel.steel_sub) appendCalcTabSheet(wb, "steel_sub", "Steel_Sub");
    if (sel.support) appendCalcTabSheet(wb, "support", "Support");

    const fileName = `FIN_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  }

  function appendCalcTabSheet(wb, tabId, sheetName) {
    const bucket = state[tabId];
    if (!bucket || !Array.isArray(bucket.sections)) return;

    const prev = bucket.activeSection;
    const out = [];

    for (let sIdx = 0; sIdx < bucket.sections.length; sIdx++) {
      bucket.activeSection = sIdx;
      recomputeSection(tabId);

      const sec = bucket.sections[sIdx];
      const sectionName = sec.name ?? `구분 ${sIdx + 1}`;
      const count = sec.count ?? "";

      for (const v of (sec.vars || [])) {
        if (!v.key && !v.expr && !v.note) continue;
        out.push({
          type: "VAR",
          sectionName,
          count,
          key: v.key ?? "",
          expr: v.expr ?? "",
          value: v.value ?? 0,
          note: v.note ?? "",
        });
      }

      (sec.rows || []).forEach((r, i) => {
        const hasAny =
          (r.code || r.formula || r.value || r.converted || r.name || r.spec || r.unit || r.surchargePct != null);
        if (!hasAny) return;

        out.push({
          type: "ROW",
          sectionName,
          count,
          no: i + 1,
          code: r.code ?? "",
          name: r.name ?? "",
          spec: r.spec ?? "",
          unit: r.unit ?? "",
          formula: r.formula ?? "",
          value: r.value ?? 0,
          surchargePct: r.surchargePct ?? "",
          convUnit: r.convUnit ?? "",
          convFactor: r.convFactor ?? "",
          converted: r.converted ?? 0,
          note: r.note ?? "",
        });
      });
    }

    bucket.activeSection = prev;

    const ws = XLSX.utils.json_to_sheet(out, { skipHeader: false });
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  async function importExcelToCodes(file) {
    if (typeof XLSX === "undefined" || !XLSX?.read) {
      throw new Error("XLSX not loaded");
    }

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheetNames = wb.SheetNames || [];
    const pickSheet = (candidates) => {
      for (const cand of candidates) {
        const hit = sheetNames.find(n => String(n).trim().toLowerCase() === String(cand).trim().toLowerCase());
        if (hit) return hit;
      }
      return null;
    };

    const sn =
      pickSheet(["Codes", "Code", "코드", "CODE"]) ||
      (sheetNames[0] || null);

    if (!sn) throw new Error("No sheet");
    const ws = wb.Sheets[sn];
    if (!ws) throw new Error("Sheet missing");

    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

    const get = (row, keys) => {
      for (const k of keys) {
        if (row[k] !== undefined) return row[k];
      }
      return "";
    };

    const next = [];
    for (const row of json) {
      const code = String(get(row, ["code", "Code", "CODE", "코드"])).trim();
      if (!code) continue;

      const name = String(get(row, ["name", "품명", "Product name", "품명\n(Product name)"]));
      const spec = String(get(row, ["spec", "규격", "Specifications", "규격\n(Specifications)"]));
      const unit = String(get(row, ["unit", "단위", "Unit", "단위\n(unit)"]));
      const note = String(get(row, ["note", "비고", "Note", "비고\n(Note)"]));

      const surchargeRaw = get(row, ["surcharge", "할증", "할증\n(surcharge)"]);
      const convUnit = String(get(row, ["convUnit", "환산단위", "Conversion unit", "환산단위\n(Conversion unit)"]));
      const convFactorRaw = get(row, ["convFactor", "환산계수", "Conversion factor", "환산계수\n(Conversion factor)"]);

      const surcharge = (String(surchargeRaw).trim() === "") ? null : Number(surchargeRaw);
      const convFactor = (String(convFactorRaw).trim() === "") ? null : Number(convFactorRaw);

      next.push({
        code: code.toUpperCase(),
        name,
        spec,
        unit,
        surcharge: Number.isFinite(surcharge) ? surcharge : null,
        convUnit,
        convFactor: Number.isFinite(convFactor) ? convFactor : null,
        note,
      });
    }

    if (!next.length) {
      throw new Error("No valid rows");
    }

    state.codeMaster = next;
    saveState();
  }

   /***************
 * ✅ Project UI (Open/Add/Edit/Delete/Select)
 * - 프로젝트 선택 전: 주요 버튼 잠금
 * - 프로젝트 선택 후: 프로젝트별 state 로드/저장
 *
 * (필요 id)
 * - projectName, projectCode
 * - btnProjectOpen, btnProjectClose, btnProjectAdd
 * - projectModal, projectList
 * - btnOpenPicker, btnExport, btnReset, fileImport
 * - btnImportWrap (fileImport를 감싸는 label/div)
 * - btnHelp
 ***************/

function setTopButtonsEnabled(enabled) {
  const btnOpen = document.getElementById("btnOpenPicker");
  const btnExport = document.getElementById("btnExport");
  const btnReset = document.getElementById("btnReset");
  const fileImport = document.getElementById("fileImport");
  const btnImportWrap = document.getElementById("btnImportWrap"); // ✅ label wrapper 잠금 처리

  if (btnOpen) btnOpen.disabled = !enabled;
  if (btnExport) btnExport.disabled = !enabled;
  if (btnReset) btnReset.disabled = !enabled;
  if (fileImport) fileImport.disabled = !enabled;

  // ✅ label wrapper는 disabled가 안 먹어서 스타일로 잠금
  if (btnImportWrap) {
    btnImportWrap.style.opacity = enabled ? "1" : "0.45";
    btnImportWrap.style.pointerEvents = enabled ? "auto" : "none";
    btnImportWrap.setAttribute("aria-disabled", enabled ? "false" : "true");
  }

  // 도움말/프로젝트 버튼은 항상 사용 가능
  const help = document.getElementById("btnHelp");
  if (help) help.disabled = false;

  const btnProjectOpen = document.getElementById("btnProjectOpen");
  if (btnProjectOpen) btnProjectOpen.disabled = false;
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
  modal.hidden = false;
  renderProjectList();
}

function closeProjectModal() {
  const modal = document.getElementById("projectModal");
  if (!modal) return;
  modal.hidden = true;
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
      // ✅ 현재 active인 프로젝트는 저장 후 메타 변경(안전)
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

  // ✅ 신규 프로젝트는 DEFAULT_STATE로 저장해두고 바로 선택
  ProjectStore.saveProjectState(pid, deepClone(DEFAULT_STATE));
  selectProject(pid);

  renderProjectList();
}

function deleteProject(projectId) {
  // ✅ active 삭제면 먼저 active 저장/해제
  if (projectId === activeProjectId) {
    // 저장은 굳이 안 해도 되지만(삭제니까) 혹시 모를 안전망
    try { saveProjectState(activeProjectId); } catch {}
    activeProjectId = "";
    ProjectStore.saveActiveId("");
  }

  projectIndex.projects = projectIndex.projects.filter(p => p.id !== projectId);
  saveProjectIndex(projectIndex);
  ProjectStore.deleteProject(projectId);

  updateProjectHeaderUI();
  renderProjectList();

  // ✅ active가 없으면 화면은 기본 상태로
  if (!activeProjectId) {
    state = deepClone(DEFAULT_STATE);
    render();
  }
}

function selectProject(projectId) {
  const meta = projectIndex.projects.find(p => p.id === projectId);
  if (!meta) return;

  // ✅ 전환 전: 기존 active 프로젝트 state 저장
  if (activeProjectId) saveProjectState(activeProjectId);

  activeProjectId = projectId;
  ProjectStore.saveActiveId(projectId);

  meta.updatedAt = Date.now();
  saveProjectIndex(projectIndex);

  // ✅ 새 프로젝트 state 로드
  state = loadProjectState(activeProjectId);

  updateProjectHeaderUI();
  render();
}

let __boundTopOnce = false;

function bindTopButtons() {
  if (__boundTopOnce) return;
  __boundTopOnce = true;

  const btnHelp = document.getElementById("btnHelp");
  const btnOpen = document.getElementById("btnOpenPicker");
  const btnExport = document.getElementById("btnExport");
  const btnReset = document.getElementById("btnReset");
  const fileImport = document.getElementById("fileImport");

  const btnProjectOpen = document.getElementById("btnProjectOpen");
  const btnProjectClose = document.getElementById("btnProjectClose");
  const btnProjectAdd = document.getElementById("btnProjectAdd");
  const projectModal = document.getElementById("projectModal");

  if (btnHelp) btnHelp.onclick = openHelpWindow;
  if (btnOpen) btnOpen.onclick = openCodePicker;
  if (btnExport) btnExport.onclick = openExcelExportModal;

  if (btnProjectOpen) btnProjectOpen.onclick = openProjectModal;
  if (btnProjectClose) btnProjectClose.onclick = closeProjectModal;
  if (btnProjectAdd) btnProjectAdd.onclick = createProject;

  if (projectModal) {
    projectModal.addEventListener("click", (e) => {
      if (e.target === projectModal) closeProjectModal();
    });
  }

  // ✅ Excel import (Codes 시트 → codeMaster 갱신)
  if (fileImport) fileImport.onchange = async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;

    try {
      await importExcelToCodes(f);
      alert("가져오기(Excel) 완료: codeMaster(코드)가 갱신되었습니다.");
      render();
    } catch (err) {
      console.error(err);
      alert("가져오기(Excel) 실패: XLSX 로드 여부 / Codes 시트 존재 여부를 확인해 주세요.");
    } finally {
      e.target.value = "";
    }
  };

  // ✅ Reset: 현재 프로젝트만 초기화
  if (btnReset) btnReset.onclick = () => {
    if (!activeProjectId) return alert("먼저 프로젝트를 선택하세요.");
    if (!confirm("정말 초기화할까요? (현재 프로젝트의 로컬 저장 데이터가 초기화됩니다)")) return;

    ProjectStore.saveProjectState(activeProjectId, deepClone(DEFAULT_STATE));
    state = loadProjectState(activeProjectId);
    render();
  };
}

/***************
 * Init
 ***************/
updateProjectHeaderUI();
render();


  function applyPanelStickyTop() {
    const root = document.documentElement;
    const isCalcTab = (state.activeTab === "steel" || state.activeTab === "steel_sub" || state.activeTab === "support");
    root.style.setProperty("--panelStickyTop", isCalcTab ? "var(--stickyWithTopSplitTop)" : "var(--stickyBaseTop)");
  }

  function render() {
    // ✅ (v13.2b) topSplit 높이 먼저 적용
    applyTopSplitH();

         // ✅ 탭이 바뀔 때만 멀티선택 해제
if (__calcMulti.active && __calcMulti.tabId !== state.activeTab) {
  __calcMultiClear();
}



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

    raf2(() => {
  // ✅ (PATCH) zoom(--uiScale)일 때 렌더 직후 view 높이 보정이 가장 중요

  updateViewFillHeight();     
  updateStickyVars();
  applyPanelStickyTop();
  updateScrollHeights();

  if (__pendingSectionFocus && __pendingSectionFocus.tabId === state.activeTab) {
    const { tabId, index } = __pendingSectionFocus;
    __pendingSectionFocus = null;

    const list = document.querySelector(`.section-list[data-tab="${tabId}"]`);
    if (list) {
      const items = [...list.querySelectorAll(".section-item")];
      const idx = clamp(Number(index || 0), 0, items.length - 1);
      const target = items[idx];
      if (target) {
        safeFocus(target);
        try { target.scrollIntoView({ block: "nearest" }); } catch {}
      } else {
        safeFocus(list);
      }
    }
  }
});

  }

  function renderSummaryTabByCodeOrder(srcTabId, title) {
    const bucket = state[srcTabId];

    const orderMap = new Map();
    (state.codeMaster || []).forEach((cm, idx) => {
      const c = String(cm.code || "").trim().toUpperCase();
      if (c) orderMap.set(c, idx);
    });

    const map = new Map();
    const prev = bucket.activeSection;

    for (let sIdx = 0; sIdx < bucket.sections.length; sIdx++) {
      bucket.activeSection = sIdx;
      recomputeSection(srcTabId);

      const sec = bucket.sections[sIdx];

      let countMul = 1;
      const rawCount = (sec.count ?? "").toString().trim();
      if (rawCount === "") countMul = 1;
      else {
        const n = Number(rawCount);
        countMul = Number.isFinite(n) ? n : 1;
      }

      for (const r of sec.rows) {
        const code = String(r.code || "").trim().toUpperCase();
        if (!code) continue;

        const name = r.name || "";
        const spec = r.spec || "";

        const baseQty = (Number(r.value) || 0) * countMul;
        const mul = (Number(r.surchargeMul) || 1);
        const afterQty = (Number(r.value) || 0) * mul * countMul;

        const convUnit = String(r.convUnit || "").trim();
        const convFactorNum = Number(r.convFactor);
        const hasConv = convUnit !== "" && Number.isFinite(convFactorNum) && convFactorNum !== 0;

        const unitShown = hasConv ? convUnit : (r.unit || "");
        const preShown  = hasConv ? (baseQty  * convFactorNum) : baseQty;
        const postShown = hasConv ? (afterQty * convFactorNum) : afterQty;

        const pct =
          (r.surchargePct == null || r.surchargePct === "" || !Number.isFinite(Number(r.surchargePct)))
            ? null
            : Number(r.surchargePct);

        if (!map.has(code)) {
          map.set(code, {
            code,
            name,
            spec,
            unit: unitShown,
            pre: 0,
            post: 0,
            pctSet: new Set(),
            unitSet: new Set(),
          });
        }

        const agg = map.get(code);
        agg.pre += preShown;
        agg.post += postShown;
        agg.unitSet.add(unitShown || "");

        if (pct == null) agg.pctSet.add("__NULL__");
        else agg.pctSet.add(String(pct));
      }
    }

    bucket.activeSection = prev;
    saveState();

    const items = [...map.values()].sort((a, b) => {
      const ai = orderMap.has(a.code) ? orderMap.get(a.code) : Number.POSITIVE_INFINITY;
      const bi = orderMap.has(b.code) ? orderMap.get(b.code) : Number.POSITIVE_INFINITY;
      if (ai !== bi) return ai - bi;
      return a.code.localeCompare(b.code);
    });

    const panelHeader = el("div", { class: "panel-header sticky-head", dataset: { sticky: "panel" } }, [
      el("div", {}, [
        el("div", { class: "panel-title" }, [title]),
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
              const unitText = (x.unitSet && x.unitSet.size > 1) ? "혼합" : (x.unit || "");

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
                el("td", {}, [unitText]),
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
updateProjectHeaderUI();   // ✅ 최초 1회 (프로젝트 미선택이면 버튼 잠금)
render();

