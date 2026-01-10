/* app.js (v10+)
   - 상단(구분명 리스트 + 변수표) 영역을 산출표 위에 "고정(sticky)" 되는 구조로 렌더
   - 구분(Ctrl+F3) 추가 / 선택(↑↓, Enter) / 마우스 클릭
   - 변수표: 변수명(최대 3자) / 산식 / 값(자동) / 비고
     * "<...>" 주석은 계산에서 제외
     * 변수는 A~Z로 시작, A / AB / A1 / AB1 같은 조합 허용(최대 3자)
   - 철골/부자재/동바리 탭별로 "구분별" 산출행을 독립적으로 유지
   - 탭 명칭 변경:
     * "철골(Steel)" -> "철골"
     * "동바리(support)" -> "구조이기/동바리"
     * "동바리_집계" -> "구조이기/동바리_집계"
*/

(() => {
  "use strict";

  /* =========================
     DOM Helpers
     ========================= */
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

  function clamp(n, a, b) { return Math.max(a, Math.min(b, n)); }

  /* =========================
     Storage
     ========================= */
  const LS_KEY = "FIN_WEB_V10_STATE";

  function loadState() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }

  function saveState() {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(state));
    } catch {}
  }

  /* =========================
     Tabs (labels 변경 포함)
     ========================= */
  const TAB_DEFS = [
    { id: "code", label: "코드(Ctrl+.)", type: "code" },
    { id: "steel", label: "철골", type: "calc_steel" },
    { id: "steel_sum", label: "철골_집계", type: "summary_steel" },
    { id: "steel_aux", label: "철골_부자재", type: "calc_aux" },
    { id: "support", label: "구조이기/동바리", type: "calc_support" },
    { id: "support_sum", label: "구조이기/동바리_집계", type: "summary_support" },
  ];

  /* =========================
     State Model
     - 탭별 sections를 가짐
     - 각 section: name, count, vars[], rows[]
     ========================= */
  function makeDefaultVars() {
    // 기본 12행 (요청: 많이 보이게 / 설명문구 삭제는 렌더에서)
    return Array.from({ length: 12 }).map(() => ({
      key: "",
      expr: "",
      value: 0,
      note: ""
    }));
  }

  function makeDefaultRows(type) {
    // 산출표 기본 12행
    // steel/support/aux 공통 형태로 쓰되, 집계 탭은 사용 안 함.
    const base = Array.from({ length: 12 }).map(() => ({
      code: "",
      name: "",
      spec: "",
      unit: "",
      expr: "",
      value: 0,
      mult: "",
      conv: ""
    }));

    // 타입별 기본 단위 힌트 정도만 넣기(빈칸 유지)
    if (type === "calc_steel") {
      base.forEach(r => { r.unit = "M"; });
    } else if (type === "calc_aux") {
      base.forEach(r => { r.unit = "M2"; });
    } else if (type === "calc_support") {
      base.forEach(r => { r.unit = ""; });
    }
    return base;
  }

  function makeSection(name = "구분 1", count = "") {
    return {
      id: cryptoRandomId(),
      name,
      count,
      vars: makeDefaultVars(),
      // rows는 탭별로 따로 들고가므로 여기엔 공용으로 두지 않음
    };
  }

  function cryptoRandomId() {
    try {
      return crypto.getRandomValues(new Uint32Array(2)).join("-");
    } catch {
      return String(Date.now()) + "-" + Math.floor(Math.random() * 1e9);
    }
  }

  function makeDefaultState() {
    const tabs = {};
    for (const t of TAB_DEFS) {
      tabs[t.id] = {
        id: t.id,
        label: t.label,
        type: t.type,
        // 산출이 필요한 탭만 sections를 운용
        sections: (t.type.startsWith("calc_") ? [makeSection("1층 바닥 철골보", "1")] : []),
        activeSectionId: null,
        // 섹션별 산출행은 tab.sectionsRows[sectionId] 형태로 저장
        sectionsRows: {}
      };

      if (t.type.startsWith("calc_")) {
        const sec = tabs[t.id].sections[0];
        tabs[t.id].activeSectionId = sec.id;
        tabs[t.id].sectionsRows[sec.id] = makeDefaultRows(t.type);
      }
    }

    return {
      activeTabId: "steel",
      tabs
    };
  }

  let state = loadState() || makeDefaultState();

  /* =========================
     Sticky Top 계산 (topbar + tabs)
     ========================= */
  function updateStickyTopVar() {
    const topbar = $(".topbar");
    const tabs = $("#tabs");
    const topbarH = topbar ? topbar.getBoundingClientRect().height : 0;
    const tabsH = tabs ? tabs.getBoundingClientRect().height : 0;
    // topbar sticky(0) + tabs sticky(topbar 아래) → top-split이 그 아래에서 시작
    const stickyTop = Math.ceil(topbarH + tabsH);
    document.documentElement.style.setProperty("--stickyTop", `${stickyTop}px`);
  }

  /* =========================
     Parsing / Evaluation
     - "<...>" 제거
     - 변수명: [A-Z][A-Z0-9]{0,2}
     - 허용문자: 숫자/연산자/괄호/공백/소수점/변수
     ========================= */
  function stripAngleComments(s) {
    if (!s) return "";
    return String(s).replace(/<[^>]*>/g, "");
  }

  function normalizeExpr(s) {
    return stripAngleComments(s).replace(/\s+/g, " ").trim();
  }

  function isValidVarName(v) {
    return /^[A-Z][A-Z0-9]{0,2}$/.test(v || "");
  }

  function safeEvalMath(expr, varMap) {
    const raw = normalizeExpr(expr);
    if (!raw) return 0;

    // 허용 토큰 체크 (숫자/연산자/괄호/점/공백/변수명)
    // 변수는 일단 토큰으로 치환 후 숫자식으로 eval
    const tokenized = raw.replace(/[A-Z][A-Z0-9]{0,2}/g, (m) => {
      const v = varMap[m];
      if (v == null || Number.isNaN(Number(v))) return "0";
      return String(Number(v));
    });

    // 허용 문자만 남았는지 검사
    if (!/^[0-9+\-*/().\s]+$/.test(tokenized)) return 0;

    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(`"use strict"; return (${tokenized});`);
      const out = fn();
      const num = Number(out);
      if (!Number.isFinite(num)) return 0;
      return num;
    } catch {
      return 0;
    }
  }

  function buildVarMap(section) {
    // 1) 먼저 값이 직접 숫자면 그걸 우선(여기서는 expr로 산출)
    // 2) 순차적으로 여러 번 돌려서 의존성 완화(간단한 고정점 반복)
    const map = {};
    for (const r of section.vars) {
      if (isValidVarName(r.key)) map[r.key] = 0;
    }

    // 반복 계산
    for (let iter = 0; iter < 4; iter++) {
      for (const r of section.vars) {
        if (!isValidVarName(r.key)) continue;
        const v = safeEvalMath(r.expr, map);
        map[r.key] = v;
      }
    }
    return map;
  }

  /* =========================
     Rendering
     ========================= */
  const tabsEl = $("#tabs");
  const viewEl = $("#view");

  function renderTabs() {
    tabsEl.innerHTML = "";
    for (const t of TAB_DEFS) {
      const btn = el("button", {
        class: "tab" + (state.activeTabId === t.id ? " active" : ""),
        type: "button",
        dataset: { tabId: t.id },
        onclick: () => {
          state.activeTabId = t.id;
          saveState();
          render();
        }
      }, t.label);
      tabsEl.appendChild(btn);
    }
  }

  function getActiveTab() {
    return state.tabs[state.activeTabId];
  }

  function getActiveSection(tab) {
    if (!tab || !tab.sections?.length) return null;
    const id = tab.activeSectionId || tab.sections[0].id;
    return tab.sections.find(s => s.id === id) || tab.sections[0];
  }

  /* ===== Top (구분명 + 변수표) ===== */
  function renderTopSplit(tab) {
    // 상단 고정 래퍼
    const wrapper = el("div", { class: "top-split" });
    const grid = el("div", { class: "calc-layout" });

    const sectionBox = renderSectionBox(tab);
    const varBox = renderVarBox(tab);

    grid.append(sectionBox, varBox);
    wrapper.append(grid);
    return wrapper;
  }

  function renderSectionBox(tab) {
    const section = getActiveSection(tab);

    const box = el("div", { class: "rail-box", id: "sectionBox" },
      el("div", { class: "rail-title" }, "구분명 리스트 (↑/↓ 이동)")
    );

    const list = el("div", { class: "section-list", id: "sectionList", tabindex: "0" });

    tab.sections.forEach((s, idx) => {
      const item = el("div", {
        class: "section-item" + (section && s.id === section.id ? " active" : ""),
        tabindex: "0",
        dataset: { secId: s.id, idx: String(idx) },
        onclick: () => {
          tab.activeSectionId = s.id;
          // 섹션별 rows 없으면 생성
          if (!tab.sectionsRows[s.id] && tab.type.startsWith("calc_")) {
            tab.sectionsRows[s.id] = makeDefaultRows(tab.type);
          }
          saveState();
          render();
          // 선택된 아이템 포커스
          requestAnimationFrame(() => {
            const newly = $(`.section-item[data-sec-id="${cssEscape(s.id)}"]`);
            if (newly) newly.focus();
          });
        },
        onkeydown: (e) => {
          // 리스트 항목에서만 ↑↓로 이동
          if (e.key === "ArrowUp" || e.key === "ArrowDown") {
            e.preventDefault();
            const items = $$(".section-item", list);
            const cur = items.indexOf(e.currentTarget);
            const next = clamp(cur + (e.key === "ArrowDown" ? 1 : -1), 0, items.length - 1);
            items[next]?.focus();
          }
          if (e.key === "Enter") {
            e.preventDefault();
            e.currentTarget.click();
          }
        }
      },
        el("div", { class: "name" }, s.name || `구분 ${idx + 1}`),
        el("div", { class: "meta" }, `개소: ${s.count ?? ""}`),
        el("div", { class: "meta" }, "선택")
      );
      list.appendChild(item);
    });

    // 편집 행(구분명 / 개소 / 저장)
    const editor = el("div", { class: "section-editor" });

    const inpName = el("input", {
      type: "text",
      placeholder: "구분명(예: 2층 바닥 철골보)",
      value: section?.name ?? "",
      oninput: (e) => {
        const s = getActiveSection(tab);
        if (!s) return;
        s.name = e.target.value;
        saveState();
        // 리스트 텍스트 즉시 반영은 전체 렌더가 깔끔
      }
    });

    const inpCount = el("input", {
      type: "text",
      placeholder: "개소(예: 0,1,2...)",
      value: section?.count ?? "",
      oninput: (e) => {
        const s = getActiveSection(tab);
        if (!s) return;
        s.count = e.target.value;
        saveState();
      }
    });

    const btnSave = el("button", {
      class: "smallbtn",
      type: "button",
      onclick: () => {
        saveState();
        render(); // 리스트 텍스트/개소 즉시 반영
      }
    }, "저장");

    editor.append(inpName, inpCount, btnSave);

    const btnRow = el("div", { class: "row-actions", style: "margin-top:8px;" },
      el("button", {
        class: "smallbtn",
        type: "button",
        onclick: () => addSection(tab)
      }, "구분 추가 (Ctrl+F3)"),
      el("button", {
        class: "smallbtn",
        type: "button",
        onclick: () => deleteActiveSection(tab)
      }, "구분 삭제")
    );

    box.append(list, editor, btnRow);

    // Ctrl+F3 (구분추가) - 문서 전역 단축키
    return box;
  }

  function addSection(tab) {
    if (!tab.type.startsWith("calc_")) return;
    const nextIdx = tab.sections.length + 1;
    const sec = makeSection(`구분 ${nextIdx}`, "");
    tab.sections.push(sec);
    tab.activeSectionId = sec.id;
    tab.sectionsRows[sec.id] = makeDefaultRows(tab.type);
    saveState();
    render();
    requestAnimationFrame(() => {
      // 새로 추가된 항목에 포커스
      const items = $$(".section-item", $("#sectionBox"));
      items[items.length - 1]?.focus();
    });
  }

  function deleteActiveSection(tab) {
    if (!tab.type.startsWith("calc_")) return;
    if (tab.sections.length <= 1) return; // 최소 1개 유지
    const cur = getActiveSection(tab);
    if (!cur) return;
    const idx = tab.sections.findIndex(s => s.id === cur.id);
    tab.sections.splice(idx, 1);
    delete tab.sectionsRows[cur.id];

    const next = tab.sections[clamp(idx, 0, tab.sections.length - 1)];
    tab.activeSectionId = next.id;
    saveState();
    render();
  }

  function renderVarBox(tab) {
    const sec = getActiveSection(tab);

    const box = el("div", { class: "rail-box", id: "varBox" },
      el("div", { class: "rail-title" }, "변수표 (A, AB, A1, AB1... 최대 3자)")
    );

    const wrap = el("div", { class: "var-tablewrap", id: "varTableWrap", tabindex: "0" });

    const table = el("table", { class: "var-table" });
    const thead = el("thead", {}, el("tr", {},
      el("th", {}, "변수"),
      el("th", {}, "산식"),
      el("th", {}, "값"),
      el("th", {}, "비고")
    ));
    const tbody = el("tbody");

    const varMap = sec ? buildVarMap(sec) : {};

    (sec?.vars || []).forEach((r, rowIdx) => {
      // 계산된 값 갱신(표시용)
      if (sec && isValidVarName(r.key)) {
        r.value = Number(varMap[r.key] ?? 0);
      } else {
        r.value = 0;
      }

      const tr = el("tr", {},

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          inputmode: "text",
          autocomplete: "off",
          spellcheck: "false",
          placeholder: "예: A / AB / A1",
          value: r.key ?? "",
          "data-scope": "var",
          "data-row": String(rowIdx),
          "data-col": "key",
          oninput: (e) => {
            if (!sec) return;
            // 대문자/숫자만 + 최대3자
            let v = String(e.target.value || "").toUpperCase();
            v = v.replace(/[^A-Z0-9]/g, "").slice(0, 3);
            e.target.value = v;
            r.key = v;
            saveState();
          },
          onkeydown: (e) => {
            // ✅ 여기서 ArrowKey를 가로채지 않음 (커서 이동 보장)
            // Enter만 다음 셀로 이동
            if (e.key === "Enter") {
              e.preventDefault();
              focusVarCell(rowIdx, "expr");
            }
          },
          onblur: () => {
            if (!sec) return;
            saveState();
            // 값 재계산/반영
            render(); // 간단하게 전체 갱신
            requestAnimationFrame(() => focusVarCell(rowIdx, "key"));
          }
        }))),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          placeholder: "예: (A+0.5)*2  (<...> 주석)",
          value: r.expr ?? "",
          "data-scope": "var",
          "data-row": String(rowIdx),
          "data-col": "expr",
          oninput: (e) => {
            if (!sec) return;
            r.expr = e.target.value;
            saveState();
          },
          onkeydown: (e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              // Enter 시 즉시 계산/반영
              saveState();
              render();
              requestAnimationFrame(() => focusVarCell(rowIdx, "expr"));
            }
          }
        }))),

        el("td", {}, el("input", {
          class: "cell readonly",
          type: "text",
          readonly: true,
          value: formatNumber(r.value)
        })),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          placeholder: "비고",
          value: r.note ?? "",
          "data-scope": "var",
          "data-row": String(rowIdx),
          "data-col": "note",
          oninput: (e) => {
            if (!sec) return;
            r.note = e.target.value;
            saveState();
          },
          onkeydown: (e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              // 다음 행 변수명으로 이동
              focusVarCell(clamp(rowIdx + 1, 0, (sec.vars.length - 1)), "key");
            }
          }
        })))
      );

      tbody.appendChild(tr);
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    box.appendChild(wrap);

    return box;
  }

  function focusVarCell(rowIdx, col) {
    const q = `input[data-scope="var"][data-row="${rowIdx}"][data-col="${col}"]`;
    const target = $(q, document);
    if (target) {
      target.focus();
      target.select?.();
    }
  }

  /* ===== Calc Table (철골/부자재/동바리) ===== */
  function renderCalcTab(tab) {
    const sec = getActiveSection(tab);
    if (!sec) return el("div", { class: "panel" }, "구분을 먼저 생성해 주세요.");

    const rows = tab.sectionsRows[sec.id] || (tab.sectionsRows[sec.id] = makeDefaultRows(tab.type));
    const varMap = buildVarMap(sec);

    const panel = el("div", { class: "panel" });

    // 헤더
    panel.appendChild(
      el("div", { class: "panel-header" },
        el("div", {},
          el("div", { class: "panel-title" }, tab.label),
          el("div", { class: "panel-desc" },
            "구분(↑/↓) 이동 → 해당 구분의 변수/산출표 전환 | 산출식 Enter 계산 | Ctrl+. 코드선택 | Ctrl+F3 구분추가"
          )
        ),
        el("div", { class: "row-actions" },
          el("button", {
            class: "smallbtn",
            type: "button",
            onclick: () => addCalcRows(tab, sec.id, 10)
          }, "+10행")
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
      // 계산값 표시
      r.value = safeEvalMath(r.expr, varMap);

      const tr = el("tr", {},

        el("td", {}, String(i + 1)),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          value: r.code ?? "",
          placeholder: "코드 입력",
          oninput: (e) => { r.code = e.target.value; saveState(); }
        })),

        el("td", {}, el("input", {
          class: "cell readonly",
          type: "text",
          value: r.name ?? "",
          readonly: true
        })),

        el("td", {}, el("input", {
          class: "cell readonly",
          type: "text",
          value: r.spec ?? "",
          readonly: true
        })),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          value: r.unit ?? "",
          oninput: (e) => { r.unit = e.target.value; saveState(); }
        })),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          value: r.expr ?? "",
          placeholder: "예: (A+0.5)*2  (<...>는 주석)",
          oninput: (e) => { r.expr = e.target.value; saveState(); },
          onkeydown: (e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              saveState();
              render(); // 즉시 계산 반영
              // 포커스 유지
              requestAnimationFrame(() => {
                const all = $$(".cell", tbody);
                // 현재 입력 그대로 유지 시도
                e.target.focus();
              });
            }
          }
        })),

        el("td", {}, el("input", {
          class: "cell readonly",
          type: "text",
          readonly: true,
          value: formatNumber(r.value)
        })),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          value: r.mult ?? "",
          oninput: (e) => { r.mult = e.target.value; saveState(); }
        })),

        el("td", {}, el("input", {
          class: "cell",
          type: "text",
          value: r.conv ?? "",
          oninput: (e) => { r.conv = e.target.value; saveState(); }
        }))
      );

      tbody.appendChild(tr);
    });

    table.append(thead, tbody);
    wrap.appendChild(table);
    panel.appendChild(wrap);
    return panel;
  }

  function addCalcRows(tab, secId, n) {
    const rows = tab.sectionsRows[secId] || (tab.sectionsRows[secId] = makeDefaultRows(tab.type));
    const extra = makeDefaultRows(tab.type).slice(0, n);
    tab.sectionsRows[secId] = rows.concat(extra);
    saveState();
    render();
  }

  /* ===== Summary Tabs (간단 표시) ===== */
  function renderSummaryTab(tab) {
    const panel = el("div", { class: "panel" },
      el("div", { class: "panel-header" },
        el("div", {},
          el("div", { class: "panel-title" }, tab.label),
          el("div", { class: "panel-desc" }, "※ 집계 탭은 추후 로직 연결(현재는 구분/산출 데이터를 기반으로 합계를 계산하도록 확장 예정)")
        )
      ),
      el("div", { style: "color:rgba(0,0,0,.65); font-size:12px; padding:8px 4px;" },
        "현재 버전은 UI/구조 및 구분-변수-산출 연동을 우선 구성했습니다. 집계 규칙(할증후수량 합계 등)은 다음 단계에서 탭별로 반영합니다."
      )
    );
    return panel;
  }

  function renderCodeTab() {
    const panel = el("div", { class: "panel" },
      el("div", { class: "panel-header" },
        el("div", {},
          el("div", { class: "panel-title" }, "코드(Ctrl+.)"),
          el("div", { class: "panel-desc" }, "코드 선택 팝업/엑셀 연동은 별도 연결(현재는 입력 유지 및 저장/불러오기만 활성)")
        )
      ),
      el("div", { style: "font-size:12px; color:rgba(0,0,0,.65);" },
        "※ 코드 선택 창(Ctrl+.) 기능은 기존 구현체가 있다면 그 로직을 이 파일에 붙여넣어 연결하면 됩니다."
      )
    );
    return panel;
  }

  function renderView() {
    viewEl.innerHTML = "";
    const tab = getActiveTab();

    // 산출 탭이면 상단 고정(구분+변수) 먼저
    if (tab.type.startsWith("calc_")) {
      viewEl.appendChild(renderTopSplit(tab));
      viewEl.appendChild(renderCalcTab(tab));
      return;
    }

    // 그 외 탭
    if (tab.type === "code") {
      viewEl.appendChild(renderCodeTab());
      return;
    }

    if (tab.type.startsWith("summary_")) {
      viewEl.appendChild(renderSummaryTab(tab));
      return;
    }

    viewEl.appendChild(el("div", { class: "panel" }, "준비중"));
  }

  function render() {
    renderTabs();
    renderView();
    updateStickyTopVar();
  }

  /* =========================
     Global Hotkeys
     ========================= */
  document.addEventListener("keydown", (e) => {
    // 입력 중이면 대부분 단축키 무시(단, Ctrl 계열은 허용)
    const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : "";
    const isTyping = (tag === "input" || tag === "textarea" || e.target?.isContentEditable);

    // Ctrl+F3 : 구분 추가
    if (e.ctrlKey && e.key === "F3") {
      e.preventDefault();
      const tab = getActiveTab();
      if (tab.type.startsWith("calc_")) addSection(tab);
      return;
    }

    // 섹션 리스트 이동(↑↓) : "입력 중"이 아닐 때만
    if (!isTyping && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
      const tab = getActiveTab();
      if (!tab.type.startsWith("calc_")) return;

      const items = $$(".section-item", $("#sectionBox"));
      if (!items.length) return;

      const curSec = getActiveSection(tab);
      const curIdx = tab.sections.findIndex(s => s.id === curSec?.id);
      const nextIdx = clamp(curIdx + (e.key === "ArrowDown" ? 1 : -1), 0, tab.sections.length - 1);
      const nextSec = tab.sections[nextIdx];
      if (!nextSec) return;

      e.preventDefault();
      tab.activeSectionId = nextSec.id;
      if (!tab.sectionsRows[nextSec.id]) tab.sectionsRows[nextSec.id] = makeDefaultRows(tab.type);
      saveState();
      render();
      requestAnimationFrame(() => {
        const focusItem = $(`.section-item[data-sec-id="${cssEscape(nextSec.id)}"]`);
        focusItem?.focus();
      });
      return;
    }
  }, { capture: true });

  /* =========================
     Export / Import / Reset
     ========================= */
  $("#btnExport")?.addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = el("a", { href: url, download: "FIN_WEB_export.json" });
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  $("#fileImport")?.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const obj = JSON.parse(text);
      // 최소 검증
      if (!obj || !obj.tabs) throw new Error("Invalid");
      state = obj;
      saveState();
      render();
    } catch {
      alert("가져오기 실패: JSON 형식을 확인해주세요.");
    } finally {
      e.target.value = "";
    }
  });

  $("#btnReset")?.addEventListener("click", () => {
    if (!confirm("초기화 하시겠습니까? (저장된 데이터가 삭제됩니다)")) return;
    state = makeDefaultState();
    saveState();
    render();
  });

  // 코드 선택 (Ctrl+.) - 현재는 버튼만 유지 (기존 팝업 연결 지점)
  $("#btnOpenPicker")?.addEventListener("click", () => {
    alert("코드 선택 창(Ctrl+.) 로직은 기존 구현을 app.js에 연결해 주세요.");
  });

  /* =========================
     Utils
     ========================= */
  function formatNumber(n) {
    const num = Number(n);
    if (!Number.isFinite(num)) return "0";
    // 소수점 너무 길면 정리
    const s = Math.abs(num) < 1e-9 ? "0" : String(Math.round(num * 1000000) / 1000000);
    return s;
  }

  function cssEscape(s) {
    // querySelector용 최소 escape
    return String(s).replace(/"/g, '\\"');
  }

  /* =========================
     Boot
     ========================= */
  window.addEventListener("resize", () => updateStickyTopVar());
  window.addEventListener("load", () => updateStickyTopVar());

  render();
})();
