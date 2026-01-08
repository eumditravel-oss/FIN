/* =========================================================
   FIN 산출자료 (Web)
   - 탭: 코드 / 철골 / 철골집계 / 부자재 / 동바리 / 동바리집계
   - 저장: localStorage
   - 엑셀 로직 반영:
     철골/부자재: 환산수량 = (K가 비었거나 0이면 E, 아니면 E*K)
                 할증후수량 = (I가 비었거나 0이면 환산수량, 아니면 환산수량*I)
     철골 집계: "할증후수량" 합계
     동바리 집계: 엑셀은 E(물량) 합계로 되어 있어 그대로 반영
   ========================================================= */

const STORAGE_KEY = "FIN_WEB_V1";

/** 엑셀에서 읽힌 코드 마스터(기본 47개) */
const SEED_CODES = [
  {"code":"A0SM355150","name":"RH형강 / SM355","spec":"150*150*7*10","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0315,"note":""},
  {"code":"A0SM355200","name":"RH형강 / SM355","spec":"200*100*5.5*8","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0213,"note":""},
  {"code":"A0SM355201","name":"RH형강 / SM355","spec":"200*200*8*12","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355250","name":"RH형강 / SM355","spec":"250*150*6*9","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355300","name":"RH형강 / SM355","spec":"300*150*6.5*9","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355301","name":"RH형강 / SM355","spec":"300*300*10*15","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355350","name":"RH형강 / SM355","spec":"350*350*12*19","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355400","name":"RH형강 / SM355","spec":"400*400*13*21","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355450","name":"RH형강 / SM355","spec":"450*450*14*23","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355500","name":"RH형강 / SM355","spec":"500*500*16*25","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355550","name":"RH형강 / SM355","spec":"550*550*17*27","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355600","name":"RH형강 / SM355","spec":"600*600*19*30","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355650","name":"RH형강 / SM355","spec":"650*650*21*33","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355700","name":"RH형강 / SM355","spec":"700*700*24*38","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355750","name":"RH형강 / SM355","spec":"750*750*26*42","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355800","name":"RH형강 / SM355","spec":"800*800*28*46","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM355900","name":"RH형강 / SM355","spec":"900*300*16*28","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"A0SM3551000","name":"RH형강 / SM355","spec":"1000*500*20*35","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},

  {"code":"B0H398200","name":"H형강 / SS275","spec":"H-398*199*7*11","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.0656,"note":""},
  {"code":"B0H400200","name":"H형강 / SS275","spec":"H-400*200*8*13","unit":"M","surcharge":7,"conv_unit":"TON","conv_factor":0.066,"note":""},

  {"code":"C0PLT","name":"PLATE / SS275","spec":"t= (사용자 입력)","unit":"M2","surcharge":7,"conv_unit":"TON","conv_factor":"","note":"환산계수는 사용자 입력 가능"},
  {"code":"D0BOLT","name":"BOLT","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"E0WELD","name":"WELD","spec":"(사용자 입력)","unit":"M","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"F0PAINT","name":"PAINT","spec":"(사용자 입력)","unit":"M2","surcharge":"","conv_unit":"","conv_factor":"","note":""},

  {"code":"S0SUPPORT","name":"동바리(서포트)","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"S0PIPE","name":"파이프동바리","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"S0FRAME","name":"시스템동바리","spec":"(사용자 입력)","unit":"EA","surcharge":"","conv_unit":"","conv_factor":"","note":""},

  {"code":"Z0ETC001","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC002","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC003","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC004","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC005","name":"기타","spec":"(사용자 입력)","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},

  // 아래는 엑셀에 존재하는 코드 수(47개)를 맞추기 위한 기본 슬롯(빈 값 허용)
  {"code":"Z0ETC006","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC007","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC008","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC009","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC010","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC011","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC012","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC013","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC014","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC015","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC016","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC017","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC018","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC019","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""},
  {"code":"Z0ETC020","name":"기타","spec":"","unit":"","surcharge":"","conv_unit":"","conv_factor":"","note":""}
];

function makeEmptyCalcRow(){
  return {
    code: "",
    name: "",
    spec: "",
    unit: "",
    value: "",     // E (물량)
    formula: "",   // 산출식(문자)
    note: "",
    surchargeMul: "", // I (배수)
    convUnit: "",
    convFactor: "",
    convQty: 0,     // M (환산수량)
    finalQty: 0     // N (할증후수량)
  };
}

function makeState(){
  return {
    codes: SEED_CODES,
    steel: Array.from({length: 20}, makeEmptyCalcRow),
    steelSub: Array.from({length: 20}, makeEmptyCalcRow),
    support: Array.from({length: 20}, makeEmptyCalcRow),
  };
}

function loadState(){
  const raw = localStorage.getItem(STORAGE_KEY);
  if(!raw) return makeState();
  try{
    const s = JSON.parse(raw);
    // 최소 보정
    s.codes = Array.isArray(s.codes) ? s.codes : SEED_CODES;
    s.steel = Array.isArray(s.steel) ? s.steel : [];
    s.steelSub = Array.isArray(s.steelSub) ? s.steelSub : [];
    s.support = Array.isArray(s.support) ? s.support : [];
    return s;
  }catch{
    return makeState();
  }
}
function saveState(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function num(v){
  const x = (v ?? "").toString().trim();
  if(x === "") return 0;
  const n = Number(x.replaceAll(",",""));
  return Number.isFinite(n) ? n : 0;
}

function surchargeToMul(surchargePercent){
  const p = num(surchargePercent);
  if(!p) return "";
  // 엑셀은 I열에 "배수" 형태로 들어가므로 7%면 1.07
  return (1 + p/100);
}

function findCode(code){
  const key = (code ?? "").toString().trim();
  if(!key) return null;
  return state.codes.find(x => (x.code ?? "").toString().trim() === key) ?? null;
}

/** value가 "=10*100" 같은 형태면 계산해주기 */
function evalValueExpr(valueRaw){
  const s = (valueRaw ?? "").toString().trim();
  if(!s) return 0;

  if(s.startsWith("=")){
    // 매우 제한적으로: 숫자/연산자/괄호/공백만 허용
    const expr = s.slice(1);
    if(!/^[0-9+\-*/().\s,]+$/.test(expr)) return 0;
    try{
      // eslint-disable-next-line no-new-func
      const f = new Function(`return (${expr.replaceAll(",","")});`);
      const out = f();
      return Number.isFinite(out) ? out : 0;
    }catch{
      return 0;
    }
  }
  return num(s);
}

function recalcRow(row){
  const m = findCode(row.code);

  // VLOOKUP 반영: 코드가 있으면 자동 채움(사용자 수정도 가능하지만 기본은 코드 기반)
  if(m){
    row.name = m.name ?? "";
    row.spec = m.spec ?? "";
    row.unit = m.unit ?? "";
    // 할증(%) → 배수 자동
    const mul = surchargeToMul(m.surcharge);
    row.surchargeMul = mul === "" ? "" : mul;
    // 환산단위/계수
    row.convUnit = m.conv_unit ?? "";
    row.convFactor = m.conv_factor ?? "";
  } else {
    // 코드가 없으면 자동필드 비움(원하면 유지로 바꿀 수도 있음)
    row.name = row.name ?? "";
    row.spec = row.spec ?? "";
    row.unit = row.unit ?? "";
    row.surchargeMul = row.surchargeMul ?? "";
    row.convUnit = row.convUnit ?? "";
    row.convFactor = row.convFactor ?? "";
  }

  const E = evalValueExpr(row.value);
  // 산출식은 "FORMULATEXT"처럼 '=' 없이 보여주고 싶으면: valueRaw가 '='로 시작할 때 자동으로 formula에 채우기
  const vraw = (row.value ?? "").toString().trim();
  if(vraw.startsWith("=")){
    row.formula = vraw.slice(1);
  }

  const K = num(row.convFactor);
  const I = num(row.surchargeMul);

  // 엑셀 로직:
  // M = IF(K=0 or blank, E, E*K)
  row.convQty = (K === 0 ? E : E*K);

  // N = IF(I=0 or blank, M, M*I)
  row.finalQty = (I === 0 ? row.convQty : row.convQty * I);

  // 소수 3자리 올림/반올림은 화면에서 표시만 처리 (필요하면 여기서 고정 가능)
}

function recalcAll(){
  state.steel.forEach(recalcRow);
  state.steelSub.forEach(recalcRow);
  state.support.forEach(recalcRow);
}

function groupSum(rows, valueSelector){
  const map = new Map();
  for(const r of rows){
    const code = (r.code ?? "").toString().trim();
    if(!code) continue;
    const m = findCode(code);
    const key = code;
    const cur = map.get(key) ?? {
      code,
      name: m?.name ?? r.name ?? "",
      spec: m?.spec ?? r.spec ?? "",
      unit: m?.unit ?? r.unit ?? "",
      sum: 0
    };
    cur.sum += valueSelector(r);
    map.set(key, cur);
  }
  return Array.from(map.values()).sort((a,b)=>a.code.localeCompare(b.code));
}

function roundUp3(x){
  // 엑셀 ROUNDUP(...,3) 유사
  const n = Number(x);
  if(!Number.isFinite(n)) return 0;
  const p = 1000;
  return Math.ceil(n * p) / p;
}

const tabsDef = [
  { id:"codes", label:"0.코드(Code)" },
  { id:"steel", label:"철골(Steel)" },
  { id:"steelTotal", label:"철골_집계" },
  { id:"steelSub", label:"철골_부자재" },
  { id:"support", label:"동바리(support)" },
  { id:"supportTotal", label:"동바리_집계" }
];

let state = loadState();
recalcAll();
saveState();

const $tabs = document.getElementById("tabs");
const $view = document.getElementById("view");

function renderTabs(activeId){
  $tabs.innerHTML = "";
  for(const t of tabsDef){
    const b = document.createElement("button");
    b.className = "tab" + (t.id===activeId ? " active" : "");
    b.textContent = t.label;
    b.onclick = () => go(t.id);
    $tabs.appendChild(b);
  }
}

function go(id){
  recalcAll();
  saveState();
  renderTabs(id);

  if(id==="codes") renderCodes();
  else if(id==="steel") renderCalcSheet("철골(Steel)", state.steel, {totalMode:"steel"});
  else if(id==="steelSub") renderCalcSheet("철골_부자재(Processing and assembly)", state.steelSub, {totalMode:"steel"});
  else if(id==="support") renderCalcSheet("동바리(support)", state.support, {totalMode:"support"});
  else if(id==="steelTotal") renderSteelTotal();
  else if(id==="supportTotal") renderSupportTotal();
}

function panel(title, desc){
  const wrap = document.createElement("div");
  wrap.className = "panel";
  const h = document.createElement("div");
  h.className = "panel-header";
  const left = document.createElement("div");
  left.innerHTML = `<div class="panel-title">${title}</div><div class="panel-desc">${desc || ""}</div>`;
  h.appendChild(left);
  wrap.appendChild(h);
  return {wrap, header:h};
}

function renderCodes(){
  $view.innerHTML = "";
  const {wrap, header} = panel("0.코드(Code)", "코드 마스터(엑셀 VLOOKUP 원천). 수정/추가 가능");

  const right = document.createElement("div");
  right.style.display="flex";
  right.style.gap="8px";
  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가";
  addBtn.onclick = ()=>{
    state.codes.push({code:"",name:"",spec:"",unit:"",surcharge:"",conv_unit:"",conv_factor:"",note:""});
    saveState(); go("codes");
  };
  right.appendChild(addBtn);

  header.appendChild(right);

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:160px;">코드</th>
          <th style="min-width:220px;">품명</th>
          <th style="min-width:220px;">규격</th>
          <th style="min-width:90px;">단위</th>
          <th style="min-width:110px;">할증(%)</th>
          <th style="min-width:120px;">환산단위</th>
          <th style="min-width:140px;">환산계수</th>
          <th style="min-width:260px;">비고</th>
          <th style="min-width:120px;">작업</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;
  const tbody = tableWrap.querySelector("tbody");

  state.codes.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${inputCell(r.code, v=>{ r.code=v; }, "예: A0SM355150")}</td>
      <td>${inputCell(r.name, v=>{ r.name=v; })}</td>
      <td>${inputCell(r.spec, v=>{ r.spec=v; })}</td>
      <td>${inputCell(r.unit, v=>{ r.unit=v; })}</td>
      <td>${inputCell(r.surcharge, v=>{ r.surcharge=v; }, "예: 7")}</td>
      <td>${inputCell(r.conv_unit, v=>{ r.conv_unit=v; })}</td>
      <td>${inputCell(r.conv_factor, v=>{ r.conv_factor=v; }, "예: 0.0315")}</td>
      <td>${textAreaCell(r.note, v=>{ r.note=v; })}</td>
      <td></td>
    `;
    const tdAct = tr.lastElementChild;
    const act = document.createElement("div");
    act.className="row-actions";
    const del = document.createElement("button");
    del.className="smallbtn";
    del.textContent="삭제";
    del.onclick=()=>{
      state.codes.splice(idx,1);
      saveState(); go("codes");
    };
    act.appendChild(del);
    tdAct.appendChild(act);
    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);
  wireCells();
}

function renderCalcSheet(title, rows, opts){
  $view.innerHTML = "";
  const desc = (opts.totalMode==="steel")
    ? '※ 물량(E)에 "=10*100"처럼 "="로 계산식 입력 가능(간단 사칙연산). 환산/할증은 코드마스터 기반 자동.'
    : '※ 동바리도 동일 구조로 입력하되, 집계는 엑셀과 동일하게 "물량(Value)" 합계 기준으로 표시됩니다.';

  const {wrap, header} = panel(title, desc);

  const right = document.createElement("div");
  right.style.display="flex";
  right.style.gap="8px";

  const addBtn = document.createElement("button");
  addBtn.className="smallbtn";
  addBtn.textContent="행 추가";
  addBtn.onclick=()=>{
    rows.push(makeEmptyCalcRow());
    saveState(); go(tabsDef.find(x=>x.label===title)?.id || "steel");
  };

  const add10Btn = document.createElement("button");
  add10Btn.className="smallbtn";
  add10Btn.textContent="+10행";
  add10Btn.onclick=()=>{
    for(let i=0;i<10;i++) rows.push(makeEmptyCalcRow());
    saveState(); go(tabsDef.find(x=>x.label===title)?.id || "steel");
  };

  right.appendChild(addBtn);
  right.appendChild(add10Btn);
  header.appendChild(right);

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:70px;">No</th>
          <th style="min-width:170px;">코드</th>
          <th style="min-width:220px;">품명(자동)</th>
          <th style="min-width:220px;">규격(자동)</th>
          <th style="min-width:90px;">단위(자동)</th>
          <th style="min-width:150px;">물량(Value)</th>
          <th style="min-width:220px;">산출식(자동표시)</th>
          <th style="min-width:220px;">비고</th>
          <th style="min-width:120px;">할증(배수)</th>
          <th style="min-width:120px;">환산단위</th>
          <th style="min-width:140px;">환산계수</th>
          <th style="min-width:140px;">환산수량</th>
          <th style="min-width:160px;">할증후수량</th>
          <th style="min-width:120px;">작업</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;

  const tbody = tableWrap.querySelector("tbody");

  rows.forEach((r, idx)=>{
    recalcRow(r);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${inputCell(r.code, v=>{ r.code=v; }, "코드 입력")}</td>
      <td>${readonlyCell(r.name)}</td>
      <td>${readonlyCell(r.spec)}</td>
      <td>${readonlyCell(r.unit)}</td>
      <td>${inputCell(r.value, v=>{ r.value=v; }, '예: 100 또는 "=10*100"')}</td>
      <td>${readonlyCell(r.formula)}</td>
      <td>${textAreaCell(r.note, v=>{ r.note=v; })}</td>
      <td>${readonlyCell(r.surchargeMul === "" ? "" : String(r.surchargeMul))}</td>
      <td>${readonlyCell(r.convUnit)}</td>
      <td>${inputCell(r.convFactor, v=>{ r.convFactor=v; }, "비워도 됨")}</td>
      <td>${readonlyCell(String(roundUp3(r.convQty)))}</td>
      <td>${readonlyCell(String(roundUp3(r.finalQty)))}</td>
      <td></td>
    `;

    const tdAct = tr.lastElementChild;
    const act = document.createElement("div");
    act.className="row-actions";

    const del = document.createElement("button");
    del.className="smallbtn";
    del.textContent="삭제";
    del.onclick=()=>{
      rows.splice(idx,1);
      saveState();
      // 현재 탭 유지
      if(title.includes("부자재")) go("steelSub");
      else if(title.includes("동바리")) go("support");
      else go("steel");
    };

    const dup = document.createElement("button");
    dup.className="smallbtn";
    dup.textContent="복제";
    dup.onclick=()=>{
      rows.splice(idx+1,0, JSON.parse(JSON.stringify(r)));
      saveState();
      if(title.includes("부자재")) go("steelSub");
      else if(title.includes("동바리")) go("support");
      else go("steel");
    };

    act.appendChild(dup);
    act.appendChild(del);
    tdAct.appendChild(act);

    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);

  // 요약 배지
  const sumBox = document.createElement("div");
  sumBox.style.marginTop="10px";
  const sumVal = (opts.totalMode==="steel")
    ? rows.reduce((a,b)=>a+roundUp3(b.finalQty),0)
    : rows.reduce((a,b)=>a+evalValueExpr(b.value),0);

  sumBox.innerHTML = `<span class="badge">합계: ${roundUp3(sumVal)}</span>`;
  wrap.appendChild(sumBox);

  $view.appendChild(wrap);
  wireCells();
}

function renderSteelTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("철골_집계(Steel_Total quantity)", "코드별 합계(엑셀과 동일하게 철골/부자재의 '할증후수량' 합산)");

  recalcAll();
  const all = [...state.steel, ...state.steelSub];
  const grouped = groupSum(all, r => roundUp3(r.finalQty));

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:170px;">코드</th>
          <th style="min-width:220px;">품명</th>
          <th style="min-width:220px;">규격</th>
          <th style="min-width:90px;">단위</th>
          <th style="min-width:160px;">할증후수량</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;
  const tbody = tableWrap.querySelector("tbody");

  grouped.forEach(g=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(g.code)}</td>
      <td>${escapeHtml(g.name)}</td>
      <td>${escapeHtml(g.spec)}</td>
      <td>${escapeHtml(g.unit)}</td>
      <td>${roundUp3(g.sum)}</td>
    `;
    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);
}

function renderSupportTotal(){
  $view.innerHTML = "";
  const {wrap} = panel("동바리_집계(Support_Total quantity)", "코드별 합계(엑셀과 동일: 동바리 탭 '물량(Value)' E열 합계)");

  recalcAll();
  const grouped = groupSum(state.support, r => evalValueExpr(r.value));

  const tableWrap = document.createElement("div");
  tableWrap.className="table-wrap";
  tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="min-width:170px;">코드</th>
          <th style="min-width:220px;">품명</th>
          <th style="min-width:220px;">규격</th>
          <th style="min-width:90px;">단위</th>
          <th style="min-width:160px;">물량(Value)</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;
  const tbody = tableWrap.querySelector("tbody");

  grouped.forEach(g=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(g.code)}</td>
      <td>${escapeHtml(g.name)}</td>
      <td>${escapeHtml(g.spec)}</td>
      <td>${escapeHtml(g.unit)}</td>
      <td>${roundUp3(g.sum)}</td>
    `;
    tbody.appendChild(tr);
  });

  wrap.appendChild(tableWrap);
  $view.appendChild(wrap);
}

/* ---------------- UI helpers ---------------- */

const cellRegistry = [];
function inputCell(value, onChange, placeholder=""){
  const id = crypto.randomUUID();
  cellRegistry.push({id, type:"input", onChange});
  return `<input class="cell" data-cell="${id}" value="${escapeAttr(value ?? "")}" placeholder="${escapeAttr(placeholder)}" />`;
}
function textAreaCell(value, onChange){
  const id = crypto.randomUUID();
  cellRegistry.push({id, type:"textarea", onChange});
  return `<textarea class="cell" data-cell="${id}">${escapeHtml(value ?? "")}</textarea>`;
}
function readonlyCell(value){
  return `<input class="cell readonly" value="${escapeAttr(value ?? "")}" readonly />`;
}

function wireCells(){
  // 이벤트 바인딩
  document.querySelectorAll("[data-cell]").forEach(el=>{
    const id = el.getAttribute("data-cell");
    const meta = cellRegistry.find(x=>x.id===id);
    if(!meta) return;

    const handler = ()=>{
      meta.onChange(el.value);
      recalcAll();
      saveState();
      // 현재 뷰 재렌더(입력 반영)
      const active = document.querySelector(".tab.active")?.textContent || "";
      if(active.includes("0.코드")) go("codes");
      else if(active.includes("철골_집계")) go("steelTotal");
      else if(active.includes("동바리_집계")) go("supportTotal");
      else if(active.includes("부자재")) go("steelSub");
      else if(active.includes("동바리")) go("support");
      else go("steel");
    };

    // input: typing 즉시 반영은 부담될 수 있으니 change/blur 중심
    el.addEventListener("change", handler);
    el.addEventListener("blur", handler);
  });
  cellRegistry.length = 0;
}

function escapeHtml(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;");
}
function escapeAttr(s){
  return escapeHtml(s).replaceAll("\n","&#10;");
}

/* ---------------- Export / Import / Reset ---------------- */

document.getElementById("btnExport").addEventListener("click", ()=>{
  recalcAll();
  saveState();
  const blob = new Blob([JSON.stringify(state, null, 2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `FIN_WEB_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
});

document.getElementById("fileImport").addEventListener("change", async (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;
  const txt = await f.text();
  try{
    const parsed = JSON.parse(txt);
    state = parsed;
    recalcAll();
    saveState();
    go("codes");
  }catch{
    alert("JSON 파싱 실패: 파일 내용을 확인해 주세요.");
  }finally{
    e.target.value = "";
  }
});

document.getElementById("btnReset").addEventListener("click", ()=>{
  if(!confirm("정말 초기화할까요? (로컬 저장 데이터가 삭제됩니다)")) return;
  localStorage.removeItem(STORAGE_KEY);
  state = makeState();
  recalcAll();
  saveState();
  go("codes");
});

/* ---------------- start ---------------- */
go("steel");
