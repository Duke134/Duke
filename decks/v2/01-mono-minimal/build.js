// Mono Minimal — 양남규 경영지원 프로필 덱 v2 (14장)
// Run: NODE_PATH=<node_modules> node build.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fi = require("react-icons/fi");
const { applyTheme } = require("/root/.claude/skills/synced/7d6f9636-edbe-41a2-aab0-0f84fceac3dd_307ca375-45f7-4235-8c53-4eacd5b231ab/pptx/scripts/apply_theme.js");

const OUT = path.join(__dirname, "deck.pptx");

// ---------- Theme ----------
const THEME = {
  name: "Mono Minimal",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "212121", lt1: "FFFFFF", dk2: "36454F", lt2: "F2F2F2",
    accent1: "D64545", accent2: "36454F", accent3: "212121",
    accent4: "F2F2F2", accent5: "36454F", accent6: "212121",
    hlink: "D64545", folHlink: "36454F",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "양남규 경영지원 프로필 — Mono Minimal";
pres.author = "양남규";
const C = pres.SchemeColor;

// ---------- Icons ----------
async function icon(Comp, hex, px = 512) {
  let svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: px }));
  svg = svg.replace(/currentColor/g, "#" + hex);
  if (!/xmlns=/.test(svg)) svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: HEX.lt1 },
  objects: [],
});
const footer = { text: { text: "양남규 · 경영지원 프로필 2026", options: { x: 0.6, y: 5.22, w: 3.5, h: 0.25, fontSize: 9, color: HEX.dk2, margin: 0, isTextBox: true } } };
const slideNo = { x: 8.9, y: 5.22, w: 0.5, h: 0.25, fontSize: 9, color: HEX.dk2, align: "right", margin: 0 };
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  objects: [
    footer,
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 8.8, h: 0.75, fontSize: 36, bold: true, color: HEX.dk1, margin: 0, valign: "top", align: "left" } } },
  ],
  slideNumber: slideNo,
});
// Case slides carry long titles that wrap to two lines
pres.defineSlideMaster({
  title: "CASE",
  background: { color: HEX.lt1 },
  objects: [
    footer,
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 8.8, h: 1.25, fontSize: 36, bold: true, color: HEX.dk1, margin: 0, valign: "top", align: "left", lineSpacingMultiple: 0.95 } } },
  ],
  slideNumber: slideNo,
});

// ---------- Helpers ----------
const T = (s, o) => Object.assign({ isTextBox: true, margin: 0, color: C.text1 }, o);

function contentSlide(num, label, title, section, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  // Motif: giant light-gray number behind content
  s.addText(num, T(null, { x: 0.3, y: -0.75, w: 5.5, h: 3.3, fontSize: 250, bold: true, color: C.background2, valign: "top", align: "left", objectName: "Giant number " + num }));
  s.addText(label, T(null, { x: 0.6, y: 0.36, w: 5, h: 0.25, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, objectName: "Section label" }));
  const tt = Array.isArray(title) ? title.map((ln, i) => ({ text: ln, options: { breakLine: i < title.length - 1 } })) : title;
  s.addText(tt, { placeholder: "title" });
  return s;
}

function pill(s, text, x, y, w, h, opts = {}) {
  s.addText(text, Object.assign({
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: Math.min(h / 2, 0.2),
    fill: { color: C.background2 }, x, y, w, h,
    fontSize: 11, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true,
    objectName: "Tag " + text,
  }, opts));
}

function heading(s, text, x, y, w) {
  s.addText(text, T(null, { x, y, w, h: 0.25, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2, objectName: "Heading " + text }));
}

function iconCircle(s, data, x, y, d, name) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, objectName: "Icon circle " + name });
  const pad = d * 0.26;
  s.addImage({ data, x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: "Icon " + name });
}

// Horizontal step flow: pills joined by accent arrows
function flow(s, steps, x, y, w, h, fontSize = 12) {
  const gap = 0.32;
  const pw = (w - gap * (steps.length - 1)) / steps.length;
  steps.forEach((st, i) => {
    const px = x + i * (pw + gap);
    pill(s, st, px, y, pw, h, { fontSize, bold: true, objectName: "Step " + st });
    if (i < steps.length - 1) {
      s.addText("→", T(null, { x: px + pw, y, w: gap, h, fontSize: 14, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: "Arrow " + i }));
    }
  });
}

// Vertical step flow
function vflow(s, steps, x, y, w, h, fontSize = 12) {
  const gap = 0.18;
  steps.forEach((st, i) => {
    const py = y + i * (h + gap);
    pill(s, st, x, py, w, h, { fontSize, bold: true, objectName: "Step " + st });
    if (i < steps.length - 1) {
      s.addText("↓", T(null, { x, y: py + h, w, h: gap, fontSize: 12, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: "Arrow down " + i }));
    }
  });
}

(async () => {
  const W = HEX.lt1, D = HEX.dk1;
  const ic = {
    fin: await icon(Fi.FiDollarSign, W), hr: await icon(Fi.FiUsers, W), ir: await icon(Fi.FiTrendingUp, W),
    gov: await icon(Fi.FiFileText, W), glb: await icon(Fi.FiGlobe, W), ga: await icon(Fi.FiArchive, W),
    brief: await icon(Fi.FiBriefcase, W), layers: await icon(Fi.FiLayers, W), check: await icon(Fi.FiCheckCircle, W),
    mail: await icon(Fi.FiMail, D),
  };

  // ===== 1. 표지 =====
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    s.addText("경영지원 매니저", T(null, { x: 0.8, y: 1.2, w: 5, h: 0.4, fontSize: 18, color: C.text2, objectName: "Kicker" }));
    s.addText([{ text: "양남규" }, { text: ".", options: { color: C.accent1 } }], T(null, { x: 0.8, y: 1.6, w: 7, h: 1.4, fontSize: 88, bold: true, objectName: "Name" }));
    s.addText("Duke Yang", T(null, { x: 0.8, y: 3.05, w: 8, h: 0.45, fontSize: 22, bold: true, color: C.text2, objectName: "English name" }));
    s.addText("IT 스타트업에서 재무·인사·IR·정부사업·해외업무·총무를 혼자 맡아 운영하며, 그 일을 AI와 지식관리 도구로 시스템화하는 경영지원 매니저입니다.",
      T(null, { x: 0.8, y: 3.6, w: 8.4, h: 0.8, fontSize: 15, lineSpacingMultiple: 1.3, valign: "top", objectName: "Tagline" }));
    s.addText("기준일 2026-10-27", T(null, { x: 0.8, y: 4.75, w: 5, h: 0.3, fontSize: 11, color: C.text2, objectName: "Date" }));
  }

  // ===== 2. 일하는 원칙 =====
  pres.addSection({ title: "About" });
  {
    const s = contentSlide("02", "PRINCIPLES", "일하는 원칙", "About");
    const items = [
      ["시스템", "사람의 기억이 아닌 시스템에 남긴다."],
      ["원문 확인", "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보낸다."],
      ["스킬 · SOP", "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다."],
    ];
    items.forEach(([kw, sent], i) => {
      const x = 0.6 + i * 2.95, w = 2.65;
      s.addText(String(i + 1).padStart(2, "0"), T(null, { x, y: 1.9, w, h: 0.35, fontSize: 14, bold: true, color: C.accent1, objectName: "Principle no " + (i + 1) }));
      s.addText(kw, T(null, { x, y: 2.3, w, h: 0.7, fontSize: 32, bold: true, objectName: "Keyword " + kw }));
      s.addText(sent, T(null, { x, y: 3.2, w, h: 1.4, fontSize: 15, lineSpacingMultiple: 1.3, valign: "top", color: C.text2, objectName: "Principle " + (i + 1) }));
    });
  }

  // ===== 3. 소개 =====
  {
    const s = contentSlide("03", "ABOUT", "소개", "About");
    const paras = [
      [ic.brief, "IT 스타트업 경영지원팀 매니저로 일하며 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다."],
      [ic.glb, "국제통상학을 전공했고, 미국 현지 유통기업에서 근무한 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다."],
      [ic.layers, "여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해 옵시디언 기반 업무 운영체계와 AI 자동화 워크플로를 직접 설계해 쓰고 있습니다."],
      [ic.check, "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다."],
    ];
    const ys = [1.6, 2.45, 3.3, 4.15];
    paras.forEach(([img, text], i) => {
      const y = ys[i];
      iconCircle(s, img, 0.6, y, 0.46, "About " + i);
      s.addText(text, T(null, { x: 1.3, y: y - 0.04, w: 7.9, h: i === 2 ? 0.95 : 0.7, fontSize: 14, lineSpacingMultiple: 1.25, valign: "top", objectName: "About para " + (i + 1) }));
    });
  }

  // ===== 4. 담당 영역 =====
  pres.addSection({ title: "Areas" });
  {
    const s = contentSlide("04", "AREAS", "담당 영역", "Areas");
    s.addText("여섯 영역을 한 사람이 맡아 운영합니다", T(null, { x: 0.6, y: 1.32, w: 8, h: 0.3, fontSize: 14, color: C.text2, objectName: "Subtitle" }));
    const cards = [
      ["재무·회계", "자금 운용, 세무 전 주기, 결산·외부감사", ic.fin],
      ["인사·노무", "급여·4대보험, 채용, 해외 원격 인력 계약", ic.hr],
      ["IR·투자", "주총·이사회, 등기, 스톡옵션, 투자 계약", ic.ir],
      ["정부사업", "공고 분석부터 정산까지 전 주기", ic.gov],
      ["해외업무", "해외 자회사 증빙, 해외 행사·출장", ic.glb],
      ["총무", "법인 행정, 기업 인증, 사무실·자산, 대외행사", ic.ga],
    ];
    cards.forEach(([t, d, img], i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * 2.95, y = 1.85 + row * 1.6, w = 2.65;
      iconCircle(s, img, x, y, 0.5, t);
      s.addText(t, T(null, { x: x + 0.65, y: y + 0.05, w: w - 0.65, h: 0.4, fontSize: 17, bold: true, objectName: "Area " + t }));
      s.addText(d, T(null, { x, y: y + 0.65, w, h: 0.7, fontSize: 13, color: C.text2, lineSpacingMultiple: 1.2, valign: "top", objectName: "Area desc " + t }));
    });
  }

  // ===== 5. 담당 영역 상세 1 =====
  {
    const s = contentSlide("05", "AREAS · DETAIL 1", "담당 영역 상세 1", "Areas");
    s.addText("재무·회계 / 인사·노무 / IR·투자", T(null, { x: 0.6, y: 1.32, w: 8, h: 0.3, fontSize: 14, color: C.text2, objectName: "Subtitle" }));
    const rows = [
      ["재무·회계", ic.fin, ["자금 운용과 송금·결제를 관리하고, 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응합니다.", "결산과 외부감사 대응, 별도·연결 재무제표 관리를 맡고 외부 회계법인과의 커뮤니케이션을 담당합니다."]],
      ["인사·노무", ic.hr, ["급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다.", "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다."]],
      ["IR·투자", ic.ir, ["정기주주총회·이사회 운영, 정관 개정, 등기, 주주명부와 스톡옵션 관리를 맡고 있습니다.", "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다."]],
    ];
    rows.forEach(([t, img, paras], i) => {
      const y = 1.72 + i * 1.15;
      iconCircle(s, img, 0.6, y, 0.46, t);
      s.addText(t, T(null, { x: 1.2, y: y + 0.06, w: 1.5, h: 0.35, fontSize: 16, bold: true, objectName: "Detail " + t }));
      s.addText(paras.map((p, j) => ({ text: p, options: { breakLine: j < paras.length - 1 } })),
        T(null, { x: 2.7, y: y - 0.02, w: 6.7, h: 1.0, fontSize: 13, lineSpacingMultiple: 1.15, paraSpaceAfter: 3, valign: "top", objectName: "Detail text " + t }));
    });
  }

  // ===== 6. 담당 영역 상세 2 =====
  {
    const s = contentSlide("06", "AREAS · DETAIL 2", "담당 영역 상세 2", "Areas");
    s.addText("정부사업 / 해외업무 / 총무", T(null, { x: 0.6, y: 1.32, w: 8, h: 0.3, fontSize: 14, color: C.text2, objectName: "Subtitle" }));
    // 정부사업 — 5-step flow
    iconCircle(s, ic.gov, 0.6, 1.78, 0.46, "정부사업");
    s.addText("정부사업", T(null, { x: 1.2, y: 1.84, w: 1.5, h: 0.35, fontSize: 16, bold: true, objectName: "Detail 정부사업" }));
    flow(s, ["공고 분석", "사업계획서 작성", "신청", "협약", "정산"], 2.75, 1.78, 6.65, 0.42, 11);
    s.addText([
      { text: "공고 분석 → 사업계획서 작성 → 신청 → 협약 → 정산까지 전 주기를 실무로 수행합니다.", options: { breakLine: true } },
      { text: "입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다." },
    ], T(null, { x: 2.75, y: 2.32, w: 6.65, h: 0.8, fontSize: 13, lineSpacingMultiple: 1.2, paraSpaceAfter: 4, valign: "top", objectName: "Detail text 정부사업" }));
    // 해외업무 / 총무 — two columns
    const cols = [
      ["해외업무", ic.glb, ["해외 자회사의 인보이스와 용역계약 증빙을 관리하고, 해외 사업 행사 참가와 해외 출장 서류 실무를 지원합니다."]],
      ["총무", ic.ga, ["본점 이전 같은 법인 행정 절차와 기업 인증 관리, 사무실·자산 관리를 담당합니다.", "대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다."]],
    ];
    cols.forEach(([t, img, paras], i) => {
      const x = 0.6 + i * 4.5, y = 3.4, w = 4.2;
      iconCircle(s, img, x, y, 0.46, t);
      s.addText(t, T(null, { x: x + 0.6, y: y + 0.06, w: 2, h: 0.35, fontSize: 16, bold: true, objectName: "Detail " + t }));
      s.addText(paras.map((p, j) => ({ text: p, options: { breakLine: j < paras.length - 1 } })),
        T(null, { x, y: y + 0.62, w, h: 1.1, fontSize: 13, lineSpacingMultiple: 1.2, paraSpaceAfter: 4, valign: "top", objectName: "Detail text " + t }));
    });
  }

  // ===== 7-11. 사례 =====
  pres.addSection({ title: "Cases" });
  const cases = [
    { n: "07", label: "CASE 01", title: ["1인 경영지원을 \"인계 가능한", "시스템\"으로 바꾸기"],
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 커스텀 AI 스킬로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다." },
    { n: "08", label: "CASE 02", title: ["단일 통화 자금 시트를", "의사결정 도구로 재설계"],
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리하고 있어, 경영진의 자금 판단 질문에 바로 답하기 어려웠습니다.",
      did: "대시보드·통화별 입출금·정기 비용·미지급 항목으로 시트를 나누고, 다중 통화 환산과 \"부족액 식별 → 송금액 계산 → 송금 후 잔액\" 시뮬레이션을 자동화했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.",
      result: "경영진의 자금 의사결정 질문에 한 화면에서 바로 답할 수 있는 상태가 됐습니다.",
      steps: ["부족액 식별", "송금액 계산", "송금 후 잔액"] },
    { n: "09", label: "CASE 03", title: ["법률 1차 검토를", "법령 원문 기반으로 내재화"],
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했습니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 실무 상식만으로는 놓칠 수 있는 별도 조항을 미리 잡게 됐습니다." },
    { n: "10", label: "CASE 04", title: ["외부로 나가기 전 \"받은 산출물\"을", "역검증하는 습관"],
      bg: "회계법인·기관·플랫폼에서 받은 자료나 내부 문서가 확인 없이 외부로 나가면 세금·등기·평가에 바로 영향을 줍니다.",
      did: "회계법인 산출물, 기관·플랫폼 자료, 계약서와 내부 문서를 원장·법령 원문과 대조한 뒤 제출하는 절차를 분기 내내 운영했습니다.",
      result: "기간 오류, 견적 오기, 정관과 계약의 충돌 같은 문제를 외부 제출 전에 찾아냈습니다." },
    { n: "11", label: "CASE 05", title: "예산 대비 실적 관리 워크북 구축",
      bg: "월별 지출 추적이 여러 엑셀 파일에 흩어져 있어 예산을 통제하기 어려웠습니다.",
      did: "실제 계정원장을 기준으로 계정과목을 정리해 카테고리로 묶었습니다. 월별 예산 수립 → 실적 입력 → 차이 분석 흐름을 자동화했고, 이후 중앙 가정 시트와 연동되도록 확장했습니다.",
      result: "예산 대비 실적을 한 체계에서 관리하는 관리회계 기반을 만들었습니다.",
      steps: ["예산 수립", "실적 입력", "차이 분석"] },
  ];
  for (const c of cases) {
    const s = contentSlide(c.n, c.label, c.title, "Cases", "CASE");
    const LX = 0.6, LW = 4.9, RX = 5.9, RW = 3.5, TOP = 2.0;
    heading(s, "배경", LX, TOP, LW);
    s.addText(c.bg, T(null, { x: LX, y: TOP + 0.3, w: LW, h: 0.95, fontSize: 13, lineSpacingMultiple: 1.25, valign: "top", objectName: "Background" }));
    heading(s, "한 일", LX, TOP + 1.3, LW);
    s.addText(c.did, T(null, { x: LX, y: TOP + 1.6, w: LW, h: 1.5, fontSize: 13, lineSpacingMultiple: 1.25, valign: "top", objectName: "What I did" }));
    // 결과 — tinted box with accent label
    const boxH = c.steps ? 1.5 : 2.1;
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y: TOP, w: RW, h: boxH, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: "Result box" });
    heading(s, "결과", RX + 0.3, TOP + 0.25, RW - 0.6);
    s.addText(c.result, T(null, { x: RX + 0.3, y: TOP + 0.53, w: RW - 0.6, h: boxH - 0.65, fontSize: 13, bold: true, lineSpacingMultiple: 1.25, valign: "top", objectName: "Result" }));
    if (c.steps) vflow(s, c.steps, RX, TOP + boxH + 0.2, RW, 0.28, 12);
  }

  // ===== 12. 스킬·도구 =====
  pres.addSection({ title: "Skills" });
  {
    const s = contentSlide("12", "SKILLS & TOOLS", "스킬·도구", "Skills");
    const skills = [
      ["엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·자금 시뮬레이션 설계"],
      ["옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인"],
      ["AI 업무 자동화", "Claude 커스텀 스킬 설계·운영, 예약 작업 자동화, MCP 연동(법령 원문 조회, 옵시디언 연동)"],
      ["협업 도구", "Gmail, Google Drive, Slack"],
      ["ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)"],
      ["문서 표준화", "공문 서식, docx·pptx 폰트 표준, 인보이스·인사·급여 템플릿, SOP"],
      ["실무 지식", "결산·세무 신고 대응, 주총·이사회·등기, 정부사업 전 주기"],
      ["언어", "비즈니스 영어(미국 현지 근무 경험)"],
    ];
    const rows = skills.map(([k, v], i) => {
      const o = i % 2 === 0 ? { fill: { color: C.background2 } } : {};
      return [
        { text: k, options: Object.assign({ bold: true, color: C.text1 }, o) },
        { text: v, options: Object.assign({ color: C.text1 }, o) },
      ];
    });
    s.addTable(rows, { x: 0.6, y: 1.6, w: 8.8, colW: [2.0, 6.8], rowH: 0.4, fontSize: 13, fontFace: "Malgun Gothic",
      border: { type: "none" }, margin: [0.06, 0.15, 0.06, 0.15], valign: "middle", objectName: "Skills table" });
  }

  // ===== 13. 타임라인 =====
  pres.addSection({ title: "Timeline" });
  {
    const s = contentSlide("13", "TIMELINE", "타임라인", "Timeline");
    const items = [
      ["2022 – 현재", "IT 스타트업", "경영지원팀 매니저"],
      ["2019 – 2022", "미국 소재 유통기업", "재고관리 인턴 → 구매(Purchasing) 담당"],
      ["2018 – 2019", "미국 비즈니스 취업연수 과정 수료", ""],
      ["2017", "중국 교환학생", ""],
      ["2013 – 2019", "학부", "국제통상학 전공 (중어중문학 복수전공)"],
    ];
    items.forEach(([date, t, d], i) => {
      const y = 1.65 + i * 0.68;
      s.addText(date, T(null, { x: 0.6, y: y + 0.02, w: 1.5, h: 0.3, fontSize: 14, bold: true, color: C.text2, objectName: "Date " + i }));
      s.addShape(pres.shapes.OVAL, { x: 2.3, y: y + 0.08, w: 0.16, h: 0.16, fill: { color: i === 0 ? C.accent1 : C.text2 }, line: { color: i === 0 ? C.accent1 : C.text2, width: 0 }, objectName: "Dot " + i });
      s.addText(t, T(null, { x: 2.7, y, w: 6.7, h: 0.32, fontSize: 16, bold: true, objectName: "Event " + i }));
      if (d) s.addText(d, T(null, { x: 2.7, y: y + 0.33, w: 6.7, h: 0.3, fontSize: 12, color: C.text2, valign: "top", objectName: "Event desc " + i }));
    });
  }

  // ===== 14. 연락 =====
  pres.addSection({ title: "Contact" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Contact" });
    s.addText("CONTACT", T(null, { x: 0.8, y: 1.45, w: 5, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 3, objectName: "Contact label" }));
    s.addText([{ text: "양남규" }, { text: ".", options: { color: C.accent1 } }], T(null, { x: 0.8, y: 1.8, w: 7, h: 1.3, fontSize: 80, bold: true, objectName: "Name" }));
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: 3.35, w: 0.44, h: 0.44, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, objectName: "Mail circle" });
    s.addImage({ data: ic.mail, x: 0.91, y: 3.46, w: 0.22, h: 0.22, objectName: "Mail icon" });
    s.addText("namgyu1412@naver.com", T(null, { x: 1.4, y: 3.35, w: 6, h: 0.44, fontSize: 18, bold: true, valign: "middle", objectName: "Email" }));
    s.addText("© 2026 Duke Yang", T(null, { x: 0.8, y: 4.95, w: 4, h: 0.25, fontSize: 10, color: C.text2, objectName: "Copyright" }));
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch(e => { console.error(e); process.exit(1); });
