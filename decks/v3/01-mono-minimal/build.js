// Mono Minimal — 양남규 경영지원 프로필 덱 v3 (16장)
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

// ---------- Text-fit estimator (inches) ----------
// Rough advance widths in em: Hangul/CJK 1.0, middle dot 1.0 (renders wide in CJK fonts), arrows 1.0,
// Latin ~0.56, digits 0.56, space 0.3, other punctuation 0.4.
function textWidthIn(str, pt) {
  let em = 0;
  for (const ch of str) {
    const cp = ch.codePointAt(0);
    if (cp === 0x20) em += 0.3;
    else if (cp === 0xB7 || cp === 0x2192 || cp === 0x2193) em += 1.0;
    else if (cp >= 0x1100 && cp <= 0x11FF || cp >= 0x3000 && cp <= 0x9FFF || cp >= 0xAC00 && cp <= 0xD7AF || cp >= 0xFF00 && cp <= 0xFFEF) em += 1.0;
    else if (/[A-Z]/.test(ch)) em += 0.66;
    else if (/[a-z0-9]/.test(ch)) em += 0.56;
    else em += 0.4;
  }
  return em * pt / 72;
}
function estLines(str, wIn, pt) {
  // Korean wraps between characters, so the line count follows the total width closely
  // (calibrated against LibreOffice renders of this deck's paragraphs).
  return Math.max(1, Math.ceil(textWidthIn(str, pt) / wIn * 0.93));
}
// height of N paragraphs at pt with line spacing multiple ls and paraSpaceAfter psa (pt)
function parasHeight(paras, wIn, pt, ls = 1.15, psa = 4) {
  let n = 0;
  paras.forEach(p => { n += estLines(p, wIn, pt); });
  return n * (pt * ls * 1.08 / 72) + (paras.length - 1) * psa / 72 + 0.04;
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
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 8.8, h: 1.2, fontSize: 36, bold: true, color: HEX.dk1, margin: 0, valign: "top", align: "left", lineSpacingMultiple: 0.95 } } },
  ],
  slideNumber: slideNo,
});

// ---------- Helpers ----------
const T = (s, o) => Object.assign({ isTextBox: true, margin: 0, color: C.text1 }, o);
const BODY = { fontSize: 13, lineSpacingMultiple: 1.15, paraSpaceAfter: 4, valign: "top" };
const BODY_TIGHT = { fontSize: 13, lineSpacingMultiple: 1.1, paraSpaceAfter: 3, valign: "top" };
const paraRuns = (paras) => paras.map((p, j) => ({ text: p, options: { breakLine: j < paras.length - 1 } }));

function contentSlide(num, label, title, section, master = "CONTENT", subtitle) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  // Motif: giant light-gray number behind content
  s.addText(num, T(null, { x: 0.3, y: -0.75, w: 5.5, h: 3.3, fontSize: 250, bold: true, color: C.background2, valign: "top", align: "left", objectName: "Giant number " + num }));
  s.addText(label, T(null, { x: 0.6, y: 0.36, w: 5, h: 0.25, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, objectName: "Section label" }));
  let tt;
  if (Array.isArray(title)) tt = title.map((ln, i) => ({ text: ln, options: { breakLine: i < title.length - 1 } }));
  else if (subtitle) tt = [{ text: title }, { text: "   " + subtitle, options: { fontSize: 14, bold: false, color: C.text2 } }];
  else tt = title;
  s.addText(tt, { placeholder: "title" });
  return s;
}

function pill(s, text, x, y, w, h, opts = {}) {
  const name = Array.isArray(text) ? text.map(r => r.text).join("") : text;
  s.addText(text, Object.assign({
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: Math.min(h / 2, 0.2),
    fill: { color: C.background2 }, x, y, w, h,
    fontSize: 12, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true,
    objectName: "Tag " + name,
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
function flow(s, steps, x, y, w, h, fontSize = 12, gap = 0.3) {
  const pw = (w - gap * (steps.length - 1)) / steps.length;
  steps.forEach((st, i) => {
    const px = x + i * (pw + gap);
    pill(s, st, px, y, pw, h, { fontSize, bold: true, lineSpacingMultiple: 1.0 });
    if (i < steps.length - 1) {
      s.addText("→", T(null, { x: px + pw, y, w: gap, h, fontSize: 14, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: "Arrow " + i }));
    }
  });
}

// Vertical step flow
function vflow(s, steps, x, y, w, h, fontSize = 12, gap = 0.17) {
  steps.forEach((st, i) => {
    const py = y + i * (h + gap);
    pill(s, st, x, py, w, h, { fontSize, bold: true });
    if (i < steps.length - 1) {
      s.addText("↓", T(null, { x, y: py + h, w, h: gap, fontSize: 12, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: "Arrow down " + st }));
    }
  });
}

(async () => {
  const W = HEX.lt1, D = HEX.dk1;
  const ic = {
    fin: await icon(Fi.FiDollarSign, W), hr: await icon(Fi.FiUsers, W), ir: await icon(Fi.FiTrendingUp, W),
    gov: await icon(Fi.FiFileText, W), glb: await icon(Fi.FiGlobe, W), ga: await icon(Fi.FiArchive, W),
    brief: await icon(Fi.FiBriefcase, W), layers: await icon(Fi.FiLayers, W), check: await icon(Fi.FiCheckCircle, W),
    book: await icon(Fi.FiBookOpen, W),
    mail: await icon(Fi.FiMail, D), target: await icon(Fi.FiTarget, D), award: await icon(Fi.FiAward, D),
  };

  // ===== 1. 표지 =====
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    s.addText("경영지원 매니저", T(null, { x: 0.8, y: 1.1, w: 5, h: 0.4, fontSize: 18, color: C.text2, objectName: "Kicker" }));
    s.addText([{ text: "양남규" }, { text: ".", options: { color: C.accent1 } }], T(null, { x: 0.8, y: 1.5, w: 7, h: 1.4, fontSize: 88, bold: true, objectName: "Name" }));
    s.addText("Duke Yang", T(null, { x: 0.8, y: 2.95, w: 8, h: 0.45, fontSize: 22, bold: true, color: C.text2, objectName: "English name" }));
    s.addText("IT 스타트업에서 국내 법인과 해외 자회사의 재무·인사·IR·정부사업·해외업무·총무를 혼자 맡아 운영하며, 그 일을 AI와 지식관리 도구로 시스템화하는 경영지원 매니저입니다.",
      T(null, { x: 0.8, y: 3.5, w: 8.4, h: 1.0, fontSize: 15, lineSpacingMultiple: 1.3, valign: "top", objectName: "Tagline" }));
    s.addText("기준일 2026-11-01", T(null, { x: 0.8, y: 4.75, w: 5, h: 0.3, fontSize: 11, color: C.text2, objectName: "Date" }));
  }

  // ===== 2. 일하는 원칙 =====
  pres.addSection({ title: "About" });
  {
    const s = contentSlide("02", "PRINCIPLES", "일하는 원칙", "About");
    const items = [
      ["시스템", "사람의 기억이 아닌 시스템에 남긴다."],
      ["원문 확인", "받은 자료를 그대로 쓰지 않는다. 회계법인·기관·플랫폼 산출물도 원장과 법령 원문에 대조한 뒤 내보낸다."],
      ["스킬·SOP", "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다."],
    ];
    items.forEach(([kw, sent], i) => {
      const x = 0.6 + i * 2.95, w = 2.65;
      s.addText(String(i + 1).padStart(2, "0"), T(null, { x, y: 1.9, w, h: 0.35, fontSize: 14, bold: true, color: C.accent1, objectName: "Principle no " + (i + 1) }));
      s.addText(kw, T(null, { x, y: 2.3, w, h: 0.7, fontSize: 32, bold: true, objectName: "Keyword " + kw }));
      s.addText(sent, T(null, { x, y: 3.2, w, h: 1.5, fontSize: 15, lineSpacingMultiple: 1.3, valign: "top", color: C.text2, objectName: "Principle " + (i + 1) }));
    });
  }

  // ===== 3. 소개 =====
  {
    const s = contentSlide("03", "ABOUT", "소개", "About");
    const paras = [
      [ic.brief, "2022년 말부터 IT 스타트업 경영지원팀 매니저로 일하며, 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다."],
      [ic.glb, "국제통상학을 전공했고, 미국 현지 유통기업에서 2년 8개월 구매·재고관리를 맡은 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다."],
      [ic.layers, "여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해, 옵시디언 기반 업무 운영체계와 자작 AI 스킬 28개·예약 작업으로 짜인 자동화 워크플로를 직접 설계해 쓰고 있습니다."],
      [ic.check, "노무·세무·상법 판단은 한국 법령과 미국 연방규정 원문으로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다."],
    ];
    let y = 1.5;
    paras.forEach(([img, text], i) => {
      const h = parasHeight([text], 7.9, 14, 1.25, 0);
      iconCircle(s, img, 0.6, y, 0.46, "About " + i);
      s.addText(text, T(null, { x: 1.3, y: y - 0.04, w: 7.9, h, fontSize: 14, lineSpacingMultiple: 1.25, valign: "top", objectName: "About para " + (i + 1) }));
      y += Math.max(h, 0.5) + 0.22;
    });
  }

  // ===== 4. 담당 영역 =====
  pres.addSection({ title: "Areas" });
  {
    const s = contentSlide("04", "AREAS", "담당 영역", "Areas", "CONTENT", "여섯 영역을 한 사람이 맡아 운영합니다");
    const cards = [
      ["재무·회계", "자금 계획·집행, 세무 전 주기, 결산·연결감사", ic.fin],
      ["인사·노무", "급여·4대보험, 채용, 퇴직·휴직 처리, 해외 원격 인력 계약", ic.hr],
      ["IR·투자", "주총·이사회, 등기, 스톡옵션, 투자 계약·신주발행 공고, 분기 실적보고", ic.ir],
      ["정부사업", "공고 분석부터 정산까지 전 주기, 기업부설연구소·기업 인증", ic.gov],
      ["해외업무", "해외 자회사 송금 증빙, 해외 법인 청산, 해외 행사·출장", ic.glb],
      ["총무", "법인 행정, 사무실·자산·계정 관리, 대외행사", ic.ga],
    ];
    cards.forEach(([t, d, img], i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.6 + col * 2.95, y = 1.75 + row * 1.65, w = 2.65;
      iconCircle(s, img, x, y, 0.5, t);
      s.addText(t, T(null, { x: x + 0.65, y: y + 0.05, w: w - 0.65, h: 0.4, fontSize: 17, bold: true, objectName: "Area " + t }));
      s.addText(d, T(null, { x, y: y + 0.65, w, h: 0.85, fontSize: 13, color: C.text2, lineSpacingMultiple: 1.2, valign: "top", objectName: "Area desc " + t }));
    });
  }

  // ===== 5. 담당 영역 상세 1 =====
  {
    const s = contentSlide("05", "AREAS · DETAIL 1", "담당 영역 상세 1", "Areas", "CONTENT", "재무·회계 / 인사·노무 / IR·투자");
    const rows = [
      ["재무·회계", ic.fin, [
        "자금 운용과 송금·결제를 관리하고, 입금 일정에 맞춘 지급 우선순위를 경영진에 보고합니다. 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응하고, 과거 연도 수정신고와 과세관청 소명자료 구성까지 맡습니다.",
        "결산과 외부감사 대응, 별도·연결 재무제표를 2개년 연속 관리했고, 외부 회계법인 3곳과의 커뮤니케이션을 담당합니다."]],
      ["인사·노무", ic.hr, [
        "급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다. 퇴직급여 분할 지급 합의, IRP 과세이연, 무급휴직 중 4대보험 처리 같은 예외 케이스도 법령 기준으로 설계합니다.",
        "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다."]],
      ["IR·투자", ic.ir, [
        "정기주주총회·이사회를 2년 연속 단독 운영하고, 정관 개정·등기·주주명부·스톡옵션 관리를 맡고 있습니다. 전환사채·신주인수권부사채 발행과 유상증자 실무, 본점 이전 등기를 외부 법무법인 없이 처리했습니다.",
        "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다."]],
    ];
    const TX = 2.3, TW = 7.1;
    let y = 1.35;
    rows.forEach(([t, img, paras], i) => {
      const h = parasHeight(paras, TW, 13, 1.1, 3);
      iconCircle(s, img, 0.6, y, 0.42, t);
      s.addText(t, T(null, { x: 1.12, y: y + 0.04, w: 1.15, h: 0.35, fontSize: 14, bold: true, objectName: "Detail " + t }));
      s.addText(paraRuns(paras), T(null, Object.assign({ x: TX, y: y - 0.03, w: TW, h, objectName: "Detail text " + t }, BODY_TIGHT)));
      y += h + 0.12;
    });
  }

  // ===== 6. 담당 영역 상세 2 =====
  {
    const s = contentSlide("06", "AREAS · DETAIL 2", "담당 영역 상세 2", "Areas", "CONTENT", "정부사업 / 해외업무 / 총무");
    const TX = 2.3, TW = 7.1;
    // 정부사업 — label + 3 paragraphs, then the 7-step flow
    const govParas = [
      "공고 분석 → 사업계획서 → 신청·입찰 → 협약·계약 → 운영 → 완료보고·검수 → 정산까지 전 주기를 실무로 수행합니다.",
      "2023년부터 11건 이상 수주, 5개 이상 주관기관을 대응했고, 입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다.",
      "기업부설연구소 설립과 벤처기업 인증 등 기업 인증을 관리합니다.",
    ];
    let y = 1.35;
    const gh = parasHeight(govParas, TW, 13, 1.1, 3);
    iconCircle(s, ic.gov, 0.6, y, 0.42, "정부사업");
    s.addText("정부사업", T(null, { x: 1.12, y: y + 0.04, w: 1.15, h: 0.35, fontSize: 14, bold: true, objectName: "Detail 정부사업" }));
    s.addText(paraRuns(govParas), T(null, Object.assign({ x: TX, y: y - 0.03, w: TW, h: gh, objectName: "Detail text 정부사업" }, BODY_TIGHT)));
    y += gh + 0.12;
    flow(s, ["공고 분석", "사업계획서", "신청·입찰", "협약·계약", "운영", [{ text: "완료보고·", options: { breakLine: true } }, { text: "검수" }], "정산"], 0.6, y, 8.8, 0.5, 12, 0.22);
    y += 0.5 + 0.22;
    // 해외업무 / 총무 — two columns
    const cols = [
      ["해외업무", ic.glb, 0.6, 4.95, [
        "해외 자회사 송금에 필요한 계약·인보이스 증빙을 심사 통과형 패키지로 관리하고, 해외 법인 청산 절차(현지 회계·법무 조율, 서류 영문화)를 지원합니다.",
        "해외 사업 행사 참가와 해외 출장 서류(ABTC) 실무를 맡습니다."]],
      ["총무", ic.ga, 5.85, 3.55, [
        "본점 이전 같은 법인 행정 절차, 사무실 입주·이전·임대차, 자산·계정 관리를 담당합니다.",
        "대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다."]],
    ];
    cols.forEach(([t, img, x, w, paras]) => {
      const h = parasHeight(paras, w, 13, 1.1, 3);
      iconCircle(s, img, x, y, 0.42, t);
      s.addText(t, T(null, { x: x + 0.52, y: y + 0.04, w: 2, h: 0.35, fontSize: 14, bold: true, objectName: "Detail " + t }));
      s.addText(paraRuns(paras), T(null, Object.assign({ x, y: y + 0.52, w, h, objectName: "Detail text " + t }, BODY_TIGHT)));
    });
  }

  // ===== 7-12. 사례 =====
  pres.addSection({ title: "Cases" });
  const cases = [
    { n: "07", label: "CASE 01", title: ["1인 경영지원을 \"인계 가능한", "시스템\"으로 바꾸기"],
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 자작 AI 스킬 28개로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다. 주간 할 일·회신 대기·중장기 과제가 한 볼트 안에서 연결됩니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다." },
    { n: "08", label: "CASE 02", title: ["자금 판단과 예산 통제를", "한 체계로 (관리회계)"],
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리했고, 월별 지출은 여러 엑셀 파일에 흩어져 있어 경영진의 자금 판단 질문에도, 예산 통제에도 바로 답하기 어려웠습니다.",
      did: "자금 시트를 대시보드·통화별 입출금·정기 비용·미지급 항목으로 나누고, 다중 통화 환산과 \"부족액 식별 → 송금액 계산 → 송금 후 잔액\" 시뮬레이션을 자동화했습니다. 예산은 실제 계정원장 기준 23개 계정과목을 6개 카테고리로 묶어 \"월별 예산 수립 → 실적 입력 → 차이 분석\" 워크북을 만들고, 중앙 가정 시트와 연동했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.",
      result: "경영진의 자금 의사결정 질문과 예산 대비 실적을 한 체계에서 바로 답할 수 있게 됐습니다.",
      flows: [["부족액 식별", "송금액 계산", "송금 후 잔액"], ["월별 예산 수립", "실적 입력", "차이 분석"]] },
    { n: "09", label: "CASE 03", title: ["법률 1차 검토를 한국 법령·", "미국 연방규정 원문 기반으로 내재화"],
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다. 해외 자회사의 미국 세무 요건은 확인할 수단 자체가 없었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구, 미국 연방규정(eCFR) 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례(행정해석)를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했고, 해외 자회사의 미국 세무 신고 요건도 연방규정 원문으로 확인합니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 투자 계약 일정을 상법 조문 기준으로 짜거나 현장조사 사전통지 의무처럼 실무 상식과 다른 조항을 미리 잡게 됐습니다." },
    { n: "10", label: "CASE 04", title: ["기업부설연구소, 요건 분석부터", "설립신고까지 열흘"],
      bg: "외부 일정 때문에 짧은 기간 안에 기업부설연구소 인정이 필요했고, 현장조사 가능성이 높은 상황이었습니다.",
      did: "인정 요건(독립 공간·전담 인력·연구개발 활동)을 법령과 인정기관 기준으로 정리하고, 공간 계약 변경·현판·도면·기자재 명세·인력 자격 증빙·연구개발활동개요서를 준비해 지시 열흘 만에 온라인 설립신고를 마쳤습니다. 현장조사 사전통지 규정을 조문으로 확인해 연구원 상주 범위를 최소로 잡았고, 재택 연구 시행문과 인사발령 공문 등 내부 규정 문서를 함께 정비했습니다.",
      result: "현장조사 대비 체크리스트와 설립 SOP가 남아, 이후 변경신고·연차 실적보고를 같은 절차로 운영할 수 있게 됐습니다." },
    { n: "11", label: "CASE 05", title: ["해외 자회사 송금 증빙을", "심사 통과형 패키지로"],
      bg: "국내 법인에서 해외 자회사로 자금을 보낼 때마다 송금 플랫폼 심사용 계약·인보이스가 필요했는데, 건마다 서류 형식과 지급 조건이 달라 심사 기준을 매번 다시 맞춰야 했습니다.",
      did: "마스터 계약과 SOW·인보이스의 관계를 정리해 표준 양식을 만들고, 서비스 완료 후 정산하는 SOW 구조를 선례 기준으로 확정했습니다. 이전 송금 5건을 전수 대조해 날짜·지급조건·수취인 정보 불일치를 고쳤고, 절차를 SOP로 남겼습니다.",
      result: "신규 건은 성과 보고서에서 항목을 뽑아 SOW·인보이스를 같은 날 만들고 심사를 한 번에 통과하는 흐름이 됐습니다." },
    { n: "12", label: "CASE 06", title: ["자금이 빠듯한 시기의 세무·노무", "리스크를 법령 기준으로 정리"],
      bg: "스타트업은 자금이 빠듯한 시기에 세무·노무 이슈가 한꺼번에 겹치고, 한 건의 처리 지연이 금융기관 거래까지 번집니다.",
      did: "과거 3개 연도 매출 누락 건의 부가세·법인세 수정신고를 회계법인과 2주 안에 마치고, 과세관청 소명자료를 거래 증빙으로 구성했습니다. 납부 일정은 세목별 분납으로 설계해 월 부담을 평준화했습니다. 퇴직급여는 분할 지급 합의서(법정기한·IRP 과세이연·지연 시 구제 조항)로 설계했고, 임금 관련 진정 사건 2건은 외부 자문 없이 합의·취하로 종결했습니다.",
      result: "과세관청·금융기관·노동 당국을 직접 상대해 처리한 절차가 SOP와 합의서 양식으로 남았습니다." },
  ];
  for (const c of cases) {
    const s = contentSlide(c.n, c.label, c.title, "Cases", "CASE");
    const LX = 0.6, LW = 5.4, RX = 6.35, RW = 3.05, TOP = 1.85;
    const bgH = parasHeight([c.bg], LW, 13, 1.15, 0);
    heading(s, "배경", LX, TOP, LW);
    s.addText(c.bg, T(null, Object.assign({ x: LX, y: TOP + 0.28, w: LW, h: bgH, objectName: "Background" }, BODY)));
    const y2 = TOP + 0.28 + bgH + 0.18;
    const didH = parasHeight([c.did], LW, 13, 1.15, 0);
    heading(s, "한 일", LX, y2, LW);
    s.addText(c.did, T(null, Object.assign({ x: LX, y: y2 + 0.28, w: LW, h: didH, objectName: "What I did" }, BODY)));
    // 결과 — tinted box with accent label
    const resH = parasHeight([c.result], RW - 0.5, 13, 1.15, 0);
    const boxH = 0.25 + 0.28 + resH + 0.42; // one line of slack under the result text
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y: TOP, w: RW, h: boxH, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: "Result box" });
    heading(s, "결과", RX + 0.25, TOP + 0.25, RW - 0.5);
    s.addText(c.result, T(null, Object.assign({ x: RX + 0.25, y: TOP + 0.53, w: RW - 0.5, h: resH, bold: true, objectName: "Result" }, BODY)));
    if (c.flows) {
      const fw = (RW - 0.15) / 2;
      c.flows.forEach((steps, i) => vflow(s, steps, RX + i * (fw + 0.15), TOP + boxH + 0.16, fw, 0.28, 12, 0.15));
    }
  }

  // ===== 13. AI 운영 체계 =====
  pres.addSection({ title: "AI" });
  {
    const s = contentSlide("13", "AI OPERATIONS", "AI 운영 체계", "AI");
    const card = (kw, body, x, y, w, h, name) => {
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: "Card " + name });
      s.addText(kw, T(null, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.4, fontSize: 22, bold: true, objectName: "Card keyword " + name }));
      s.addText(body, T(null, Object.assign({ x: x + 0.25, y: y + 0.66, w: w - 0.5, h: h - 0.8, objectName: "Card text " + name }, BODY)));
    };
    const LX = 0.6, LW = 4.0, RX = 4.8, RW = 4.6, TOP = 1.35;
    card("스킬 28개", "보고문·메일·문서 검토·계약 검토·주간보고·회고·회의록·정부사업 등록 등 업무 유형별 자작 스킬 (2026-10 전수감사 기준)", LX, TOP, LW, 1.7, "skills");
    card("예약 작업", "주간보고 초안(볼트·메일·메신저 교차 조회), 월간 공휴일·일정 체크, 현장조사 준비 체크리스트, 아침 할 일 점검을 자동 발송", LX, TOP + 1.85, LW, 1.7, "scheduled");
    card("연동", "한국 법령·미국 연방규정 원문, 옵시디언, 메일·캘린더·메신저·스프레드시트", RX, TOP, RW, 1.2, "integrations");
    const rules = [
      "외부 발송·제출의 최종 버튼은 사람이 누른다",
      "자료에 없는 숫자는 만들지 않고, 계산은 결과·과정·근거·주의사항 4필드로 보고한다",
      "답변에 붙인 교정을 모아 스킬에 반영한다 (월 1회 전수감사)",
    ];
    const RY = TOP + 1.35, RH = 2.2;
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y: RY, w: RW, h: RH, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: "Card rules" });
    s.addText("운영 규칙", T(null, { x: RX + 0.25, y: RY + 0.2, w: RW - 0.5, h: 0.4, fontSize: 22, bold: true, objectName: "Card keyword rules" }));
    // numbered rules as a borderless two-column table so each row sizes to its text
    const ruleRows = rules.map((r, i) => [
      { text: String(i + 1).padStart(2, "0"), options: { bold: true, color: C.accent1, valign: "top" } },
      { text: r, options: { color: C.text1, valign: "top" } },
    ]);
    s.addTable(ruleRows, { x: RX + 0.25, y: RY + 0.62, w: RW - 0.5, colW: [0.35, RW - 0.85], fontSize: 13, fontFace: "Malgun Gothic",
      border: { type: "none" }, margin: [0.03, 0.04, 0.05, 0.04], lineSpacingMultiple: 1.1, objectName: "Rules table" });
  }

  // ===== 14. 스킬·도구 =====
  pres.addSection({ title: "Skills" });
  {
    const s = contentSlide("14", "SKILLS & TOOLS", "스킬·도구", "Skills");
    const skills = [
      ["엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·다중 통화 자금 시뮬레이션 설계"],
      ["옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인, 주간파일·회신추적·SOP 연결"],
      ["AI 업무 자동화", "Claude 커스텀 스킬 28개 설계·운영, 예약 작업, MCP 연동(한국 법령·미국 연방규정·옵시디언)"],
      ["협업 도구", "Gmail, Google Drive·Sheets, Slack"],
      ["ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)"],
      ["문서 표준화", "공문 서식, docx·pptx 폰트 표준, SOW·인보이스·인사·급여 템플릿, SOP"],
      ["실무 지식", "결산·세무 신고·수정신고 대응, 주총·이사회·등기·CB·BW, 정부사업 전 주기, 기업부설연구소, 해외 송금 증빙·해외 법인 청산, 퇴직소득 과세이연"],
      ["언어", "비즈니스 영어(미국 현지 근무 2년 8개월)"],
    ];
    const rows = skills.map(([k, v], i) => {
      const o = i % 2 === 0 ? { fill: { color: C.background2 } } : {};
      return [
        { text: k, options: Object.assign({ bold: true, color: C.text1 }, o) },
        { text: v, options: Object.assign({ color: C.text1 }, o) },
      ];
    });
    s.addTable(rows, { x: 0.6, y: 1.45, w: 8.8, colW: [1.8, 7.0], rowH: 0.4, fontSize: 13, fontFace: "Malgun Gothic",
      border: { type: "none" }, margin: [0.05, 0.15, 0.05, 0.15], valign: "middle", objectName: "Skills table" });
  }

  // ===== 15. 타임라인 =====
  pres.addSection({ title: "Timeline" });
  {
    const s = contentSlide("15", "TIMELINE", "타임라인", "Timeline");
    const items = [
      ["2022.12 – 현재", "IT 스타트업", ["경영지원팀 매니저 (국내 법인 + 해외 자회사)"]],
      ["2019.06 – 2022.01", "Stylekorean, LLC (실리콘투 미국 자회사, 뉴저지)", ["재고관리 인턴 → Purchasing 담당", "재고 관리 매뉴얼 제작으로 재고 손실률 전년 대비 8% 이상 감소"]],
      ["2018.11 – 2019.04", "산업인력공단 글로벌청년인재개발", ["미국 비즈니스 취업연수과정 수료"]],
      ["2017", "중국 교환학생", []],
      ["2013 – 2019", "학부", ["국제통상학 전공 (중어중문학 복수전공)"]],
    ];
    let y = 1.45;
    items.forEach(([date, t, ds], i) => {
      s.addText(date, T(null, { x: 0.6, y: y + 0.03, w: 1.85, h: 0.3, fontSize: 13, bold: true, color: C.text2, objectName: "Date " + i }));
      s.addShape(pres.shapes.OVAL, { x: 2.6, y: y + 0.09, w: 0.16, h: 0.16, fill: { color: i === 0 ? C.accent1 : C.text2 }, line: { color: i === 0 ? C.accent1 : C.text2, width: 0 }, objectName: "Dot " + i });
      s.addText(t, T(null, { x: 3.0, y, w: 6.4, h: 0.34, fontSize: 16, bold: true, objectName: "Event " + i }));
      if (ds.length) s.addText(paraRuns(ds), T(null, { x: 3.0, y: y + 0.34, w: 6.4, h: 0.23 * ds.length + 0.05, fontSize: 12, color: C.text2, valign: "top", paraSpaceAfter: 2, objectName: "Event desc " + i }));
      y += 0.42 + 0.23 * ds.length + (ds.length ? 0.14 : 0.1);
    });
  }

  // ===== 16. 지향점·연락 =====
  pres.addSection({ title: "Contact" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Contact" });
    s.addText("NEXT · CONTACT", T(null, { x: 0.8, y: 0.95, w: 5, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 3, objectName: "Contact label" }));
    s.addText([{ text: "양남규" }, { text: ".", options: { color: C.accent1 } }], T(null, { x: 0.8, y: 1.25, w: 7, h: 1.2, fontSize: 72, bold: true, objectName: "Name" }));
    const blocks = [
      ["지향", "해외 자회사를 둔 회사의 재무 총괄", ic.target, 0.8],
      ["준비 중", "재경관리사 → AICPA", ic.award, 5.3],
    ];
    blocks.forEach(([h, t, img, x]) => {
      s.addShape(pres.shapes.OVAL, { x, y: 2.75, w: 0.44, h: 0.44, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: "Circle " + h });
      s.addImage({ data: img, x: x + 0.11, y: 2.86, w: 0.22, h: 0.22, objectName: "Icon " + h });
      heading(s, h, x + 0.6, 2.72, 3.6);
      s.addText(t, T(null, { x: x + 0.6, y: 2.97, w: 3.8, h: 0.4, fontSize: 18, bold: true, objectName: "Block " + h }));
    });
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: 3.7, w: 0.44, h: 0.44, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, objectName: "Mail circle" });
    s.addImage({ data: ic.mail, x: 0.91, y: 3.81, w: 0.22, h: 0.22, objectName: "Mail icon" });
    s.addText("namgyu1412@naver.com", T(null, { x: 1.4, y: 3.7, w: 6, h: 0.44, fontSize: 18, bold: true, valign: "middle", objectName: "Email" }));
    s.addText("© 2026 Duke Yang", T(null, { x: 0.8, y: 4.95, w: 4, h: 0.25, fontSize: 10, color: C.text2, objectName: "Copyright" }));
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch(e => { console.error(e); process.exit(1); });
