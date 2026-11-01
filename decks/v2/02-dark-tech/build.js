// Dark Tech — 양남규 포트폴리오 덱 (02-dark-tech) · 원고 v2 (14장)
// Run: NODE_PATH=<node_modules> node build.js   (writes deck.pptx next to this file)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fi = require("react-icons/fi");

const SKILL =
  process.env.PPTX_SKILL_DIR ||
  "/root/.claude/skills/synced/7d6f9636-edbe-41a2-aab0-0f84fceac3dd_307ca375-45f7-4235-8c53-4eacd5b231ab/pptx";
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));

const OUT = path.join(__dirname, "deck.pptx");

// ---------- Theme ----------
const THEME = {
  name: "Dark Tech",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "E6EDF7", // 본문 (text1)
    lt1: "0B1221", // 바탕 (background1)
    dk2: "8B9BB4", // 보조 (text2)
    lt2: "141C2F", // 카드 (background2)
    accent1: "2EE6A6", // 민트 포인트
    accent2: "4F8CFF", // 블루 포인트2
    accent3: "1B2742", // 칩/서브 패널
    accent4: "2A3A5A", // 글래스 테두리
    accent5: "0F172A", // 깊은 패널
    accent6: "11322B", // 민트 틴트 패널 (결과 박스)
    hlink: "4F8CFF",
    folHlink: "8B9BB4",
  },
};
const HX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "양남규";
pres.title = "양남규 포트폴리오 — Dark Tech";
const C = pres.SchemeColor;

const W = 10, H = 5.625, M = 0.5;
const CT = 1.15; // content top (1-line title)
const CT2 = 1.75; // content top (2-line title, case slides)
const CB = 5.25; // content bottom

// ---------- helpers ----------
async function iconPng(Icon, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color: "#" + hex, size })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function dotGridPng() {
  const w = 1920, h = 1080, step = 40;
  let dots = "";
  for (let y = step / 2; y < h; y += step)
    for (let x = step / 2; x < w; x += step)
      dots += `<circle cx="${x}" cy="${y}" r="1.6" fill="#${HX.accent4}"/>`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><rect width="${w}" height="${h}" fill="#${HX.lt1}"/>${dots}</svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// rough text width (inches) for Malgun-ish metrics
function estW(text, pt) {
  let em = 0;
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if (ch === " ") em += 0.3;
    else if (c > 0x2e80) em += 1.0; // CJK
    else if (/[A-Z0-9]/.test(ch)) em += 0.64;
    else if (/[a-z]/.test(ch)) em += 0.54;
    else em += 0.36;
  }
  return (em * pt) / 72;
}

function shadow() {
  return { type: "outer", color: "000000", blur: 8, offset: 3, angle: 90, opacity: 0.35 };
}

// motif: 둥근 모서리 글래스 카드 (어두운 카드 + 미세한 테두리)
function glass(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    rectRadius: opts.radius ?? 0.12,
    fill: { color: opts.fill || C.background2 },
    line: { color: opts.border || C.accent4, width: opts.borderWidth || 0.75 },
    shadow: opts.noShadow ? undefined : shadow(),
    objectName: opts.name || "card",
  });
  if (opts.dots) {
    const cols = [C.accent2, C.text2, C.accent1];
    cols.forEach((col, i) => {
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.2 + i * 0.17, y: y + 0.18, w: 0.1, h: 0.1,
        fill: { color: col }, line: { color: col, width: 0 },
        objectName: "window dot",
      });
    });
  }
  if (opts.header) {
    slide.addText(opts.header, {
      x: x + (opts.dots ? 0.8 : 0.25), y: y + 0.1, w: w - (opts.dots ? 1.05 : 0.5), h: 0.26,
      fontSize: 10, color: C.text2, margin: 0, valign: "middle", isTextBox: true,
      objectName: "card header",
    });
  }
}

// small label chip; returns width used
function chip(slide, x, y, text, opts = {}) {
  const pt = opts.fontSize || 10;
  const h = opts.h || 0.26;
  const w = opts.w || estW(text, pt) * 1.12 + 0.26;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: h / 2,
    fill: { color: opts.fill || C.accent3 },
    line: { color: opts.border || C.accent4, width: 0.75 },
    objectName: "chip",
  });
  slide.addText(text, {
    x, y, w, h, fontSize: pt, color: opts.color || C.accent1, bold: !!opts.bold,
    align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "chip label",
  });
  return w;
}

// flow chips in a box, wrapping by estimated width; returns total height used
function chipFlow(slide, x, y, maxW, items, opts = {}) {
  const gap = opts.gap ?? 0.08, h = opts.h || 0.26, pt = opts.fontSize || 10;
  let cx = x, cy = y;
  for (const t of items) {
    const w = estW(t, pt) * 1.12 + 0.26;
    if (cx + w > x + maxW + 0.01 && cx > x) { cx = x; cy += h + gap; }
    chip(slide, cx, cy, t, { ...opts, fontSize: pt, h, w });
    cx += w + gap;
  }
  return cy + h - y;
}

// step flow: rounded step boxes joined by mint arrows (single row); returns width used
function stepFlow(slide, x, y, items, opts = {}) {
  const pt = opts.fontSize || 11, h = opts.h || 0.34, aw = opts.arrowW || 0.32, pad = opts.pad || 0.3;
  let sx = x;
  items.forEach((t, i) => {
    const w = estW(t, pt) * 1.12 + pad;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: sx, y, w, h, rectRadius: 0.08,
      fill: { color: opts.fill || C.accent3 }, line: { color: opts.border || C.accent4, width: 0.75 },
      objectName: "step",
    });
    slide.addText(t, {
      x: sx, y, w, h, fontSize: pt, bold: true, color: opts.color || C.text1,
      align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "step label",
    });
    sx += w;
    if (i < items.length - 1) {
      slide.addText("→", {
        x: sx, y, w: aw, h, fontSize: pt + 3, bold: true, color: C.accent1,
        align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "arrow",
      });
      sx += aw;
    }
  });
  return sx - x;
}

function iconBox(slide, x, y, size, data, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: size, h: size, rectRadius: size * 0.28,
    fill: { color: opts.fill || C.accent3 },
    line: { color: opts.border || C.accent4, width: 0.75 },
    objectName: "icon box",
  });
  const pad = size * 0.24;
  slide.addImage({ data, x: x + pad, y: y + pad, w: size - 2 * pad, h: size - 2 * pad, objectName: opts.alt || "icon", altText: opts.alt || "icon" });
}

function label(slide, x, y, w, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h: opts.h || 0.28, fontSize: opts.fontSize || 11, bold: true,
    color: opts.color || C.accent1, margin: 0, valign: "middle", isTextBox: true,
    charSpacing: 1, objectName: "label",
  });
}

function body(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h, fontSize: opts.fontSize || 14, color: opts.color || C.text1,
    margin: 0, valign: opts.valign || "top", align: "left", isTextBox: true,
    lineSpacingMultiple: opts.ls || 1.0, bold: !!opts.bold, objectName: opts.name || "body text",
  });
}

// several paragraphs in one flowing box
function paras(slide, x, y, w, h, texts, opts = {}) {
  slide.addText(
    texts.map((t, i) => ({ text: t, options: { paraSpaceAfter: opts.gap ?? 8, breakLine: i < texts.length - 1 } })),
    { x, y, w, h, fontSize: opts.fontSize || 14, color: opts.color || C.text1, margin: 0, valign: "top", align: "left",
      isTextBox: true, lineSpacingMultiple: opts.ls || 1.0, objectName: opts.name || "body text" }
  );
}

// ---------- layouts ----------
async function defineLayouts() {
  const grid = await dotGridPng();
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: HX.lt1 },
    objects: [{ image: { x: 0, y: 0, w: W, h: H, data: grid } }],
  });
  const footer = { text: { text: "양남규 · 경영지원 매니저", options: { x: M, y: 5.32, w: 5, h: 0.22, fontSize: 9, color: C.text2, margin: 0, valign: "middle" } } };
  const eyebrow = { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.2, w: 6, h: 0.22, fontSize: 10, color: C.accent1, margin: 0, valign: "middle" }, text: "" } };
  const slideNumber = { x: 8.9, y: 5.32, w: 0.6, h: 0.22, fontSize: 9, color: HX.dk2, align: "right", margin: 0 };
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: HX.lt1 },
    objects: [
      eyebrow,
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.42, w: 9, h: 0.62, fontSize: 36, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
      footer,
    ],
    slideNumber,
  });
  // same title position, two-line title region (case slides)
  pres.defineSlideMaster({
    title: "CONTENT2",
    background: { color: HX.lt1 },
    objects: [
      eyebrow,
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.42, w: 9, h: 1.25, fontSize: 36, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
      footer,
    ],
    slideNumber,
  });
}

function contentSlide(section, eyebrow, title, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(title, { placeholder: "title" });
  return s;
}

// ---------- build ----------
(async () => {
  await defineLayouts();
  const mint = HX.accent1, blue = HX.accent2;
  const ic = {
    user: await iconPng(Fi.FiUser, mint),
    finance: await iconPng(Fi.FiDollarSign, mint),
    hr: await iconPng(Fi.FiUsers, mint),
    ir: await iconPng(Fi.FiTrendingUp, mint),
    gov: await iconPng(Fi.FiFileText, mint),
    globe: await iconPng(Fi.FiGlobe, mint),
    admin: await iconPng(Fi.FiBriefcase, mint),
    database: await iconPng(Fi.FiDatabase, mint),
    check: await iconPng(Fi.FiCheckCircle, mint),
    repeat: await iconPng(Fi.FiRepeat, mint),
    layers: await iconPng(Fi.FiLayers, mint),
    grid: await iconPng(Fi.FiGrid, mint),
    book: await iconPng(Fi.FiBookOpen, mint),
    search: await iconPng(Fi.FiSearch, mint),
    table: await iconPng(Fi.FiColumns, mint),
    folder: await iconPng(Fi.FiFolder, mint),
    cpu: await iconPng(Fi.FiCpu, mint),
    chat: await iconPng(Fi.FiMessageSquare, mint),
    server: await iconPng(Fi.FiServer, mint),
    doc: await iconPng(Fi.FiFileText, mint),
    mail: await iconPng(Fi.FiMail, mint),
    terminal: await iconPng(Fi.FiTerminal, blue),
    clock: await iconPng(Fi.FiClock, mint),
  };

  // ===== 1. 표지 =====
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    glass(s, 1.1, 0.75, 7.8, 4.1, { dots: true, header: "portfolio — 양남규", name: "terminal window" });
    s.addText([
      { text: "$ ", options: { color: C.accent2 } },
      { text: "whoami", options: { color: C.accent1 } },
    ], { x: 1.5, y: 1.22, w: 4, h: 0.3, fontSize: 13, bold: true, margin: 0, isTextBox: true, objectName: "prompt" });
    chip(s, 1.5, 1.6, "경영지원 매니저", { fontSize: 11, bold: true, h: 0.3 });
    s.addText("양남규", { x: 1.5, y: 1.95, w: 4.2, h: 0.95, fontSize: 56, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "name" });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.95, y: 2.15, w: 0.14, h: 0.6, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: "cursor" });
    s.addText("Duke Yang", { x: 1.5, y: 2.92, w: 6.5, h: 0.4, fontSize: 22, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "english name" });
    body(s, 1.5, 3.42, 6.9, 0.85,
      "IT 스타트업에서 재무·인사·IR·정부사업·해외업무·총무를 혼자 맡아 운영하며, 그 일을 AI와 지식관리 도구로 시스템화하는 경영지원 매니저입니다.",
      { fontSize: 14, ls: 1.1 });
    body(s, 1.5, 4.35, 6.8, 0.3, "기준일 2026-10-27", { fontSize: 11, color: C.text2 });
    iconBox(s, 7.45, 1.8, 1.0, ic.terminal, { alt: "terminal icon" });
  }

  // ===== 2. 일하는 원칙 =====
  pres.addSection({ title: "About" });
  {
    const s = contentSlide("About", "// principles", "일하는 원칙");
    const items = [
      { k: "시스템", t: "사람의 기억이 아닌 시스템에 남긴다.", icon: ic.database },
      { k: "원문 확인", t: "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보낸다.", icon: ic.check },
      { k: "스킬 · SOP", t: "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다.", icon: ic.repeat },
    ];
    const cw = 2.93, gap = (9 - 3 * cw) / 2, y = CT, h = CB - CT;
    items.forEach((it, i) => {
      const x = M + i * (cw + gap);
      glass(s, x, y, cw, h);
      s.addText(String(i + 1).padStart(2, "0"), { x: x + cw - 0.8, y: y + 0.22, w: 0.55, h: 0.3, fontSize: 12, bold: true, color: C.text2, align: "right", margin: 0, isTextBox: true, objectName: "index" });
      iconBox(s, x + 0.3, y + 0.45, 0.7, it.icon, { alt: it.k + " icon" });
      s.addText(it.k, { x: x + 0.3, y: y + 1.55, w: cw - 0.6, h: 0.6, fontSize: 28, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "keyword" });
      body(s, x + 0.3, y + 2.35, cw - 0.6, h - 2.6, it.t, { fontSize: 15, ls: 1.1 });
    });
  }

  // ===== 3. 소개 =====
  {
    const s = contentSlide("About", "// about", "소개");
    const lw = 5.9, rx = M + lw + 0.2, rw = 9 - lw - 0.2;
    glass(s, M, CT, lw, CB - CT, { dots: true, header: "about.md" });
    paras(s, M + 0.3, CT + 0.6, lw - 0.6, CB - CT - 0.8, [
      "IT 스타트업 경영지원팀 매니저로 일하며 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다.",
      "국제통상학을 전공했고, 미국 현지 유통기업에서 근무한 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다.",
      "여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해 옵시디언 기반 업무 운영체계와 AI 자동화 워크플로를 직접 설계해 쓰고 있습니다.",
      "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다.",
    ], { fontSize: 14, ls: 1.1, gap: 9, name: "introduction" });
    glass(s, rx, CT, rw, CB - CT, { header: "profile" });
    iconBox(s, rx + rw / 2 - 0.45, CT + 0.55, 0.9, ic.user, { alt: "user icon" });
    s.addText("양남규", { x: rx, y: CT + 1.55, w: rw, h: 0.34, fontSize: 17, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true, objectName: "name" });
    s.addText("Duke Yang · 경영지원 매니저", { x: rx, y: CT + 1.9, w: rw, h: 0.26, fontSize: 11, color: C.text2, align: "center", margin: 0, isTextBox: true, objectName: "role" });
    chipFlow(s, rx + 0.3, CT + 2.4, rw - 0.6, ["재무", "인사", "법무", "정부사업", "IR", "해외업무"], { fontSize: 11, h: 0.3, gap: 0.09 });
  }

  // ===== 4. 담당 영역 =====
  pres.addSection({ title: "Domains" });
  {
    const s = contentSlide("Domains", "// domains", "담당 영역");
    body(s, M, 1.07, 9, 0.28, "여섯 영역을 한 사람이 맡아 운영합니다", { fontSize: 13, color: C.text2, valign: "middle" });
    const cards = [
      { t: "재무·회계", d: "자금 운용, 세무 전 주기, 결산·외부감사", icon: ic.finance },
      { t: "인사·노무", d: "급여·4대보험, 채용,\n해외 원격 인력 계약", icon: ic.hr },
      { t: "IR·투자", d: "주총·이사회, 등기, 스톡옵션, 투자 계약", icon: ic.ir },
      { t: "정부사업", d: "공고 분석부터\n정산까지 전 주기", icon: ic.gov },
      { t: "해외업무", d: "해외 자회사 증빙,\n해외 행사·출장", icon: ic.globe },
      { t: "총무", d: "법인 행정, 기업 인증, 사무실·자산, 대외행사", icon: ic.admin },
    ];
    const top = 1.45, gx = 0.12, gy = 0.14, cw = (9 - 2 * gx) / 3, ch = (CB - top - gy) / 2;
    cards.forEach((c, i) => {
      const x = M + (i % 3) * (cw + gx), y = top + Math.floor(i / 3) * (ch + gy);
      glass(s, x, y, cw, ch);
      s.addText(String(i + 1).padStart(2, "0"), { x: x + cw - 0.75, y: y + 0.18, w: 0.55, h: 0.3, fontSize: 12, bold: true, color: C.text2, align: "right", margin: 0, isTextBox: true, objectName: "index" });
      iconBox(s, x + 0.25, y + 0.2, 0.55, c.icon, { alt: c.t + " icon" });
      s.addText(c.t, { x: x + 0.25, y: y + 0.85, w: cw - 0.5, h: 0.34, fontSize: 17, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "card title" });
      body(s, x + 0.25, y + 1.2, cw - 0.5, ch - 1.3, c.d, { fontSize: 14, ls: 1.05 });
    });
  }

  // ===== 5. 담당 영역 상세 1 =====
  {
    const s = contentSlide("Domains", "// domains · detail 1", "담당 영역 상세 1");
    const cols = [
      { t: "재무·회계", icon: ic.finance, p: [
        "자금 운용과 송금·결제를 관리하고, 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응합니다.",
        "결산과 외부감사 대응, 별도·연결 재무제표 관리를 맡고 외부 회계법인과의 커뮤니케이션을 담당합니다." ] },
      { t: "인사·노무", icon: ic.hr, p: [
        "급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다.",
        "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다." ] },
      { t: "IR·투자", icon: ic.ir, p: [
        "정기주주총회·이사회 운영, 정관 개정, 등기, 주주명부와 스톡옵션 관리를 맡고 있습니다.",
        "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다." ] },
    ];
    const cw = 2.93, gap = (9 - 3 * cw) / 2, y = CT, h = CB - CT;
    cols.forEach((c, i) => {
      const x = M + i * (cw + gap);
      glass(s, x, y, cw, h);
      iconBox(s, x + 0.25, y + 0.25, 0.5, c.icon, { alt: c.t + " icon" });
      s.addText(c.t, { x: x + 0.88, y: y + 0.25, w: cw - 1.1, h: 0.5, fontSize: 17, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "card title" });
      paras(s, x + 0.25, y + 0.98, cw - 0.5, h - 1.15, c.p, { fontSize: 14, ls: 1.05, gap: 8 });
    });
  }

  // ===== 6. 담당 영역 상세 2 =====
  {
    const s = contentSlide("Domains", "// domains · detail 2", "담당 영역 상세 2");
    const rows = [
      { t: "정부사업", icon: ic.gov, h: 1.78, flow: ["공고 분석", "사업계획서 작성", "신청", "협약", "정산"], p: [
        "공고 분석 → 사업계획서 작성 → 신청 → 협약 → 정산까지 전 주기를 실무로 수행합니다.",
        "입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다." ] },
      { t: "해외업무", icon: ic.globe, h: 0.92, p: [
        "해외 자회사의 인보이스와 용역계약 증빙을 관리하고, 해외 사업 행사 참가와 해외 출장 서류 실무를 지원합니다." ] },
      { t: "총무", icon: ic.admin, h: 1.16, p: [
        "본점 이전 같은 법인 행정 절차와 기업 인증 관리, 사무실·자산 관리를 담당합니다.",
        "대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다." ] },
    ];
    const gap = 0.12, lw = 1.9, tx = M + lw, tw = 9 - lw - 0.25;
    let y = CT;
    rows.forEach((r) => {
      glass(s, M, y, 9, r.h);
      iconBox(s, M + 0.25, y + 0.2, 0.5, r.icon, { alt: r.t + " icon" });
      s.addText(r.t, { x: M + 0.88, y: y + 0.2, w: lw - 0.95, h: 0.5, fontSize: 17, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "row title" });
      let ty = y + 0.2;
      if (r.flow) {
        stepFlow(s, tx, ty, r.flow, { fontSize: 11, h: 0.34 });
        ty += 0.48;
      }
      paras(s, tx, ty, tw, y + r.h - ty - 0.12, r.p, { fontSize: 14, ls: 1.05, gap: 6 });
      y += r.h + gap;
    });
  }

  // ===== 7–11. 사례 =====
  pres.addSection({ title: "Cases" });
  const cases = [
    {
      no: "01", title: "1인 경영지원을 \"인계 가능한 시스템\"으로 바꾸기", icon: ic.layers,
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 커스텀 AI 스킬로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다.",
    },
    {
      no: "02", title: "단일 통화 자금 시트를 의사결정 도구로 재설계", icon: ic.grid,
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리하고 있어, 경영진의 자금 판단 질문에 바로 답하기 어려웠습니다.",
      did: "대시보드·통화별 입출금·정기 비용·미지급 항목으로 시트를 나누고, 다중 통화 환산과 \"부족액 식별 → 송금액 계산 → 송금 후 잔액\" 시뮬레이션을 자동화했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.",
      flow: ["부족액 식별", "송금액 계산", "송금 후 잔액"],
      result: "경영진의 자금 의사결정 질문에 한 화면에서 바로 답할 수 있는 상태가 됐습니다.",
    },
    {
      no: "03", title: "법률 1차 검토를 법령 원문 기반으로 내재화", icon: ic.book,
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했습니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 실무 상식만으로는 놓칠 수 있는 별도 조항을 미리 잡게 됐습니다.",
    },
    {
      no: "04", title: "외부로 나가기 전 \"받은 산출물\"을 역검증하는 습관", icon: ic.search,
      bg: "회계법인·기관·플랫폼에서 받은 자료나 내부 문서가 확인 없이 외부로 나가면 세금·등기·평가에 바로 영향을 줍니다.",
      did: "회계법인 산출물, 기관·플랫폼 자료, 계약서와 내부 문서를 원장·법령 원문과 대조한 뒤 제출하는 절차를 분기 내내 운영했습니다.",
      result: "기간 오류, 견적 오기, 정관과 계약의 충돌 같은 문제를 외부 제출 전에 찾아냈습니다.",
    },
    {
      no: "05", title: "예산 대비 실적 관리 워크북 구축", icon: ic.table,
      bg: "월별 지출 추적이 여러 엑셀 파일에 흩어져 있어 예산을 통제하기 어려웠습니다.",
      did: "실제 계정원장을 기준으로 계정과목을 정리해 카테고리로 묶었습니다. 월별 예산 수립 → 실적 입력 → 차이 분석 흐름을 자동화했고, 이후 중앙 가정 시트와 연동되도록 확장했습니다.",
      flow: ["예산 수립", "실적 입력", "차이 분석"],
      result: "예산 대비 실적을 한 체계에서 관리하는 관리회계 기반을 만들었습니다.",
    },
  ];
  for (const cs of cases) {
    const s = contentSlide("Cases", `// case_${cs.no} · 사례 ${Number(cs.no)}`, cs.title, "CONTENT2");
    const y0 = CT2, hAll = CB - y0;
    const rh = 1.05, gap = 0.14, th = hAll - rh - gap;
    const c1 = 3.1, c2 = 9 - c1 - 0.12, x2 = M + c1 + 0.12;
    // 배경
    glass(s, M, y0, c1, th);
    label(s, M + 0.22, y0 + 0.16, c1 - 0.44, "배경");
    body(s, M + 0.22, y0 + 0.52, c1 - 0.44, th - 0.65, cs.bg, { fontSize: 14, ls: 1.05 });
    // 한 일
    glass(s, x2, y0, c2, th);
    label(s, x2 + 0.22, y0 + 0.16, c2 - 0.44, "한 일");
    iconBox(s, x2 + c2 - 0.72, y0 + 0.14, 0.5, cs.icon, { alt: "case icon" });
    const didH = cs.flow ? th - 1.2 : th - 0.65;
    body(s, x2 + 0.22, y0 + 0.52, c2 - 1.1, didH, cs.did, { fontSize: 14, ls: 1.05 });
    if (cs.flow) stepFlow(s, x2 + 0.22, y0 + th - 0.56, cs.flow, { fontSize: 11, h: 0.34 });
    // 결과 — tinted box
    const ry = y0 + th + gap;
    glass(s, M, ry, 9, rh, { fill: C.accent6, border: C.accent1, name: "result box" });
    label(s, M + 0.22, ry + 0.14, 0.9, "결과", { fontSize: 11 });
    body(s, M + 1.1, ry + 0.14, 9 - 1.35, rh - 0.28, cs.result, { fontSize: 15, bold: true, ls: 1.05, valign: "middle", name: "result" });
  }

  // ===== 12. 스킬·도구 =====
  pres.addSection({ title: "Skills" });
  {
    const s = contentSlide("Skills", "// skills", "스킬·도구");
    const rows = [
      ["엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·자금 시뮬레이션 설계", ic.table],
      ["옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인", ic.folder],
      ["AI 업무 자동화", "Claude 커스텀 스킬 설계·운영, 예약 작업 자동화, MCP 연동(법령 원문 조회, 옵시디언 연동)", ic.cpu],
      ["협업 도구", "Gmail, Google Drive, Slack", ic.chat],
      ["ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)", ic.server],
      ["문서 표준화", "공문 서식, docx·pptx 폰트 표준, 인보이스·인사·급여 템플릿, SOP", ic.doc],
      ["실무 지식", "결산·세무 신고 대응, 주총·이사회·등기, 정부사업 전 주기", ic.book],
      ["언어", "비즈니스 영어(미국 현지 근무 경험)", ic.globe],
    ];
    const pitch = (CB - CT) / 8, rh = pitch - 0.07;
    rows.forEach(([n, d, icon], i) => {
      const y = CT + i * pitch;
      glass(s, M, y, 9, rh, { noShadow: true, fill: i % 2 ? C.background2 : C.accent5, radius: 0.08 });
      iconBox(s, M + 0.15, y + (rh - 0.32) / 2, 0.32, icon, { alt: n + " icon" });
      s.addText(n, { x: M + 0.6, y, w: 1.75, h: rh, fontSize: 14, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "skill name" });
      body(s, M + 2.4, y, 9 - 2.55, rh, d, { fontSize: 13, valign: "middle" });
    });
  }

  // ===== 13. 타임라인 =====
  pres.addSection({ title: "Career" });
  {
    const s = contentSlide("Career", "// timeline", "타임라인");
    const items = [
      ["2022 – 현재", "IT 스타트업", "경영지원팀 매니저"],
      ["2019 – 2022", "미국 소재 유통기업", "재고관리 인턴 → 구매(Purchasing) 담당"],
      ["2018 – 2019", "미국 비즈니스 취업연수 과정 수료", ""],
      ["2017", "중국 교환학생", ""],
      ["2013 – 2019", "학부", "국제통상학 전공 (중어중문학 복수전공)"],
    ];
    const pitch = (CB - CT) / 5, ch = pitch - 0.1;
    const lineX = 1.82;
    s.addShape(pres.shapes.LINE, { x: lineX, y: CT + 0.2, w: 0, h: CB - CT - 0.4, line: { color: C.accent4, width: 1.5 }, objectName: "timeline" });
    items.forEach(([d, t, desc], i) => {
      const y = CT + i * pitch;
      s.addText(d, { x: M, y, w: 1.2, h: ch, fontSize: 12, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "date" });
      s.addShape(pres.shapes.OVAL, { x: lineX - 0.08, y: y + ch / 2 - 0.08, w: 0.16, h: 0.16, fill: { color: i === 0 ? C.accent1 : C.background2 }, line: { color: C.accent1, width: 1.5 }, objectName: "node" });
      const cx = 2.08, cw = 9.5 - cx;
      glass(s, cx, y, cw, ch);
      s.addText(t, { x: cx + 0.25, y, w: 3.3, h: ch, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "organization" });
      if (desc) body(s, cx + 3.65, y + 0.06, cw - 3.85, ch - 0.12, desc, { fontSize: 14, color: C.text2, valign: "middle" });
    });
  }

  // ===== 14. 연락 — 마무리 =====
  pres.addSection({ title: "Outro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Outro" });
    glass(s, 1.1, 0.9, 7.8, 3.85, { dots: true, header: "contact", name: "terminal window" });
    s.addText([
      { text: "$ ", options: { color: C.accent2 } },
      { text: "contact", options: { color: C.accent1 } },
    ], { x: 1.5, y: 1.4, w: 4, h: 0.3, fontSize: 13, bold: true, margin: 0, isTextBox: true, objectName: "prompt" });
    s.addText("연락", { x: 1.5, y: 1.8, w: 5.5, h: 0.9, fontSize: 48, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "closing title" });
    const w = 4.2, h = 0.55, y = 2.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y, w, h, rectRadius: 0.1, fill: { color: C.accent3 }, line: { color: C.accent4, width: 0.75 }, objectName: "contact" });
    s.addImage({ data: ic.mail, x: 1.5 + 0.2, y: y + 0.155, w: 0.24, h: 0.24, altText: "mail icon", objectName: "mail icon" });
    s.addText("namgyu1412@naver.com", { x: 1.5 + 0.58, y, w: w - 0.7, h, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "email" });
    body(s, 1.5, 3.75, 6.8, 0.3, "© 2026 Duke Yang", { fontSize: 11, color: C.text2 });
    iconBox(s, 7.45, 1.95, 1.0, ic.terminal, { alt: "terminal icon" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
