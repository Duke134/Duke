// Dark Tech — 양남규 포트폴리오 덱 (02-dark-tech)
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
    accent6: "7FF0C8", // 연민트
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
const CT = 1.15; // content top
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
function estLines(text, pt, widthIn) {
  return Math.max(1, Math.ceil(estW(text, pt) / widthIn));
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
    line: { color: opts.border || C.accent4, width: 0.75 },
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

// ---------- layouts ----------
async function defineLayouts() {
  const grid = await dotGridPng();
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: HX.lt1 },
    objects: [{ image: { x: 0, y: 0, w: W, h: H, data: grid } }],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: HX.lt1 },
    objects: [
      { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.2, w: 6, h: 0.22, fontSize: 10, color: C.accent1, margin: 0, valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.42, w: 9, h: 0.62, fontSize: 36, bold: true, color: C.text1, margin: 0, valign: "middle", align: "left" }, text: "" } },
      { text: { text: "양남규 · Finance Operations Specialist", options: { x: M, y: 5.32, w: 5, h: 0.22, fontSize: 9, color: C.text2, margin: 0, valign: "middle" } } },
    ],
    slideNumber: { x: 8.9, y: 5.32, w: 0.6, h: 0.22, fontSize: 9, color: HX.dk2, align: "right", margin: 0 },
  });
}

function contentSlide(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
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
    ir: await iconPng(Fi.FiTrendingUp, mint),
    hr: await iconPng(Fi.FiUsers, mint),
    gov: await iconPng(Fi.FiShield, mint),
    search: await iconPng(Fi.FiSearch, mint),
    grid: await iconPng(Fi.FiGrid, mint),
    alert: await iconPng(Fi.FiAlertTriangle, mint),
    cpu: await iconPng(Fi.FiCpu, mint),
    award: await iconPng(Fi.FiAward, mint),
    book: await iconPng(Fi.FiBookOpen, mint),
    mail: await iconPng(Fi.FiMail, mint),
    linkedin: await iconPng(Fi.FiLinkedin, mint),
    terminal: await iconPng(Fi.FiTerminal, blue),
    zap: await iconPng(Fi.FiZap, mint),
    clock: await iconPng(Fi.FiClock, mint),
  };

  // ===== 1. 표지 =====
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    glass(s, 1.1, 0.9, 7.8, 3.85, { dots: true, header: "portfolio — 양남규", name: "terminal window" });
    s.addText([
      { text: "$ ", options: { color: C.accent2 } },
      { text: "whoami", options: { color: C.accent1 } },
    ], { x: 1.5, y: 1.4, w: 4, h: 0.3, fontSize: 13, bold: true, margin: 0, isTextBox: true, objectName: "prompt" });
    body(s, 1.5, 1.78, 4, 0.35, "Hello, I'm", { fontSize: 16, color: C.text2 });
    s.addText("양남규", { x: 1.5, y: 2.1, w: 4.2, h: 0.95, fontSize: 56, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "name" });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.95, y: 2.3, w: 0.14, h: 0.6, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: "cursor" });
    s.addText("Finance Operations Specialist", { x: 1.5, y: 3.1, w: 6.5, h: 0.4, fontSize: 22, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "role" });
    body(s, 1.5, 3.55, 6.8, 0.35, "스타트업의 성장을 뒷받침하는 경영지원 전문가", { fontSize: 15 });
    body(s, 1.5, 3.98, 6.8, 0.3, "재무 · IR · HR · 컴플라이언스  /  GXC 경영지원팀  /  2026", { fontSize: 12, color: C.text2 });
    iconBox(s, 7.45, 1.95, 1.0, ic.terminal, { alt: "terminal icon" });
  }

  // ===== 2. About Me =====
  pres.addSection({ title: "About" });
  {
    const s = contentSlide("About", "// about", "About Me");
    glass(s, M, CT, 5.5, CB - CT, { dots: true, header: "about.md" });
    body(s, 0.8, 1.75, 4.9, 3.3,
      "GXC에서 경영지원 업무를 담당하며, 재무/회계 관리부터 투자 관련 법률 검토, 인사/노무 행정, 정부사업 관리까지 폭넓은 영역을 수행하고 있습니다. 체계적인 학습과 실무 경험을 병행하며 스타트업의 성장에 필요한 오퍼레이션 역량을 꾸준히 쌓아가고 있습니다.",
      { fontSize: 16, ls: 1.15 });
    const rx = 6.2, rw = 3.3;
    glass(s, rx, CT, rw, CB - CT, { header: "tags[]" });
    iconBox(s, rx + rw / 2 - 0.45, 1.65, 0.9, ic.user, { alt: "user icon" });
    s.addText("양남규", { x: rx, y: 2.62, w: rw, h: 0.3, fontSize: 15, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true, objectName: "name" });
    s.addText("GXC 경영지원팀", { x: rx, y: 2.9, w: rw, h: 0.25, fontSize: 11, color: C.text2, align: "center", margin: 0, isTextBox: true, objectName: "team" });
    chipFlow(s, rx + 0.3, 3.35, rw - 0.6, ["재무/회계", "IR/투자", "HR/노무", "정부사업", "Excel 고급", "데이터분석", "컴플라이언스"], { fontSize: 11, h: 0.3, gap: 0.09 });
  }

  // ===== 3. 업무 영역 =====
  {
    const s = contentSlide("About", "// domains", "업무 영역");
    const cards = [
      { t: "재무 / 회계", icon: ic.finance, items: ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"] },
      { t: "IR / 투자", icon: ic.ir, items: ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"] },
      { t: "HR / 노무", icon: ic.hr, items: ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"] },
      { t: "정부사업 / 기술", icon: ic.gov, items: ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"] },
    ];
    const cw = 2.1, gap = (9 - 4 * cw) / 3;
    cards.forEach((c, i) => {
      const x = M + i * (cw + gap), y = CT, h = CB - CT;
      glass(s, x, y, cw, h);
      s.addText(String(i + 1).padStart(2, "0"), { x: x + cw - 0.75, y: y + 0.2, w: 0.55, h: 0.3, fontSize: 12, bold: true, color: C.text2, align: "right", margin: 0, isTextBox: true, objectName: "index" });
      iconBox(s, x + 0.25, y + 0.3, 0.7, c.icon, { alt: c.t + " icon" });
      s.addText(c.t, { x: x + 0.25, y: y + 1.2, w: cw - 0.4, h: 0.4, fontSize: 17, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "card title" });
      c.items.forEach((it, j) => {
        const iy = y + 1.85 + j * 0.5;
        s.addShape(pres.shapes.RECTANGLE, { x: x + 0.27, y: iy + 0.12, w: 0.09, h: 0.09, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: "bullet" });
        body(s, x + 0.47, iy, cw - 0.65, 0.34, it, { fontSize: 14, valign: "middle" });
      });
    });
  }

  // ===== 4. Skills =====
  pres.addSection({ title: "Skills" });
  {
    const s = contentSlide("Skills", "// skills", "Skills & Expertise");
    const groups = [
      { t: "재무/회계", labels: ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], values: [75, 70, 80, 65] },
      { t: "IR/투자", labels: ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], values: [70, 65, 60, 60] },
      { t: "HR/노무", labels: ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], values: [75, 70, 65, 60] },
      { t: "정부사업/기술", labels: ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], values: [70, 65, 85, 60] },
    ];
    const cw = 4.4, ch = 2.0, gx = 0.2, gy = 0.1;
    groups.forEach((g, i) => {
      const x = M + (i % 2) * (cw + gx), y = CT + Math.floor(i / 2) * (ch + gy);
      glass(s, x, y, cw, ch);
      label(s, x + 0.25, y + 0.12, 3, g.t, { fontSize: 12, color: C.accent1 });
      s.addText("숙련도 (%)", { x: x + cw - 1.5, y: y + 0.12, w: 1.25, h: 0.28, fontSize: 9, color: C.text2, align: "right", margin: 0, valign: "middle", isTextBox: true, objectName: "unit" });
      s.addChart(pres.charts.BAR, [{ name: g.t, labels: g.labels, values: g.values }], {
        x: x + 0.15, y: y + 0.42, w: cw - 0.3, h: ch - 0.5,
        barDir: "bar", barGapWidthPct: 45,
        chartColors: [HX.accent1],
        chartArea: { fill: { color: HX.lt2 } },
        plotArea: { fill: { color: HX.lt2 } },
        showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HX.dk1, dataLabelFontSize: 10, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: '0"%"',
        catAxisLabelColor: HX.dk1, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin", catAxisLineShow: false,
        catGridLine: { style: "none" }, valGridLine: { style: "none" },
        valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valAxisLineShow: false,
        showLegend: false, showTitle: false,
      });
    });
  }

  // ===== 5. Excel Mastery =====
  {
    const s = contentSlide("Skills", "// excel", "Excel Mastery");
    body(s, M, 1.07, 9, 0.28, "실무에서 활용하는 엑셀 고급 함수와 자동화 기법", { fontSize: 12, color: C.text2, valign: "middle" });
    const rows = [
      { n: "XLOOKUP", icon: ic.search, use: "급여 대장 자동 매칭 / 거래처별 매출 추적 / 계정과목 자동 분류", badges: ["VLOOKUP 대체", "양방향 검색", "에러 처리 내장"] },
      { n: "피벗테이블", icon: ic.grid, use: "부서별 예산 집행 분석 / 월별 비용 추이 리포트", badges: ["크로스탭 분석", "슬라이서 필터", "자동 갱신"] },
      { n: "조건부 서식", icon: ic.alert, use: "예산 초과 자동 알림 (달성률 100% 초과 빨강, 90% 미만 노랑)", badges: ["데이터 막대", "아이콘 세트", "색조 규칙"] },
      { n: "자동화 매크로(VBA)", icon: ic.cpu, use: "월말 결산 자동화 / 급여명세서 자동 발송", badges: ["VBA 매크로", "반복 업무 제거", "3시간→15분"] },
    ];
    const top = 1.45, pitch = (CB - top) / 4, rh = pitch - 0.1;
    rows.forEach((r, i) => {
      const y = top + i * pitch;
      glass(s, M, y, 9, rh);
      iconBox(s, 0.7, y + (rh - 0.5) / 2, 0.5, r.icon, { alt: r.n + " icon" });
      s.addText(String(i + 1).padStart(2, "0"), { x: 1.35, y: y + 0.1, w: 0.4, h: 0.22, fontSize: 9, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "index" });
      s.addText(r.n, { x: 1.35, y: y + 0.3, w: 2.2, h: rh - 0.4, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "technique" });
      body(s, 3.6, y + 0.08, 3.1, rh - 0.16, r.use, { fontSize: 14, valign: "middle" });
      const bh = chipFlowHeight(r.badges, 2.5, 10, 0.25, 0.07);
      chipFlow(s, 6.9, y + (rh - bh) / 2, 2.5, r.badges, { fontSize: 10, h: 0.25, gap: 0.07 });
    });
  }

  // ===== 6. 자동화 성과 =====
  {
    const s = contentSlide("Skills", "// automation", "자동화 성과");
    const lw = 4.4, rx = 5.1, rw = 4.4;
    // stat 1
    glass(s, M, CT, lw, 1.45);
    s.addText("3시간 → 15분", { x: 0.75, y: CT + 0.15, w: 3.9, h: 0.6, fontSize: 30, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "stat" });
    body(s, 0.75, CT + 0.8, 3.9, 0.3, "월말 결산 보고서 생성", { fontSize: 14, valign: "middle" });
    body(s, 0.75, CT + 1.08, 3.9, 0.26, "전표 정리 → 집계 → 보고서, 버튼 하나", { fontSize: 11, color: C.text2, valign: "middle" });
    // stat 2
    const y2 = CT + 1.57;
    glass(s, M, y2, lw, 1.45);
    s.addText("3일 → 1.2일", { x: 0.75, y: y2 + 0.15, w: 2.6, h: 0.6, fontSize: 30, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "stat" });
    chip(s, 3.4, y2 + 0.3, "약 60% 단축", { fontSize: 10, bold: true, h: 0.3 });
    body(s, 0.75, y2 + 0.85, 3.9, 0.3, "월말 결산 전체 소요 시간", { fontSize: 14, valign: "middle" });
    // pivot table
    glass(s, rx, CT, rw, 3.02, { dots: true, header: "예시 피벗 표 — 부서 / Q1 / Q2 / 증감률" });
    const hdr = (t) => ({ text: t, options: { bold: true, color: HX.accent1, fill: { color: HX.accent3 }, fontSize: 11, align: t === "부서" ? "left" : "right" } });
    const cell = (t, o = {}) => ({ text: t, options: { color: o.color || HX.dk1, fontSize: 12, bold: !!o.bold, align: o.align || "right", fill: { color: o.fill || HX.lt2 } } });
    const rowsT = [
      [hdr("부서"), hdr("Q1"), hdr("Q2"), hdr("증감률")],
      [cell("개발팀", { align: "left" }), cell("45,200"), cell("52,800"), cell("+16.8%", { color: HX.accent1, bold: true })],
      [cell("마케팅", { align: "left" }), cell("32,100"), cell("28,500"), cell("-11.2%", { color: HX.accent2, bold: true })],
      [cell("경영지원", { align: "left" }), cell("18,300"), cell("19,100"), cell("+4.4%", { color: HX.accent1, bold: true })],
      [cell("합계", { align: "left", bold: true, fill: HX.accent5 }), cell("95,600", { bold: true, fill: HX.accent5 }), cell("100,400", { bold: true, fill: HX.accent5 }), cell("+5.0%", { color: HX.accent1, bold: true, fill: HX.accent5 })],
    ];
    s.addTable(rowsT, {
      x: rx + 0.25, y: CT + 0.55, w: rw - 0.5, colW: [1.1, 0.95, 0.95, 0.9], rowH: 0.42,
      border: { type: "solid", color: HX.accent4, pt: 0.75 }, margin: [0.04, 0.1, 0.04, 0.1], valign: "middle",
      objectName: "pivot table",
    });
    // 3단계
    const y3 = CT + 3.15, h3 = CB - y3;
    glass(s, M, y3, 9, h3);
    s.addText("3단계", { x: 0.75, y: y3 + 0.1, w: 1.3, h: h3 - 0.2, fontSize: 26, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "stat" });
    const steps = [["전표 정리", 1.5], ["계정과목별 집계", 2.0], ["보고서 출력 (PDF 저장 · 메일 발송)", 3.1]];
    let sx = 2.3;
    steps.forEach(([t, w], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: sx, y: y3 + 0.22, w, h: h3 - 0.44, rectRadius: 0.08, fill: { color: C.accent3 }, line: { color: C.accent4, width: 0.75 }, objectName: "step" });
      s.addText(t, { x: sx, y: y3 + 0.22, w, h: h3 - 0.44, fontSize: 12, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "step label" });
      sx += w;
      if (i < steps.length - 1) {
        s.addText("→", { x: sx, y: y3 + 0.22, w: 0.4, h: h3 - 0.44, fontSize: 16, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "arrow" });
        sx += 0.4;
      }
    });
  }

  // ===== 7–9. Cases =====
  pres.addSection({ title: "Cases" });
  const cases = [
    {
      eyebrow: "// case_01 · IR/투자", title: "SHA 주주간계약 검토", tag: "IR/투자", no: "Case 01",
      bg: "GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.",
      steps: ["ROFR(우선매수권) 조항의 실무 적용 시나리오 분석", "Drag-Along · Tag-Along 권리의 행사 조건 정리", "희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)", "투자자 동의권(Consent Rights) 범위 및 영향 분석", "법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화"],
      result: "5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.",
      kpis: [["5개", "핵심 조항 분석"], ["1건", "가이드 문서"]],
      learn: "법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다.",
    },
    {
      eyebrow: "// case_02 · 재무/회계", title: "월말 결산 프로세스 혁신", tag: "재무/회계", no: "Case 02",
      bg: "매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.",
      steps: ["기존 결산 프로세스 전수조사 및 병목 구간 파악", "전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)", "계정과목별 자동 집계 템플릿 구축", "VBA 매크로를 활용한 보고서 자동 생성", "결산 체크리스트 및 SOP 문서 작성"],
      result: "결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.",
      kpis: [["-60%", "소요 시간"], ["SOP", "프로세스 표준화"]],
      learn: "Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다.",
    },
    {
      eyebrow: "// case_03 · HR/노무", title: "TTC 기반 보상 체계 설계", tag: "HR/노무", no: "Case 03",
      bg: "스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.",
      steps: ["국내 IT 스타트업 보상 시장조사 및 벤치마킹", "TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)", "직급/직무별 Pay Band 설계", "스톡옵션 배분 기준 및 베스팅 스케줄 수립", "경영진 보고 및 의사결정 자료 준비"],
      result: "전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.",
      kpis: [["전사", "적용 범위"], ["TTC", "체계 도입"]],
      learn: "보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다.",
    },
  ];
  for (const cs of cases) {
    const s = contentSlide("Cases", cs.eyebrow, cs.title);
    // tag chips on the title row (right)
    const cw2 = estW(cs.tag, 10) * 1.12 + 0.26;
    const cw1 = estW(cs.no, 10) * 1.12 + 0.26;
    chip(s, 9.5 - cw2, 0.58, cs.tag, { fontSize: 10, bold: true, h: 0.3 });
    chip(s, 9.5 - cw2 - 0.1 - cw1, 0.58, cs.no, { fontSize: 10, bold: true, h: 0.3, color: C.text2 });

    const c1 = 2.5, c2 = 3.3, c3 = 2.9, g = (9 - c1 - c2 - c3) / 2;
    const x1 = M, x2 = M + c1 + g, x3 = x2 + c2 + g;
    const hAll = CB - CT;
    // col1: 배경 + 지표
    const kh = 1.05;
    glass(s, x1, CT, c1, hAll - kh - 0.12);
    label(s, x1 + 0.2, CT + 0.15, c1 - 0.4, "배경");
    body(s, x1 + 0.2, CT + 0.5, c1 - 0.4, hAll - kh - 0.12 - 0.6, cs.bg, { fontSize: 14, ls: 0.95 });
    const ky = CB - kh;
    glass(s, x1, ky, c1, kh, { fill: C.accent5 });
    label(s, x1 + 0.2, ky + 0.1, c1 - 0.4, "지표", { fontSize: 10 });
    cs.kpis.forEach(([n, l], i) => {
      const kx = x1 + 0.2 + i * ((c1 - 0.4) / 2);
      s.addText(n, { x: kx, y: ky + 0.36, w: (c1 - 0.4) / 2, h: 0.38, fontSize: 20, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "kpi" });
      body(s, kx, ky + 0.74, (c1 - 0.4) / 2, 0.22, l, { fontSize: 9, color: C.text2, valign: "middle" });
    });
    // col2: 과정 (numbered list, auto-flow)
    glass(s, x2, CT, c2, hAll);
    label(s, x2 + 0.2, CT + 0.15, c2 - 0.4, "과정");
    s.addText(
      cs.steps.map((st, i) => ({
        text: st,
        options: { bullet: { type: "number", numberType: "arabicPeriod", indent: 26 }, paraSpaceAfter: 9, breakLine: i < cs.steps.length - 1 },
      })),
      { x: x2 + 0.2, y: CT + 0.52, w: c2 - 0.4, h: hAll - 0.7, fontSize: 14, color: C.text1, margin: 0, valign: "top", align: "left", isTextBox: true, lineSpacingMultiple: 0.95, objectName: "process steps" }
    );
    // col3: 결과 + 배운 점 (one flowing box so spacing stays consistent)
    glass(s, x3, CT, c3, hAll);
    s.addText([
      { text: "결과", options: { fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, paraSpaceAfter: 6, breakLine: true } },
      { text: cs.result, options: { fontSize: 14, color: C.text1, paraSpaceAfter: 18, breakLine: true } },
      { text: "배운 점", options: { fontSize: 11, bold: true, color: C.accent1, charSpacing: 1, paraSpaceAfter: 6, breakLine: true } },
      { text: cs.learn, options: { fontSize: 14, color: C.text1 } },
    ], { x: x3 + 0.2, y: CT + 0.2, w: c3 - 0.4, h: hAll - 0.35, margin: 0, valign: "top", align: "left", isTextBox: true, lineSpacingMultiple: 0.95, objectName: "result and learning" });
  }

  // ===== 10. Career Timeline =====
  pres.addSection({ title: "Career" });
  {
    const s = contentSlide("Career", "// timeline", "Career Timeline");
    const items = [
      ["2026.04", "스톡옵션 세무 체계 정립", "스톡옵션 법률 구조 및 과세 체계 분석 완료. 임직원 보상 설계의 세무 관점 가이드라인 수립."],
      ["2026.04", "ABTC 인증 처리 완료", "APEC 기업인 여행카드(ABTC) 관련 문서 작업 및 인증 프로세스 완료."],
      ["2026.04", "SHA(주주간계약) 분석", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 등 주주간계약의 핵심 조항 분석 및 실무 적용."],
      ["2026.04", "IT 스타트업 런웨이 관리 체계 수립", "현금흐름 기반 런웨이 관리 프레임워크 학습 및 GXC 적용 방안 도출."],
      ["2026 ~", "GXC 경영지원팀 합류", "재무/회계, IR, HR, 정부사업 등 경영지원 전반 업무 담당. 스타트업 운영 기반 구축."],
    ];
    const pitch = (CB - CT) / 5, ch = pitch - 0.1;
    const lineX = 1.72;
    s.addShape(pres.shapes.LINE, { x: lineX, y: CT + 0.2, w: 0, h: CB - CT - 0.4, line: { color: C.accent4, width: 1.5 }, objectName: "timeline" });
    items.forEach(([d, t, desc], i) => {
      const y = CT + i * pitch;
      s.addText(d, { x: M, y, w: 1.1, h: ch, fontSize: 12, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "date" });
      s.addShape(pres.shapes.OVAL, { x: lineX - 0.08, y: y + ch / 2 - 0.08, w: 0.16, h: 0.16, fill: { color: i === 0 ? C.accent1 : C.background2 }, line: { color: C.accent1, width: 1.5 }, objectName: "node" });
      const cx = 1.98, cw = 9.5 - cx;
      glass(s, cx, y, cw, ch);
      s.addText(t, { x: cx + 0.22, y, w: 2.5, h: ch, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "event" });
      body(s, cx + 2.85, y + 0.06, cw - 3.05, ch - 0.12, desc, { fontSize: 14, color: C.text2, valign: "middle", ls: 0.95 });
    });
  }

  // ===== 11. Certifications & Learning =====
  {
    const s = contentSlide("Career", "// certifications", "Certifications & Learning");
    const lw = 4.4, rx = 5.1, rw = 4.4, h = CB - CT;
    glass(s, M, CT, lw, h);
    label(s, 0.75, CT + 0.18, 3, "자격증 (준비중)");
    const certs = [["전산회계 1급", "한국세무사회", "2026.06 목표"], ["전산세무 2급", "한국세무사회", "2026.06 목표"]];
    certs.forEach(([n, org, goal], i) => {
      const y = CT + 0.7 + i * 1.55;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.75, y, w: lw - 0.5, h: 1.3, rectRadius: 0.1, fill: { color: C.accent5 }, line: { color: C.accent4, width: 0.75 }, objectName: "cert row" });
      iconBox(s, 0.95, y + 0.32, 0.66, ic.award, { alt: "certificate icon" });
      s.addText(n, { x: 1.85, y: y + 0.22, w: 2.7, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "cert name" });
      body(s, 1.85, y + 0.62, 2.7, 0.26, org, { fontSize: 12, color: C.text2, valign: "middle" });
      chip(s, 1.85, y + 0.92, goal, { fontSize: 10, bold: true, h: 0.26 });
    });
    glass(s, rx, CT, rw, h);
    label(s, rx + 0.25, CT + 0.18, 3, "최근 학습");
    const learn = [["2026.04.16", "스톡옵션 과세 체계", "IR/투자"], ["2026.04.14", "스톡옵션 법률 구조", "IR/투자"], ["2026.04.10", "주주간계약(SHA) 핵심 조항", "IR/투자"], ["2026.04.08", "IT 스타트업 런웨이 관리", "재무"]];
    const rp = (h - 0.7) / 4;
    learn.forEach(([d, t, cat], i) => {
      const y = CT + 0.62 + i * rp;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx + 0.25, y: y + 0.05, w: rw - 0.5, h: rp - 0.12, rectRadius: 0.08, fill: { color: C.accent5 }, line: { color: C.accent4, width: 0.75 }, objectName: "learning row" });
      iconBox(s, rx + 0.4, y + 0.05 + (rp - 0.12 - 0.44) / 2, 0.44, ic.book, { alt: "book icon" });
      s.addText(d, { x: rx + 1.0, y: y + 0.12, w: 1.5, h: 0.22, fontSize: 10, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "date" });
      body(s, rx + 1.0, y + 0.34, rw - 1.4, 0.3, t, { fontSize: 14, valign: "middle" });
      const cwid = estW(cat, 9) * 1.12 + 0.24;
      chip(s, rx + rw - 0.4 - cwid, y + 0.11, cat, { fontSize: 9, h: 0.22 });
    });
  }

  // ===== 12. Project Gallery =====
  pres.addSection({ title: "Projects" });
  {
    const s = contentSlide("Projects", "// projects", "Project Gallery");
    const projects = [
      ["IR/투자", "SHA 주주간계약 검토", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 분석 및 투자 계약 실무 적용", "투자 계약 핵심 조항 가이드 문서화"],
      ["IR/투자", "스톡옵션 체계 설계", "법률 구조 분석, 과세 체계 정립, 임직원 보상 설계 가이드라인 수립", "스톡옵션 세무/법률 통합 가이드 완성"],
      ["HR/노무", "급여 체계 설계", "TTC(Total Target Compensation) 기반 급여 체계 설계 및 보상 구조 수립", "전사 보상 체계 프레임워크 구축"],
      ["정부사업", "ABTC 인증 처리", "APEC 기업인 여행카드 관련 서류 준비 및 인증 프로세스 관리", "인증 완료 및 프로세스 SOP 문서화"],
      ["재무/회계", "런웨이 관리 체계", "IT 스타트업 현금흐름 기반 런웨이 관리 프레임워크 구축", "월별 런웨이 모니터링 대시보드 구축"],
      ["재무/회계", "월말 결산 프로세스", "매월 재무제표 마감 및 결산 프로세스 체계화, 전표 처리 자동화", "결산 소요 시간 단축 및 프로세스 표준화"],
    ];
    const cw = 4.45, gx = 0.1, gy = 0.08, ch = (CB - CT - 2 * gy) / 3;
    projects.forEach(([cat, t, desc, res], i) => {
      const x = M + (i % 2) * (cw + gx), y = CT + Math.floor(i / 2) * (ch + gy);
      glass(s, x, y, cw, ch);
      const cwid = chip(s, x + 0.2, y + 0.13, cat, { fontSize: 10, h: 0.26 });
      s.addText(t, { x: x + 0.2 + cwid + 0.12, y: y + 0.09, w: cw - cwid - 0.9, h: 0.34, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "project title" });
      s.addText(String(i + 1).padStart(2, "0"), { x: x + cw - 0.65, y: y + 0.12, w: 0.45, h: 0.28, fontSize: 10, bold: true, color: C.text2, align: "right", margin: 0, valign: "middle", isTextBox: true, objectName: "index" });
      body(s, x + 0.2, y + 0.48, cw - 0.4, 0.56, desc, { fontSize: 14, ls: 0.95 });
      s.addText([
        { text: "→ ", options: { color: C.accent1, bold: true } },
        { text: res, options: { color: C.accent6 } },
      ], { x: x + 0.2, y: y + ch - 0.36, w: cw - 0.4, h: 0.26, fontSize: 12, margin: 0, valign: "middle", isTextBox: true, objectName: "outcome" });
    });
  }

  // ===== 13. 마무리 =====
  pres.addSection({ title: "Outro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Outro" });
    glass(s, 1.1, 0.9, 7.8, 3.85, { dots: true, header: "contact", name: "terminal window" });
    s.addText([
      { text: "$ ", options: { color: C.accent2 } },
      { text: "connect", options: { color: C.accent1 } },
    ], { x: 1.5, y: 1.4, w: 4, h: 0.3, fontSize: 13, bold: true, margin: 0, isTextBox: true, objectName: "prompt" });
    s.addText("Let's Connect", { x: 1.5, y: 1.8, w: 6.5, h: 0.9, fontSize: 48, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "closing title" });
    body(s, 1.5, 2.8, 6.8, 0.4, "새로운 기회나 협업에 대해 이야기하고 싶으시다면 연락주세요.", { fontSize: 16 });
    const contacts = [["Email", ic.mail], ["LinkedIn", ic.linkedin]];
    let cx = 1.5;
    contacts.forEach(([t, icon]) => {
      const w = 1.75, h = 0.5, y = 3.45;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y, w, h, rectRadius: 0.1, fill: { color: C.accent3 }, line: { color: C.accent4, width: 0.75 }, objectName: "contact" });
      s.addImage({ data: icon, x: cx + 0.18, y: y + 0.13, w: 0.24, h: 0.24, altText: t + " icon", objectName: t + " icon" });
      s.addText(t, { x: cx + 0.52, y, w: w - 0.6, h, fontSize: 13, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "contact label" });
      cx += w + 0.2;
    });
    body(s, 1.5, 4.25, 6.8, 0.3, "© 2026 양남규", { fontSize: 11, color: C.text2 });
    iconBox(s, 7.45, 1.95, 1.0, ic.terminal, { alt: "terminal icon" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });

function chipFlowHeight(items, maxW, pt, h, gap) {
  let cx = 0, rows = 1;
  for (const t of items) {
    const w = estW(t, pt) * 1.12 + 0.26;
    if (cx + w > maxW + 0.01 && cx > 0) { cx = 0; rows++; }
    cx += w + gap;
  }
  return rows * h + (rows - 1) * gap;
}
