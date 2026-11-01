// 06-korean-formal — "Korean Formal" portfolio deck (양남규) — content v2 (14 slides)
// Run: NODE_PATH=<node_modules with pptxgenjs/sharp/react-icons> node build.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fa = require("react-icons/fa");

const SKILL_DIR =
  process.env.PPTX_SKILL_DIR ||
  "/root/.claude/skills/synced/7d6f9636-edbe-41a2-aab0-0f84fceac3dd_307ca375-45f7-4235-8c53-4eacd5b231ab/pptx";
const { applyTheme } = require(path.join(SKILL_DIR, "scripts", "apply_theme.js"));

const OUT = path.join(__dirname, "deck.pptx");

// ---------- Theme (brief palette) ----------
const P = {
  navy: "1E2761", // 감청
  sky: "CADCFC", // 하늘
  pale: "EEF3FB", // 연하늘
  text: "1A1A1A", // 본문
  muted: "5A6478", // 보조
  red: "C8102E", // 포인트 (최소)
  white: "FFFFFF",
};
const THEME = {
  name: "Korean Formal",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: P.text,
    lt1: P.white,
    dk2: P.navy,
    lt2: P.pale,
    accent1: P.navy,
    accent2: P.sky,
    accent3: P.muted,
    accent4: P.red,
    accent5: P.pale,
    accent6: P.sky,
    hlink: P.navy,
    folHlink: P.muted,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "양남규";
pres.title = "양남규 경영지원 프로필 — Korean Formal";
const C = pres.SchemeColor;
const NAVY = C.text2; // dk2
const TEXT = C.text1; // dk1
const PALE = C.background2; // lt2
const WHITE = C.background1; // lt1
const SKY = C.accent2;
const MUTED = C.accent3;
const RED = C.accent4;

// ---------- Icons (react-icons -> sharp PNG) ----------
const iconCache = {};
async function icon(name, hex, size = 256) {
  const key = `${name}-${hex}`;
  if (iconCache[key]) return iconCache[key];
  const Comp = Fa[name];
  if (!Comp) throw new Error("No icon " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + hex, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}

// ---------- Layouts ----------
const FOOT_Y = 5.27;
pres.defineSlideMaster({
  title: "COVER",
  background: { color: P.navy },
  objects: [],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: P.white },
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.5, y: 0.64, w: 9.0, h: 1.25, fontSize: 36, bold: true, color: NAVY, margin: 0, valign: "top", align: "left" },
        text: "",
      },
    },
    { text: { text: "양남규 · 경영지원 매니저 포트폴리오", options: { x: 0.5, y: FOOT_Y, w: 5, h: 0.25, fontSize: 9, color: MUTED, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: 8.9, y: FOOT_Y, w: 0.6, h: 0.25, fontSize: 9, color: MUTED, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "CLOSING",
  background: { color: P.white },
  objects: [],
});

// ---------- Helpers ----------
let n = 0;
function content(section, labelText, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  n++;
  // Motif: small navy square label at top-left
  const w = Math.max(1.0, labelText.length * 0.085 + 0.35);
  s.addText(labelText, {
    x: 0.5, y: 0.34, w, h: 0.26, fill: { color: NAVY }, color: WHITE, fontSize: 10, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Section label", charSpacing: 1,
  });
  s.addText(title, { placeholder: "title" });
  return s;
}
// section box (구획 박스): pale fill or white with sky border
function box(s, x, y, w, h, opts = {}) {
  s.addShape(pres.ShapeType.rect, {
    x, y, w, h,
    fill: { color: opts.fill || PALE },
    line: opts.line === false ? { color: opts.fill || PALE, width: 0 } : { color: opts.lineColor || SKY, width: 0.75 },
    objectName: opts.name || "Panel",
  });
}
function label(s, text, x, y, w = 0.7, h = 0.26, opts = {}) {
  s.addText(text, {
    x, y, w, h, fill: { color: opts.fill || NAVY }, color: opts.color || WHITE, fontSize: opts.fontSize || 10, bold: true,
    align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Label " + text,
  });
}
function chip(s, text, x, y, w, h = 0.26, opts = {}) {
  s.addText(text, {
    x, y, w, h, fill: { color: opts.fill || WHITE }, line: { color: opts.line || SKY, width: 0.75 },
    color: opts.color || NAVY, fontSize: opts.fontSize || 10, bold: !!opts.bold, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Chip " + text,
  });
}
function txt(s, text, x, y, w, h, opts = {}) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize: 14, color: TEXT, margin: 0, valign: "top", align: "left", isTextBox: true }, opts));
}
function bullets(items, opts = {}) {
  return items.map((t, i) => ({
    text: t,
    options: Object.assign({ bullet: opts.numbered ? { type: "number" } : { code: "25A0" }, breakLine: i < items.length - 1, paraSpaceAfter: opts.space || 4 }, opts.run || {}),
  }));
}

// step flow: boxed steps joined by arrows (공고 분석 → 사업계획서 → …)
function flow(s, steps, x, y, w, h = 0.34, opts = {}) {
  const gap = opts.gap || 0.34;
  const cw = (w - gap * (steps.length - 1)) / steps.length;
  steps.forEach((t, i) => {
    const cx = x + i * (cw + gap);
    s.addText(t, {
      x: cx, y, w: cw, h, fill: { color: opts.fill || PALE }, line: { color: opts.line || SKY, width: 0.75 },
      color: opts.color || NAVY, fontSize: opts.fontSize || 11, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Step " + t,
    });
    if (i < steps.length - 1) {
      s.addText("→", { x: cx + cw, y, w: gap, h, fontSize: 14, bold: true, color: opts.arrow || NAVY, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Arrow" });
    }
  });
}

(async () => {
  const navyIcon = (name) => icon(name, P.white);
  const paleIcon = (name) => icon(name, P.navy);
  // icon tile: navy square with white icon
  async function tile(s, name, x, y, size = 0.5, alt = "아이콘") {
    s.addShape(pres.ShapeType.rect, { x, y, w: size, h: size, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Icon tile" });
    const iw = size * 0.56;
    s.addImage({ data: await navyIcon(name), x: x + (size - iw) / 2, y: y + (size - iw) / 2, w: iw, h: iw, objectName: alt });
  }
  const AREAS = [
    ["재무·회계", "FaCalculator"],
    ["인사·노무", "FaUsers"],
    ["IR·투자", "FaHandshake"],
    ["정부사업", "FaLandmark"],
    ["해외업무", "FaGlobeAmericas"],
    ["총무", "FaBuilding"],
  ];

  // ===== 1. 표지 =====
  pres.addSection({ title: "표지" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
    n++;
    s.addText("PORTFOLIO", { x: 0.7, y: 0.6, w: 1.3, h: 0.3, fill: { color: SKY }, color: NAVY, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Cover label", charSpacing: 2 });
    txt(s, "경영지원 매니저", 0.7, 1.25, 5.4, 0.4, { fontSize: 18, color: SKY, bold: true });
    s.addText("양남규", { x: 0.7, y: 1.65, w: 5.4, h: 1.05, fontSize: 60, bold: true, color: WHITE, margin: 0, valign: "middle", isTextBox: true, objectName: "이름" });
    txt(s, "Duke Yang", 0.7, 2.75, 5.4, 0.42, { fontSize: 20, color: SKY, bold: true });
    txt(
      s,
      "IT 스타트업에서 재무·인사·IR·정부사업·해외업무·총무를 혼자 맡아 운영하며, 그 일을 AI와 지식관리 도구로 시스템화하는 경영지원 매니저입니다.",
      0.7, 3.25, 5.3, 0.95,
      { fontSize: 14, color: WHITE, lineSpacingMultiple: 1.2 }
    );
    // 구획 grid: six areas (3 x 2) on the right
    const gx = 6.35, gy = 1.25, cs = 1.0, gg = 0.1;
    for (let i = 0; i < 6; i++) {
      const cx = gx + (i % 3) * (cs + gg), cy = gy + Math.floor(i / 3) * (cs + gg);
      s.addShape(pres.ShapeType.rect, { x: cx, y: cy, w: cs, h: cs, fill: { color: P.navy }, line: { color: SKY, width: 0.75 }, objectName: "Grid cell" });
      s.addImage({ data: await icon(AREAS[i][1], P.sky), x: cx + 0.32, y: cy + 0.2, w: 0.36, h: 0.36, objectName: AREAS[i][0] + " 아이콘" });
      s.addText(AREAS[i][0], { x: cx, y: cy + 0.62, w: cs, h: 0.3, fontSize: 9.5, color: SKY, align: "center", margin: 0, isTextBox: true, objectName: AREAS[i][0] });
    }
    // bottom info panel
    s.addShape(pres.ShapeType.rect, { x: 0.7, y: 4.45, w: 8.6, h: 0.6, fill: { color: P.navy }, line: { color: SKY, width: 0.75 }, objectName: "Info panel" });
    s.addText(
      [
        { text: "재무 · 인사 · IR · 정부사업 · 해외업무 · 총무", options: { color: WHITE, bold: true } },
        { text: "     /     기준일 2026-10-27", options: { color: SKY } },
      ],
      { x: 0.9, y: 4.45, w: 8.2, h: 0.6, fontSize: 13, margin: 0, valign: "middle", isTextBox: true, objectName: "보조 정보" }
    );
  }

  // ===== 2. 일하는 원칙 =====
  pres.addSection({ title: "원칙" });
  {
    const s = content("원칙", "01  PRINCIPLES", "일하는 원칙");
    const items = [
      ["시스템", "사람의 기억이 아닌 시스템에 남긴다.", "FaDatabase"],
      ["원문 확인", "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보낸다.", "FaSearch"],
      ["스킬 · SOP", "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다.", "FaClipboardList"],
    ];
    const pw = 2.87, gap = 0.195, y0 = 1.5, ph = 3.6;
    for (let i = 0; i < 3; i++) {
      const x = 0.5 + i * (pw + gap);
      box(s, x, y0, pw, ph, { fill: i === 1 ? PALE : WHITE });
      txt(s, String(i + 1).padStart(2, "0"), x + 0.25, y0 + 0.2, 1.2, 0.5, { fontSize: 28, bold: true, color: SKY });
      await tile(s, items[i][2], x + pw - 0.8, y0 + 0.25, 0.55, items[i][0] + " 아이콘");
      txt(s, items[i][0], x + 0.25, y0 + 1.05, pw - 0.5, 0.6, { fontSize: 26, bold: true, color: NAVY, valign: "middle" });
      s.addShape(pres.ShapeType.rect, { x: x + 0.25, y: y0 + 1.85, w: 0.14, h: 0.14, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Marker" });
      txt(s, items[i][1], x + 0.25, y0 + 2.1, pw - 0.5, 1.3, { fontSize: 15, lineSpacingMultiple: 1.25 });
    }
  }

  // ===== 3. 소개 =====
  pres.addSection({ title: "소개" });
  {
    const s = content("소개", "02  ABOUT", "소개");
    // left profile tile
    s.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 2.3, h: 3.6, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Profile tile" });
    s.addShape(pres.ShapeType.rect, { x: 0.75, y: 1.8, w: 0.7, h: 0.7, fill: { color: P.navy }, line: { color: SKY, width: 0.75 }, objectName: "Profile icon cell" });
    s.addImage({ data: await icon("FaUserTie", P.sky), x: 0.9, y: 1.95, w: 0.4, h: 0.4, objectName: "프로필 아이콘" });
    txt(s, "양남규", 0.75, 2.7, 1.9, 0.5, { fontSize: 24, bold: true, color: WHITE });
    txt(s, "Duke Yang", 0.75, 3.2, 1.9, 0.3, { fontSize: 12, color: SKY, bold: true });
    txt(s, "경영지원 매니저", 0.75, 3.5, 1.9, 0.3, { fontSize: 13, color: WHITE });
    txt(s, "국제통상학 전공", 0.75, 4.35, 1.9, 0.3, { fontSize: 11, color: SKY });
    txt(s, "IT 스타트업 경영지원팀", 0.75, 4.62, 1.9, 0.3, { fontSize: 11, color: SKY });
    // right: four paragraphs in a 구획 box
    box(s, 3.0, 1.5, 6.5, 3.6);
    const paras = [
      "IT 스타트업 경영지원팀 매니저로 일하며 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다.",
      "국제통상학을 전공했고, 미국 현지 유통기업에서 근무한 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다.",
      "여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해 옵시디언 기반 업무 운영체계와 AI 자동화 워크플로를 직접 설계해 쓰고 있습니다.",
      "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다.",
    ];
    const ys = [1.72, 2.47, 3.1, 4.05];
    const hs = [0.7, 0.58, 0.9, 0.6];
    paras.forEach((p, i) => {
      s.addShape(pres.ShapeType.rect, { x: 3.25, y: ys[i] + 0.07, w: 0.13, h: 0.13, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Marker" });
      txt(s, p, 3.55, ys[i], 5.7, hs[i], { fontSize: 14, lineSpacingMultiple: 1.2 });
    });
  }

  // ===== 4. 담당 영역 =====
  pres.addSection({ title: "담당 영역" });
  {
    const s = content("담당 영역", "03  AREAS", "담당 영역");
    txt(s, "여섯 영역을 한 사람이 맡아 운영합니다", 0.5, 1.32, 8.5, 0.3, { fontSize: 14, color: MUTED });
    const cards = [
      ["재무·회계", "자금 운용, 세무 전 주기, 결산·외부감사"],
      ["인사·노무", "급여·4대보험, 채용, 해외 원격 인력 계약"],
      ["IR·투자", "주총·이사회, 등기, 스톡옵션, 투자 계약"],
      ["정부사업", "공고 분석부터 정산까지 전 주기"],
      ["해외업무", "해외 자회사 증빙, 해외 행사·출장"],
      ["총무", "법인 행정, 기업 인증, 사무실·자산, 대외행사"],
    ];
    const cw = 2.87, ch = 1.5, gx = 0.195, gy = 0.18, x0 = 0.5, y0 = 1.78;
    for (let i = 0; i < 6; i++) {
      const x = x0 + (i % 3) * (cw + gx), y = y0 + Math.floor(i / 3) * (ch + gy);
      box(s, x, y, cw, ch, { fill: i % 2 === 0 ? PALE : WHITE });
      await tile(s, AREAS[i][1], x + 0.2, y + 0.2, 0.5, cards[i][0] + " 아이콘");
      txt(s, String(i + 1).padStart(2, "0"), x + cw - 0.65, y + 0.18, 0.5, 0.3, { fontSize: 13, bold: true, color: SKY, align: "right" });
      txt(s, cards[i][0], x + 0.85, y + 0.2, cw - 1.1, 0.5, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
      txt(s, cards[i][1], x + 0.2, y + 0.85, cw - 0.4, 0.6, { fontSize: 13, color: TEXT, lineSpacingMultiple: 1.15 });
    }
  }

  // ===== 5. 담당 영역 상세 1 =====
  {
    const s = content("담당 영역", "03  AREAS · 1", "담당 영역 상세 1");
    const cols = [
      ["재무·회계", "FaCalculator", [
        "자금 운용과 송금·결제를 관리하고, 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응합니다.",
        "결산과 외부감사 대응, 별도·연결 재무제표 관리를 맡고 외부 회계법인과의 커뮤니케이션을 담당합니다.",
      ]],
      ["인사·노무", "FaUsers", [
        "급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다.",
        "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다.",
      ]],
      ["IR·투자", "FaHandshake", [
        "정기주주총회·이사회 운영, 정관 개정, 등기, 주주명부와 스톡옵션 관리를 맡고 있습니다.",
        "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다.",
      ]],
    ];
    const cw = 2.87, gap = 0.195, y0 = 1.5, ch = 3.6;
    for (let i = 0; i < 3; i++) {
      const x = 0.5 + i * (cw + gap);
      box(s, x, y0, cw, ch, { fill: i === 1 ? PALE : WHITE });
      await tile(s, cols[i][1], x + 0.2, y0 + 0.2, 0.5, cols[i][0] + " 아이콘");
      txt(s, cols[i][0], x + 0.85, y0 + 0.2, cw - 1.0, 0.5, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
      s.addText(
        cols[i][2].map((p, j) => ({ text: p, options: { breakLine: j === 0, paraSpaceAfter: 10 } })),
        { x: x + 0.2, y: y0 + 0.9, w: cw - 0.4, h: ch - 1.05, fontSize: 13, color: TEXT, lineSpacingMultiple: 1.2, margin: 0, valign: "top", isTextBox: true, objectName: cols[i][0] + " 상세" }
      );
    }
  }

  // ===== 6. 담당 영역 상세 2 =====
  {
    const s = content("담당 영역", "03  AREAS · 2", "담당 영역 상세 2");
    // 정부사업: full-width box with 5-step flow
    box(s, 0.5, 1.5, 9.0, 1.62, { fill: PALE });
    await tile(s, "FaLandmark", 0.7, 1.68, 0.5, "정부사업 아이콘");
    txt(s, "정부사업", 1.35, 1.68, 3, 0.5, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
    txt(s, "공고 분석 → 사업계획서 작성 → 신청 → 협약 → 정산까지 전 주기를 실무로 수행합니다.", 4.0, 1.68, 5.3, 0.5, { fontSize: 13, valign: "middle", lineSpacingMultiple: 1.15 });
    flow(s, ["공고 분석", "사업계획서", "신청", "협약", "정산"], 0.7, 2.3, 8.6, 0.34, { fill: WHITE });
    txt(s, "입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다.", 0.7, 2.74, 8.6, 0.3, { fontSize: 13, valign: "middle" });
    // 해외업무 / 총무
    const bw = 4.4, by = 3.3, bh = 1.8;
    box(s, 0.5, by, bw, bh, { fill: WHITE });
    await tile(s, "FaGlobeAmericas", 0.7, by + 0.2, 0.5, "해외업무 아이콘");
    txt(s, "해외업무", 1.35, by + 0.2, 3, 0.5, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
    txt(s, "해외 자회사의 인보이스와 용역계약 증빙을 관리하고, 해외 사업 행사 참가와 해외 출장 서류 실무를 지원합니다.", 0.7, by + 0.8, bw - 0.4, 0.95, { fontSize: 13, lineSpacingMultiple: 1.2 });
    box(s, 5.1, by, bw, bh, { fill: WHITE });
    await tile(s, "FaBuilding", 5.3, by + 0.2, 0.5, "총무 아이콘");
    txt(s, "총무", 5.95, by + 0.2, 3, 0.5, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
    s.addText(
      [
        { text: "본점 이전 같은 법인 행정 절차와 기업 인증 관리, 사무실·자산 관리를 담당합니다.", options: { breakLine: true, paraSpaceAfter: 4 } },
        { text: "대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다.", options: {} },
      ],
      { x: 5.3, y: by + 0.8, w: bw - 0.4, h: 0.95, fontSize: 13, color: TEXT, lineSpacingMultiple: 1.2, margin: 0, valign: "top", isTextBox: true, objectName: "총무 상세" }
    );
  }

  // ===== 7–11. 사례 =====
  pres.addSection({ title: "사례" });
  const cases = [
    {
      num: "01", title: "1인 경영지원을 \"인계 가능한 시스템\"으로 바꾸기", icon: "FaSitemap",
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 커스텀 AI 스킬로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다.",
    },
    {
      num: "02", title: "단일 통화 자금 시트를 의사결정 도구로 재설계", icon: "FaTable",
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리하고 있어, 경영진의 자금 판단 질문에 바로 답하기 어려웠습니다.",
      did: "대시보드·통화별 입출금·정기 비용·미지급 항목으로 시트를 나누고, 다중 통화 환산과 \"부족액 식별 → 송금액 계산 → 송금 후 잔액\" 시뮬레이션을 자동화했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.",
      flow: ["부족액 식별", "송금액 계산", "송금 후 잔액"],
      result: "경영진의 자금 의사결정 질문에 한 화면에서 바로 답할 수 있는 상태가 됐습니다.",
    },
    {
      num: "03", title: "법률 1차 검토를 법령 원문 기반으로 내재화", icon: "FaBalanceScale",
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했습니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 실무 상식만으로는 놓칠 수 있는 별도 조항을 미리 잡게 됐습니다.",
    },
    {
      num: "04", title: "외부로 나가기 전 \"받은 산출물\"을 역검증하는 습관", icon: "FaCheckDouble",
      bg: "회계법인·기관·플랫폼에서 받은 자료나 내부 문서가 확인 없이 외부로 나가면 세금·등기·평가에 바로 영향을 줍니다.",
      did: "회계법인 산출물, 기관·플랫폼 자료, 계약서와 내부 문서를 원장·법령 원문과 대조한 뒤 제출하는 절차를 분기 내내 운영했습니다.",
      result: "기간 오류, 견적 오기, 정관과 계약의 충돌 같은 문제를 외부 제출 전에 찾아냈습니다.",
    },
    {
      num: "05", title: "예산 대비 실적 관리 워크북 구축", icon: "FaChartBar",
      bg: "월별 지출 추적이 여러 엑셀 파일에 흩어져 있어 예산을 통제하기 어려웠습니다.",
      did: "실제 계정원장을 기준으로 계정과목을 정리해 카테고리로 묶었습니다. 월별 예산 수립 → 실적 입력 → 차이 분석 흐름을 자동화했고, 이후 중앙 가정 시트와 연동되도록 확장했습니다.",
      flow: ["예산 수립", "실적 입력", "차이 분석"],
      result: "예산 대비 실적을 한 체계에서 관리하는 관리회계 기반을 만들었습니다.",
    },
  ];
  for (const cs of cases) {
    const s = content("사례", `04  CASE ${cs.num}`, cs.title);
    const BODY = 14;
    // Row A: 배경 — body starts below a one- or two-line title (title position itself is fixed)
    const twoLine = cs.title.replace(/\s/g, "").length >= 16;
    const yA = twoLine ? 2.0 : 1.7, hA = 0.78;
    box(s, 0.5, yA, 9.0, hA);
    label(s, "배경", 0.62, yA + 0.12, 0.6);
    txt(s, cs.bg, 1.35, yA + 0.1, 7.95, hA - 0.15, { fontSize: BODY, lineSpacingMultiple: 1.15 });
    // Row B: 한 일 (+ optional flow)
    const yB = yA + hA + 0.1, hC = 0.62, hB = 5.12 - hC - 0.1 - (2.0 + hA + 0.1); // same box heights on every case slide
    box(s, 0.5, yB, 9.0, hB, { fill: WHITE });
    label(s, "한 일", 0.62, yB + 0.12, 0.6);
    const didW = cs.flow ? 7.95 : 7.3;
    txt(s, cs.did, 1.35, yB + 0.12, didW, cs.flow ? hB - 0.75 : hB - 0.2, { fontSize: BODY, lineSpacingMultiple: 1.15 });
    if (cs.flow) {
      flow(s, cs.flow, 1.35, yB + hB - 0.46, 4.6, 0.34);
      s.addImage({ data: await paleIcon(cs.icon), x: 8.85, y: yB + hB - 0.57, w: 0.45, h: 0.45, objectName: "사례 아이콘" });
    } else {
      s.addImage({ data: await paleIcon(cs.icon), x: 8.85, y: yB + 0.12, w: 0.45, h: 0.45, objectName: "사례 아이콘" });
    }
    // Row C: 결과 — navy tile so it stands out
    const yC = yB + hB + 0.1;
    s.addShape(pres.ShapeType.rect, { x: 0.5, y: yC, w: 9.0, h: hC, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "결과 tile" });
    label(s, "결과", 0.62, yC + 0.12, 0.6, 0.26, { fill: SKY, color: NAVY });
    txt(s, cs.result, 1.35, yC + 0.1, 7.95, hC - 0.15, { fontSize: BODY, color: WHITE, bold: true, lineSpacingMultiple: 1.15, valign: "middle" });
  }

  // ===== 12. 스킬·도구 =====
  pres.addSection({ title: "스킬" });
  {
    const s = content("스킬", "05  SKILLS", "스킬·도구");
    const H = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "left", fontSize: 11 } });
    const K = (t) => ({ text: t, options: { bold: true, color: NAVY, fontSize: 13 } });
    const V = (t) => ({ text: t, options: { color: TEXT, fontSize: 13 } });
    const rows = [
      ["엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·자금 시뮬레이션 설계"],
      ["옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인"],
      ["AI 업무 자동화", "Claude 커스텀 스킬 설계·운영, 예약 작업 자동화, MCP 연동(법령 원문 조회, 옵시디언 연동)"],
      ["협업 도구", "Gmail, Google Drive, Slack"],
      ["ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)"],
      ["문서 표준화", "공문 서식, docx·pptx 폰트 표준, 인보이스·인사·급여 템플릿, SOP"],
      ["실무 지식", "결산·세무 신고 대응, 주총·이사회·등기, 정부사업 전 주기"],
      ["언어", "비즈니스 영어(미국 현지 근무 경험)"],
    ];
    const tbl = [[H("영역"), H("내용")]].concat(
      rows.map((r, i) => {
        const fill = i % 2 === 0 ? { color: PALE } : { color: WHITE };
        const k = K(r[0]); k.options.fill = fill;
        const v = V(r[1]); v.options.fill = fill;
        return [k, v];
      })
    );
    s.addTable(tbl, {
      x: 0.5, y: 1.5, w: 9.0, colW: [1.9, 7.1], rowH: [0.3].concat(rows.map(() => 0.38)),
      border: { type: "solid", color: P.sky, pt: 0.75 }, valign: "middle", margin: [0.03, 0.12, 0.03, 0.12],
      objectName: "스킬·도구 표",
    });
  }

  // ===== 13. 타임라인 =====
  pres.addSection({ title: "경력" });
  {
    const s = content("경력", "06  TIMELINE", "타임라인");
    const items = [
      ["2022 – 현재", "IT 스타트업", "경영지원팀 매니저"],
      ["2019 – 2022", "미국 소재 유통기업", "재고관리 인턴 → 구매(Purchasing) 담당"],
      ["2018 – 2019", "미국 비즈니스 취업연수 과정 수료", ""],
      ["2017", "중국 교환학생", ""],
      ["2013 – 2019", "학부", "국제통상학 전공 (중어중문학 복수전공)"],
    ];
    const rh = 0.66, y0 = 1.5;
    s.addShape(pres.ShapeType.rect, { x: 0.5, y: y0, w: 9.0, h: 0.3, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Table head" });
    txt(s, "시기", 0.65, y0, 1.2, 0.3, { fontSize: 10, bold: true, color: WHITE, valign: "middle" });
    txt(s, "소속 · 활동", 2.45, y0, 3, 0.3, { fontSize: 10, bold: true, color: WHITE, valign: "middle" });
    txt(s, "역할", 5.7, y0, 3, 0.3, { fontSize: 10, bold: true, color: WHITE, valign: "middle" });
    txt(s, "▲ 최근  ·  ▼ 과거", 8.0, y0, 1.4, 0.3, { fontSize: 9, color: SKY, valign: "middle", align: "right" });
    s.addShape(pres.ShapeType.line, { x: 2.05, y: y0 + 0.3, w: 0, h: rh * items.length, line: { color: SKY, width: 1 }, objectName: "Timeline line" });
    items.forEach((it, i) => {
      const y = y0 + 0.3 + i * rh;
      box(s, 0.5, y, 9.0, rh, { fill: i % 2 === 0 ? PALE : WHITE, line: false });
      txt(s, it[0], 0.65, y, 1.3, rh, { fontSize: 13, bold: true, color: NAVY, valign: "middle" });
      s.addShape(pres.ShapeType.rect, { x: 1.96, y: y + rh / 2 - 0.09, w: 0.18, h: 0.18, fill: { color: i === 0 ? RED : NAVY }, line: { color: WHITE, width: 1.5 }, objectName: "Timeline marker" });
      txt(s, it[1], 2.45, y, 3.1, rh, { fontSize: 14, bold: true, color: TEXT, valign: "middle" });
      if (it[2]) txt(s, it[2], 5.7, y, 3.7, rh, { fontSize: 13, color: MUTED, valign: "middle" });
    });
  }

  // ===== 14. 연락 =====
  pres.addSection({ title: "마무리" });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "마무리" });
    n++;
    s.addText("07  CONTACT", { x: 0.5, y: 0.34, w: 1.2, h: 0.26, fill: { color: NAVY }, color: WHITE, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Section label", charSpacing: 1 });
    s.addText("연락", { x: 0.5, y: 1.4, w: 6, h: 1.0, fontSize: 48, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: "제목" });
    txt(s, "양남규 · Duke Yang · 경영지원 매니저", 0.5, 2.45, 6, 0.4, { fontSize: 16, color: TEXT });
    box(s, 0.5, 3.15, 4.6, 0.7);
    await tile(s, "FaEnvelope", 0.62, 3.27, 0.46, "이메일 아이콘");
    txt(s, "namgyu1412@naver.com", 1.22, 3.15, 3.8, 0.7, { fontSize: 16, bold: true, color: NAVY, valign: "middle" });
    // right decorative grid (motif echo): six areas
    const gx = 6.3, gy = 1.4, cs = 1.0, gg = 0.1;
    AREAS.forEach((a, i) => {
      const cx = gx + (i % 3) * (cs + gg), cy = gy + Math.floor(i / 3) * (cs + gg);
      const dark = i === 5;
      s.addShape(pres.ShapeType.rect, { x: cx, y: cy, w: cs, h: cs, fill: { color: dark ? NAVY : PALE }, line: { color: SKY, width: 0.75 }, objectName: "Grid cell" });
      txt(s, a[0], cx, cy, cs, cs, { fontSize: 12, bold: true, color: dark ? WHITE : NAVY, align: "center", valign: "middle" });
    });
    txt(s, "© 2026 Duke Yang", 0.5, 4.9, 4, 0.3, { fontSize: 11, color: MUTED });
    txt(s, "양남규 · 경영지원 매니저 포트폴리오", 5.5, 4.9, 4.0, 0.3, { fontSize: 9, color: MUTED, align: "right" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT, "slides:", n);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
