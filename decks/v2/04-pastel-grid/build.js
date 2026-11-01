// Pastel Grid — 양남규 경영지원 프로필 덱 v2 (14장)
// Run: NODE_PATH=<node_modules> node build.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fi = require("react-icons/fi");

const SKILL_DIR =
  "/root/.claude/skills/synced/7d6f9636-edbe-41a2-aab0-0f84fceac3dd_307ca375-45f7-4235-8c53-4eacd5b231ab/pptx";
const { applyTheme } = require(path.join(SKILL_DIR, "scripts/apply_theme.js"));

const OUT = path.join(__dirname, "deck.pptx");

// ---------- THEME ----------
const THEME = {
  name: "Pastel Grid",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "2B3A42", // 진회색 (titles)
    lt1: "FFFFFF", // 바탕
    dk2: "50808E", // 슬레이트 (body)
    lt2: "F3F6F4", // 아주 연한 세이지 틴트
    accent1: "84B59F", // 세이지
    accent2: "B8B5E1", // 라벤더
    accent3: "F5C6AA", // 피치
    accent4: "D9E8DF", // 세이지 틴트 (타일)
    accent5: "E7E6F4", // 라벤더 틴트 (타일)
    accent6: "FBE6D7", // 피치 틴트 (타일)
    hlink: "50808E",
    folHlink: "50808E",
  },
};
const H = THEME.colors;

// ---------- Icons (react-icons -> PNG via sharp) ----------
async function iconPng(IconComp, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color: "#" + hex, size: size, strokeWidth: 2 })
  );
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// rough text width (inches) for chip sizing — Korean ~0.95em, latin ~0.55em
function estW(t, fs) {
  let w = 0;
  for (const ch of t) {
    if (ch === " ") w += 0.3;
    else if (/[ᄀ-ᇿ㄰-㆏가-힯一-鿿　-〿＀-￯]/.test(ch)) w += 0.98;
    else w += 0.58;
  }
  return (w * fs) / 72;
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.author = "양남규";
  pres.title = "양남규 경영지원 프로필 — Pastel Grid";
  const C = pres.SchemeColor;

  // Pre-render icons (dark gray on pastel badges)
  const names = [
    "FiUser", "FiDollarSign", "FiTrendingUp", "FiUsers", "FiShield", "FiBriefcase",
    "FiSearch", "FiGrid", "FiZap", "FiClock", "FiCalendar", "FiFileText", "FiBookOpen",
    "FiMail", "FiCheckCircle", "FiLayers", "FiTarget", "FiFlag", "FiGlobe", "FiHome",
    "FiDatabase", "FiRepeat", "FiMapPin", "FiMessageCircle", "FiServer", "FiEdit3",
    "FiCpu", "FiTool", "FiClipboard", "FiBook", "FiArchive", "FiCompass", "FiSliders",
    "FiBarChart2", "FiFolder", "FiAward",
  ];
  const ICON = {};
  for (const n of names) {
    ICON[n] = await iconPng(Fi[n], H.dk1);
  }

  // ---------- Layouts ----------
  // Title box is two lines tall (case titles wrap); valign top keeps the first
  // line at the same height on every content slide.
  const TITLE_Y = 0.3, TITLE_H = 1.2;
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: H.accent1 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: H.lt1 },
    objects: [
      {
        placeholder: {
          options: {
            name: "title", type: "title", x: 0.5, y: TITLE_Y, w: 7.6, h: TITLE_H,
            fontSize: 36, bold: true, color: C.text1, align: "left", valign: "top", margin: 0,
            lineSpacingMultiple: 1.05,
          },
          text: "",
        },
      },
      {
        text: {
          text: "양남규 · 경영지원 매니저",
          options: { x: 0.5, y: 5.28, w: 5, h: 0.25, fontSize: 9, color: C.text2, margin: 0, valign: "middle" },
        },
      },
    ],
    slideNumber: { x: 8.9, y: 5.28, w: 0.6, h: 0.25, fontSize: 9, color: C.text2, align: "right", margin: 0 },
  });
  pres.defineSlideMaster({
    title: "CLOSING",
    background: { color: H.accent1 },
    objects: [],
  });

  // ---------- Helpers ----------
  let objN = 0;
  const nm = (s) => `${s} ${++objN}`;

  function tile(slide, x, y, w, h, color, opts = {}) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, fill: { color }, line: { color, width: 0 }, rectRadius: opts.r ?? 0.14,
      objectName: nm("Tile"), ...(opts.shadow ? { shadow: { type: "outer", blur: 6, offset: 1, angle: 90, color: H.dk1, opacity: 0.08 } } : {}),
    });
  }
  function badge(slide, x, y, d, color, icon, iconData) {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color, width: 0 }, objectName: nm("Badge") });
    const pad = d * 0.26;
    slide.addImage({ data: iconData || ICON[icon], x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: nm("Icon"), altText: icon });
  }
  function text(slide, t, o) {
    slide.addText(t, { isTextBox: true, margin: 0, fontFace: undefined, objectName: nm("Text"), ...o });
  }
  function chip(slide, t, x, y, w, h, fill, color = C.text1, fs = 10) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: h / 2, objectName: nm("Chip") });
    text(slide, t, { x, y, w, h, fontSize: fs, bold: true, color, align: "center", valign: "middle" });
  }
  function label(slide, t, x, y, w) {
    text(slide, t, { x, y, w, h: 0.24, fontSize: 11, bold: true, color: C.text2, charSpacing: 1 });
  }
  // horizontal step flow: chips joined by arrows
  function flow(slide, steps, x, y, h, fill, fs = 11, arrowColor = C.text2) {
    let cx = x;
    const aw = 0.3;
    steps.forEach((st, i) => {
      const w = estW(st, fs) + 0.32;
      chip(slide, st, cx, y, w, h, fill, C.text1, fs);
      cx += w;
      if (i < steps.length - 1) {
        text(slide, "→", { x: cx, y, w: aw, h, fontSize: fs + 2, bold: true, color: arrowColor, align: "center", valign: "middle" });
        cx += aw;
      }
    });
    return cx - x;
  }
  // vertical numbered steps
  function vsteps(slide, steps, x, y, w, pitch, circleFill, fs = 13) {
    steps.forEach((t, i) => {
      const yy = y + i * pitch;
      slide.addShape(pres.shapes.OVAL, { x, y: yy, w: 0.34, h: 0.34, fill: { color: circleFill }, line: { color: circleFill, width: 0 }, objectName: nm("Step") });
      text(slide, String(i + 1), { x, y: yy, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: C.text1, align: "center", valign: "middle" });
      text(slide, t, { x: x + 0.45, y: yy - 0.03, w: w - 0.45, h: 0.4, fontSize: fs, color: C.text1, valign: "middle" });
      if (i < steps.length - 1) {
        slide.addShape(pres.shapes.LINE, { x: x + 0.17, y: yy + 0.34, w: 0, h: pitch - 0.34, line: { color: circleFill, width: 1.5 }, objectName: nm("Connector") });
      }
    });
  }

  const PASTEL = [C.accent1, C.accent2, C.accent3];
  const TINT = [C.accent4, C.accent5, C.accent6];

  // =============== 1. 표지 ===============
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    // left big white tile
    tile(s, 0.5, 0.55, 5.35, 4.5, C.background1, { r: 0.2 });
    text(s, "경영지원 매니저", { x: 0.9, y: 0.9, w: 4.6, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    text(s, "양남규", { x: 0.9, y: 1.3, w: 4.6, h: 1.0, fontSize: 54, bold: true, color: C.text1 });
    text(s, "Duke Yang", { x: 0.9, y: 2.3, w: 4.6, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
    text(
      s,
      "IT 스타트업에서 재무·인사·IR·정부사업·\n해외업무·총무를 혼자 맡아 운영하며,\n그 일을 AI와 지식관리 도구로 시스템화하는\n경영지원 매니저입니다.",
      { x: 0.9, y: 2.85, w: 4.55, h: 1.3, fontSize: 14, color: C.text1, valign: "top", lineSpacingMultiple: 1.3 }
    );
    chip(s, "기준일 2026-10-27", 0.9, 4.3, 1.75, 0.36, TINT[1], C.text1, 10);
    // right 2x3 white tiles with badges — six areas
    const grid = [
      ["FiDollarSign", "재무·회계"], ["FiUsers", "인사·노무"],
      ["FiTrendingUp", "IR·투자"], ["FiFileText", "정부사업"],
      ["FiGlobe", "해외업무"], ["FiHome", "총무"],
    ];
    const gx = 6.15, gy = 0.55, tw = 1.6, th = 1.4, gap = 0.15;
    grid.forEach(([ic, lab], i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      tile(s, x, y, tw, th, C.background1, { r: 0.2 });
      badge(s, x + (tw - 0.62) / 2, y + 0.22, 0.62, TINT[i % 3], ic);
      text(s, lab, { x: x + 0.08, y: y + 0.92, w: tw - 0.16, h: 0.32, fontSize: 12, bold: true, color: C.text1, align: "center", valign: "middle" });
    });
    s.addNotes("표지 — 양남규, 경영지원 매니저 (기준일 2026-10-27)");
  }

  // =============== 2. 일하는 원칙 ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intro" });
    s.addText("일하는 원칙", { placeholder: "title" });
    const items = [
      ["FiDatabase", "시스템", "사람의 기억이 아닌 시스템에 남긴다."],
      ["FiBookOpen", "원문 확인", "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보낸다."],
      ["FiRepeat", "스킬 · SOP", "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다."],
    ];
    const gx = 0.5, gy = 1.1, tw = 2.87, th = 4.05, gap = 0.195;
    items.forEach(([ic, kw, sent], i) => {
      const x = gx + i * (tw + gap), y = gy;
      tile(s, x, y, tw, th, TINT[i]);
      text(s, `0${i + 1}`, { x: x + 0.3, y: y + 0.25, w: 1.5, h: 0.6, fontSize: 30, bold: true, color: C.text2 });
      badge(s, x + tw - 0.3 - 0.7, y + 0.25, 0.7, PASTEL[i], ic);
      text(s, kw, { x: x + 0.3, y: y + 1.65, w: tw - 0.6, h: 0.6, fontSize: 26, bold: true, color: C.text1, valign: "middle" });
      text(s, sent, { x: x + 0.3, y: y + 2.4, w: tw - 0.6, h: 1.5, fontSize: 16, color: C.text1, valign: "top", lineSpacingMultiple: 1.3 });
    });
    s.addNotes("일하는 원칙 3개 — 시스템, 원문 확인, 스킬·SOP");
  }

  // =============== 3. 소개 ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intro" });
    s.addText("소개", { placeholder: "title" });
    const paras = [
      ["FiUser", "IT 스타트업 경영지원팀 매니저로 일하며 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다."],
      ["FiGlobe", "국제통상학을 전공했고, 미국 현지 유통기업에서 근무한 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다."],
      ["FiLayers", "여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해 옵시디언 기반 업무 운영체계와 AI 자동화 워크플로를 직접 설계해 쓰고 있습니다."],
      ["FiCheckCircle", "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다."],
    ];
    const gx = 0.5, gy = 1.1, tw = 4.4, th = 1.95, gap = 0.15;
    paras.forEach(([ic, p], i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      tile(s, x, y, tw, th, TINT[i % 3]);
      badge(s, x + 0.25, y + 0.22, 0.46, PASTEL[i % 3], ic);
      text(s, p, { x: x + 0.25, y: y + 0.8, w: tw - 0.5, h: th - 0.95, fontSize: 14, color: C.text1, valign: "top", lineSpacingMultiple: 1.15 });
    });
    s.addNotes("소개 4문단");
  }

  // =============== 4. 담당 영역 ===============
  pres.addSection({ title: "Work" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Work" });
    s.addText("담당 영역", { placeholder: "title" });
    text(s, "여섯 영역을 한 사람이 맡아 운영합니다", { x: 4.5, y: TITLE_Y + 0.1, w: 5.0, h: 0.45, fontSize: 14, color: C.text2, align: "right", valign: "middle" });
    const items = [
      ["FiDollarSign", "재무·회계", "자금 운용, 세무 전 주기,\n결산·외부감사"],
      ["FiUsers", "인사·노무", "급여·4대보험, 채용,\n해외 원격 인력 계약"],
      ["FiTrendingUp", "IR·투자", "주총·이사회, 등기,\n스톡옵션, 투자 계약"],
      ["FiFileText", "정부사업", "공고 분석부터 정산까지 전 주기"],
      ["FiGlobe", "해외업무", "해외 자회사 증빙,\n해외 행사·출장"],
      ["FiHome", "총무", "법인 행정, 기업 인증,\n사무실·자산, 대외행사"],
    ];
    const gx = 0.5, gy = 1.1, tw = 2.87, th = 1.95, gap = 0.195;
    items.forEach(([ic, h, d], i) => {
      const x = gx + (i % 3) * (tw + gap), y = gy + Math.floor(i / 3) * (th + gap);
      tile(s, x, y, tw, th, TINT[i % 3]);
      badge(s, x + 0.25, y + 0.25, 0.62, PASTEL[i % 3], ic);
      text(s, h, { x: x + 0.25, y: y + 1.0, w: tw - 0.5, h: 0.35, fontSize: 17, bold: true, color: C.text1, valign: "middle" });
      text(s, d, { x: x + 0.25, y: y + 1.35, w: tw - 0.5, h: 0.55, fontSize: 13, color: C.text2, valign: "top", lineSpacingMultiple: 1.15 });
    });
    s.addNotes("담당 영역 6개 — 여섯 영역을 한 사람이 맡아 운영");
  }

  // =============== 5. 담당 영역 상세 1 ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Work" });
    s.addText("담당 영역 상세 1", { placeholder: "title" });
    text(s, "재무·회계 / 인사·노무 / IR·투자", { x: 4.5, y: TITLE_Y + 0.1, w: 5.0, h: 0.45, fontSize: 14, color: C.text2, align: "right", valign: "middle" });
    const rows = [
      ["FiDollarSign", "재무·회계", [
        "자금 운용과 송금·결제를 관리하고, 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응합니다.",
        "결산과 외부감사 대응, 별도·연결 재무제표 관리를 맡고 외부 회계법인과의 커뮤니케이션을 담당합니다.",
      ]],
      ["FiUsers", "인사·노무", [
        "급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다.",
        "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다.",
      ]],
      ["FiTrendingUp", "IR·투자", [
        "정기주주총회·이사회 운영, 정관 개정, 등기, 주주명부와 스톡옵션 관리를 맡고 있습니다.",
        "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다.",
      ]],
    ];
    const rh = 1.25, gap = 0.15, y0 = 1.1, x0 = 0.5, w = 9.0;
    rows.forEach(([ic, h, ps], i) => {
      const y = y0 + i * (rh + gap);
      tile(s, x0, y, w, rh, TINT[i]);
      badge(s, x0 + 0.3, y + 0.22, 0.5, PASTEL[i], ic);
      text(s, h, { x: x0 + 0.3, y: y + 0.78, w: 1.6, h: 0.3, fontSize: 15, bold: true, color: C.text1, valign: "middle" });
      const runs = ps.map((p, j) => ({ text: p, options: { bullet: { indent: 12 }, breakLine: j < ps.length - 1, paraSpaceAfter: 3 } }));
      text(s, runs, { x: x0 + 2.05, y: y + 0.15, w: w - 2.35, h: rh - 0.3, fontSize: 14, color: C.text1, valign: "middle", lineSpacingMultiple: 1.12 });
    });
    s.addNotes("담당 영역 상세 1 — 재무·회계 / 인사·노무 / IR·투자");
  }

  // =============== 6. 담당 영역 상세 2 ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Work" });
    s.addText("담당 영역 상세 2", { placeholder: "title" });
    text(s, "정부사업 / 해외업무 / 총무", { x: 4.5, y: TITLE_Y + 0.1, w: 5.0, h: 0.45, fontSize: 14, color: C.text2, align: "right", valign: "middle" });
    // left tall tile: 정부사업 with 5-step flow
    {
      const x = 0.5, y = 1.1, w = 4.4, h = 4.05;
      tile(s, x, y, w, h, TINT[0]);
      badge(s, x + 0.3, y + 0.25, 0.5, PASTEL[0], "FiFileText");
      text(s, "정부사업", { x: x + 0.95, y: y + 0.25, w: 2.5, h: 0.5, fontSize: 17, bold: true, color: C.text1, valign: "middle" });
      const ps = [
        "공고 분석 → 사업계획서 작성 → 신청 → 협약 → 정산까지 전 주기를 실무로 수행합니다.",
        "입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다.",
      ];
      const runs = ps.map((p, j) => ({ text: p, options: { bullet: { indent: 12 }, breakLine: j < ps.length - 1, paraSpaceAfter: 3 } }));
      text(s, runs, { x: x + 0.3, y: y + 0.9, w: w - 0.6, h: 1.25, fontSize: 14, color: C.text1, valign: "top", lineSpacingMultiple: 1.12 });
      vsteps(s, ["공고 분석", "사업계획서", "신청", "협약", "정산"], x + 0.32, y + 2.25, w - 0.6, 0.36, C.background1, 13);
    }
    // right: 해외업무 / 총무
    const right = [
      ["FiGlobe", "해외업무", ["해외 자회사의 인보이스와 용역계약 증빙을 관리하고, 해외 사업 행사 참가와 해외 출장 서류 실무를 지원합니다."]],
      ["FiHome", "총무", ["본점 이전 같은 법인 행정 절차와 기업 인증 관리, 사무실·자산 관리를 담당합니다.", "대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다."]],
    ];
    right.forEach(([ic, h, ps], i) => {
      const x = 5.1, y = 1.1 + i * (1.95 + 0.15), w = 4.4, th = 1.95;
      tile(s, x, y, w, th, TINT[i + 1]);
      badge(s, x + 0.3, y + 0.25, 0.5, PASTEL[i + 1], ic);
      text(s, h, { x: x + 0.95, y: y + 0.25, w: 2.5, h: 0.5, fontSize: 17, bold: true, color: C.text1, valign: "middle" });
      const runs = ps.map((p, j) => ({ text: p, options: { bullet: { indent: 12 }, breakLine: j < ps.length - 1, paraSpaceAfter: 3 } }));
      text(s, runs, { x: x + 0.3, y: y + 0.88, w: w - 0.6, h: th - 1.0, fontSize: 14, color: C.text1, valign: "top", lineSpacingMultiple: 1.12 });
    });
    s.addNotes("담당 영역 상세 2 — 정부사업 5단계 흐름 / 해외업무 / 총무");
  }

  // =============== 7-11. 사례 ===============
  pres.addSection({ title: "Cases" });
  const cases = [
    {
      no: "사례 1", title: '1인 경영지원을 "인계 가능한 시스템"으로 바꾸기', tint: 0,
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 커스텀 AI 스킬로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다.",
      icon: "FiDatabase",
    },
    {
      no: "사례 2", title: "단일 통화 자금 시트를 의사결정 도구로 재설계", tint: 1,
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리하고 있어, 경영진의 자금 판단 질문에 바로 답하기 어려웠습니다.",
      did: '대시보드·통화별 입출금·정기 비용·미지급 항목으로 시트를 나누고, 다중 통화 환산과 "부족액 식별 → 송금액 계산 → 송금 후 잔액" 시뮬레이션을 자동화했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.',
      flow: ["부족액 식별", "송금액 계산", "송금 후 잔액"],
      result: "경영진의 자금 의사결정 질문에 한 화면에서 바로 답할 수 있는 상태가 됐습니다.",
      icon: "FiBarChart2",
    },
    {
      no: "사례 3", title: "법률 1차 검토를 법령 원문 기반으로 내재화", tint: 2,
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했습니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 실무 상식만으로는 놓칠 수 있는 별도 조항을 미리 잡게 됐습니다.",
      icon: "FiBook",
    },
    {
      no: "사례 4", title: '외부로 나가기 전 "받은 산출물"을 역검증하는 습관', tint: 0,
      bg: "회계법인·기관·플랫폼에서 받은 자료나 내부 문서가 확인 없이 외부로 나가면 세금·등기·평가에 바로 영향을 줍니다.",
      did: "회계법인 산출물, 기관·플랫폼 자료, 계약서와 내부 문서를 원장·법령 원문과 대조한 뒤 제출하는 절차를 분기 내내 운영했습니다.",
      result: "기간 오류, 견적 오기, 정관과 계약의 충돌 같은 문제를 외부 제출 전에 찾아냈습니다.",
      icon: "FiSearch",
    },
    {
      no: "사례 5", title: "예산 대비 실적 관리 워크북 구축", tint: 1,
      bg: "월별 지출 추적이 여러 엑셀 파일에 흩어져 있어 예산을 통제하기 어려웠습니다.",
      did: "실제 계정원장을 기준으로 계정과목을 정리해 카테고리로 묶었습니다. 월별 예산 수립 → 실적 입력 → 차이 분석 흐름을 자동화했고, 이후 중앙 가정 시트와 연동되도록 확장했습니다.",
      flow: ["예산 수립", "실적 입력", "차이 분석"],
      result: "예산 대비 실적을 한 체계에서 관리하는 관리회계 기반을 만들었습니다.",
      icon: "FiGrid",
    },
  ];
  for (const cs of cases) {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Cases" });
    s.addText(cs.title, { placeholder: "title" });
    chip(s, cs.no, 8.4, 0.42, 1.1, 0.4, PASTEL[cs.tint], C.text1, 11);
    const t = cs.tint;
    const LX = 0.5, LW = 5.0, RX = 5.7, RW = 3.8, Y0 = 1.75, YEND = 5.2;
    // 배경
    const bgH = 1.2;
    tile(s, LX, Y0, LW, bgH, C.background2);
    label(s, "배경", LX + 0.25, Y0 + 0.17, 2);
    text(s, cs.bg, { x: LX + 0.25, y: Y0 + 0.44, w: LW - 0.5, h: bgH - 0.5, fontSize: 13, color: C.text1, valign: "top", lineSpacingMultiple: 1.12 });
    // 한 일
    const dY = Y0 + bgH + 0.15, dH = YEND - dY;
    tile(s, LX, dY, LW, dH, TINT[t]);
    label(s, "한 일", LX + 0.25, dY + 0.17, 2);
    const flowH = cs.flow ? 0.5 : 0;
    text(s, cs.did, { x: LX + 0.25, y: dY + 0.44, w: LW - 0.5, h: dH - 0.5 - flowH, fontSize: 13, color: C.text1, valign: "top", lineSpacingMultiple: 1.12 });
    if (cs.flow) flow(s, cs.flow, LX + 0.25, dY + dH - 0.6, 0.34, C.background1, 11);
    // 결과 — full-height accent tile with badge
    tile(s, RX, Y0, RW, YEND - Y0, PASTEL[t]);
    badge(s, RX + 0.3, Y0 + 0.3, 0.7, C.background1, cs.icon);
    text(s, "결과", { x: RX + 1.15, y: Y0 + 0.3, w: 2.0, h: 0.7, fontSize: 20, bold: true, color: C.text1, valign: "middle" });
    text(s, cs.result, { x: RX + 0.3, y: Y0 + 1.25, w: RW - 0.6, h: YEND - Y0 - 1.45, fontSize: 15, bold: true, color: C.text1, valign: "top", lineSpacingMultiple: 1.25 });
    s.addNotes(`${cs.no} — ${cs.title}`);
  }

  // =============== 12. 스킬·도구 ===============
  pres.addSection({ title: "Skills" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Skills" });
    s.addText("스킬·도구", { placeholder: "title" });
    const rows = [
      ["FiGrid", "엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·자금 시뮬레이션 설계"],
      ["FiFolder", "옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인"],
      ["FiCpu", "AI 업무 자동화", "Claude 커스텀 스킬 설계·운영, 예약 작업 자동화, MCP 연동(법령 원문 조회, 옵시디언 연동)"],
      ["FiMessageCircle", "협업 도구", "Gmail, Google Drive, Slack"],
      ["FiServer", "ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)"],
      ["FiEdit3", "문서 표준화", "공문 서식, docx·pptx 폰트 표준, 인보이스·인사·급여 템플릿, SOP"],
      ["FiClipboard", "실무 지식", "결산·세무 신고 대응, 주총·이사회·등기, 정부사업 전 주기"],
      ["FiGlobe", "언어", "비즈니스 영어(미국 현지 근무 경험)"],
    ];
    const rh = 0.48, gap = 0.05, y0 = 1.0, x0 = 0.5, w = 9.0;
    rows.forEach(([ic, h, d], i) => {
      const y = y0 + i * (rh + gap);
      tile(s, x0, y, w, rh, TINT[i % 3], { r: 0.1 });
      badge(s, x0 + 0.15, y + (rh - 0.34) / 2, 0.34, PASTEL[i % 3], ic);
      text(s, h, { x: x0 + 0.65, y, w: 1.55, h: rh, fontSize: 13, bold: true, color: C.text1, valign: "middle" });
      text(s, d, { x: x0 + 2.25, y, w: w - 2.45, h: rh, fontSize: 12, color: C.text1, valign: "middle", lineSpacingMultiple: 1.05 });
    });
    s.addNotes("스킬·도구 8행");
  }

  // =============== 13. 타임라인 ===============
  pres.addSection({ title: "Timeline" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Timeline" });
    s.addText("타임라인", { placeholder: "title" });
    text(s, "최신이 위", { x: 6.5, y: TITLE_Y + 0.1, w: 3.0, h: 0.45, fontSize: 14, color: C.text2, align: "right", valign: "middle" });
    const rows = [
      ["FiBriefcase", "2022 – 현재", "IT 스타트업", "경영지원팀 매니저"],
      ["FiMapPin", "2019 – 2022", "미국 소재 유통기업", "재고관리 인턴 → 구매(Purchasing) 담당"],
      ["FiAward", "2018 – 2019", "미국 비즈니스 취업연수 과정 수료", ""],
      ["FiCompass", "2017", "중국 교환학생", ""],
      ["FiBookOpen", "2013 – 2019", "학부 — 국제통상학 전공 (중어중문학 복수전공)", ""],
    ];
    const rh = 0.72, gap = 0.1, y0 = 1.1;
    s.addShape(pres.shapes.LINE, { x: 0.5 + 0.3, y: y0 + 0.36, w: 0, h: (rh + gap) * 4, line: { color: C.accent4, width: 2 }, objectName: nm("Connector") });
    rows.forEach(([ic, date, h, d], i) => {
      const y = y0 + i * (rh + gap);
      badge(s, 0.5, y + 0.06, 0.6, PASTEL[i % 3], ic);
      tile(s, 1.4, y, 8.1, rh, TINT[i % 3]);
      text(s, date, { x: 1.65, y: y + 0.1, w: 1.3, h: 0.52, fontSize: 12, bold: true, color: C.text2, valign: "middle" });
      if (d) {
        text(s, h, { x: 3.0, y: y + 0.08, w: 6.3, h: 0.28, fontSize: 14, bold: true, color: C.text1 });
        text(s, d, { x: 3.0, y: y + 0.38, w: 6.3, h: 0.3, fontSize: 12, color: C.text2 });
      } else {
        text(s, h, { x: 3.0, y: y + 0.1, w: 6.3, h: 0.52, fontSize: 14, bold: true, color: C.text1, valign: "middle" });
      }
    });
    s.addNotes("타임라인 — 2013–현재, 최신이 위");
  }

  // =============== 14. 연락 ===============
  pres.addSection({ title: "Outro" });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Outro" });
    tile(s, 0.5, 0.55, 5.35, 4.5, C.background1, { r: 0.2 });
    text(s, "연락", { x: 0.9, y: 0.95, w: 4.6, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    text(s, "양남규", { x: 0.9, y: 1.35, w: 4.6, h: 1.0, fontSize: 54, bold: true, color: C.text1 });
    text(s, "Duke Yang", { x: 0.9, y: 2.35, w: 4.6, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
    text(s, "경영지원 매니저", { x: 0.9, y: 2.8, w: 4.6, h: 0.35, fontSize: 14, color: C.text1 });
    text(s, "© 2026 Duke Yang", { x: 0.9, y: 4.45, w: 4.4, h: 0.3, fontSize: 10, color: C.text2 });
    // right: email tile + three pastel motif tiles
    const gx = 6.15, tw = 3.35;
    {
      const y = 0.55, th = 2.6;
      tile(s, gx, y, tw, th, C.background1, { r: 0.2 });
      badge(s, gx + (tw - 0.9) / 2, y + 0.35, 0.9, TINT[2], "FiMail");
      text(s, "이메일", { x: gx + 0.2, y: y + 1.35, w: tw - 0.4, h: 0.3, fontSize: 12, bold: true, color: C.text2, align: "center", valign: "middle" });
      text(s, "namgyu1412@naver.com", { x: gx + 0.2, y: y + 1.7, w: tw - 0.4, h: 0.45, fontSize: 16, bold: true, color: C.text1, align: "center", valign: "middle" });
    }
    const small = [["FiDatabase", 0], ["FiBookOpen", 1], ["FiRepeat", 2]];
    const sw = (tw - 0.3) / 3, sy = 0.55 + 2.6 + 0.2, sh = 4.5 + 0.55 - sy;
    small.forEach(([ic, k], i) => {
      const x = gx + i * (sw + 0.15);
      tile(s, x, sy, sw, sh, C.background1, { r: 0.2 });
      badge(s, x + (sw - 0.6) / 2, sy + (sh - 0.6) / 2, 0.6, TINT[k], ic);
    });
    s.addNotes("연락 — namgyu1412@naver.com, © 2026 Duke Yang");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
