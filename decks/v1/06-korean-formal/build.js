// 06-korean-formal — "Korean Formal" portfolio deck (양남규)
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
pres.title = "양남규 포트폴리오 — Korean Formal";
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
        options: { name: "title", type: "title", x: 0.5, y: 0.64, w: 7.7, h: 0.74, fontSize: 36, bold: true, color: NAVY, margin: 0, valign: "middle", align: "left" },
        text: "",
      },
    },
    { text: { text: "양남규 포트폴리오 · GXC 경영지원팀", options: { x: 0.5, y: FOOT_Y, w: 5, h: 0.25, fontSize: 9, color: MUTED, margin: 0, valign: "middle" } } },
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

(async () => {
  const navyIcon = (name) => icon(name, P.white);
  const paleIcon = (name) => icon(name, P.navy);

  // ===== 1. Cover =====
  pres.addSection({ title: "표지" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
    n++;
    // motif label (sky)
    s.addText("PORTFOLIO · 2026", { x: 0.7, y: 0.6, w: 1.75, h: 0.3, fill: { color: SKY }, color: NAVY, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Cover label", charSpacing: 2 });
    // decorative grid of pale boxes on the right (구획)
    const gx = 6.4, gy = 0.6;
    const cells = [
      ["재무 · 회계", "FaCalculator"],
      ["IR · 투자", "FaHandshake"],
      ["HR · 노무", "FaUsers"],
      ["컴플라이언스", "FaBalanceScale"],
    ];
    for (let i = 0; i < 4; i++) {
      const cx = gx + (i % 2) * 1.5, cy = gy + Math.floor(i / 2) * 1.5;
      s.addShape(pres.ShapeType.rect, { x: cx, y: cy, w: 1.4, h: 1.4, fill: { color: P.navy }, line: { color: SKY, width: 0.75 }, objectName: "Grid cell" });
      s.addImage({ data: await icon(cells[i][1], P.sky), x: cx + 0.45, y: cy + 0.3, w: 0.5, h: 0.5, objectName: cells[i][0] + " 아이콘" });
      s.addText(cells[i][0], { x: cx, y: cy + 0.9, w: 1.4, h: 0.3, fontSize: 10, color: SKY, align: "center", margin: 0, isTextBox: true });
    }
    txt(s, "Hello, I'm", 0.7, 1.35, 5, 0.4, { fontSize: 18, color: SKY });
    s.addText("양남규", { x: 0.7, y: 1.75, w: 5.4, h: 1.1, fontSize: 60, bold: true, color: WHITE, margin: 0, valign: "middle", isTextBox: true, objectName: "이름" });
    txt(s, "Finance Operations Specialist", 0.7, 2.9, 5.5, 0.45, { fontSize: 20, color: SKY, bold: true });
    txt(s, "스타트업의 성장을 뒷받침하는 경영지원 전문가", 0.7, 3.4, 5.5, 0.4, { fontSize: 16, color: WHITE });
    // bottom info box
    s.addShape(pres.ShapeType.rect, { x: 0.7, y: 4.35, w: 8.6, h: 0.6, fill: { color: P.navy }, line: { color: SKY, width: 0.75 }, objectName: "Info panel" });
    s.addText(
      [
        { text: "재무 · IR · HR · 컴플라이언스", options: { color: WHITE, bold: true } },
        { text: "     /     GXC 경영지원팀     /     2026", options: { color: SKY } },
      ],
      { x: 0.9, y: 4.35, w: 8.2, h: 0.6, fontSize: 13, margin: 0, valign: "middle", isTextBox: true, objectName: "보조 정보" }
    );
  }

  // ===== 2. About Me =====
  pres.addSection({ title: "소개" });
  {
    const s = content("소개", "01  ABOUT", "About Me");
    // left: paragraph box
    box(s, 0.5, 1.5, 5.6, 3.6);
    s.addShape(pres.ShapeType.rect, { x: 0.8, y: 1.8, w: 0.7, h: 0.7, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Icon tile" });
    s.addImage({ data: await navyIcon("FaUserTie"), x: 0.95, y: 1.95, w: 0.4, h: 0.4, objectName: "프로필 아이콘" });
    txt(s, "양남규", 1.7, 1.8, 3, 0.4, { fontSize: 18, bold: true, color: NAVY });
    txt(s, "Finance Operations Specialist · GXC 경영지원팀", 1.7, 2.18, 4.2, 0.3, { fontSize: 11, color: MUTED });
    txt(
      s,
      "GXC에서 경영지원 업무를 담당하며, 재무/회계 관리부터 투자 관련 법률 검토, 인사/노무 행정, 정부사업 관리까지 폭넓은 영역을 수행하고 있습니다. 체계적인 학습과 실무 경험을 병행하며 스타트업의 성장에 필요한 오퍼레이션 역량을 꾸준히 쌓아가고 있습니다.",
      0.8, 2.75, 5.0, 2.2,
      { fontSize: 15, lineSpacingMultiple: 1.35 }
    );
    // right: tags
    label(s, "핵심 역량", 6.4, 1.5, 0.9);
    const tags = ["재무/회계", "IR/투자", "HR/노무", "정부사업", "Excel 고급", "데이터분석", "컴플라이언스"];
    tags.forEach((t, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 6.4 + col * 1.6, y = 1.95 + row * 0.5;
      const last = i === tags.length - 1;
      chip(s, t, x, y, last ? 3.1 : 1.5, 0.38, { fontSize: 12, bold: true, fill: PALE });
    });
    txt(s, "핵심 역량 태그 7개", 6.4, 3.98, 3.1, 0.3, { fontSize: 10, color: MUTED });
    s.addShape(pres.ShapeType.rect, { x: 6.4, y: 4.45, w: 3.1, h: 0.65, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Quote tile" });
    txt(s, "스타트업 운영 기반을 만드는 경영지원", 6.55, 4.45, 2.9, 0.65, { fontSize: 12, color: WHITE, bold: true, valign: "middle" });
  }

  // ===== 3. 업무 영역 =====
  pres.addSection({ title: "업무 영역" });
  {
    const s = content("업무 영역", "02  WORK AREAS", "업무 영역");
    const cards = [
      ["재무 / 회계", "FaCalculator", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"]],
      ["IR / 투자", "FaHandshake", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"]],
      ["HR / 노무", "FaUsers", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"]],
      ["정부사업 / 기술", "FaLandmark", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"]],
    ];
    const cw = 2.1, gap = 0.2, x0 = 0.5, y0 = 1.5;
    for (let i = 0; i < 4; i++) {
      const x = x0 + i * (cw + gap);
      // header cell
      s.addShape(pres.ShapeType.rect, { x, y: y0, w: cw, h: 1.25, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Card head" });
      s.addImage({ data: await navyIcon(cards[i][1]), x: x + 0.25, y: y0 + 0.22, w: 0.5, h: 0.5, objectName: cards[i][0] + " 아이콘" });
      txt(s, String(i + 1).padStart(2, "0"), x + cw - 0.65, y0 + 0.18, 0.5, 0.35, { fontSize: 14, bold: true, color: SKY, align: "right" });
      txt(s, cards[i][0], x + 0.25, y0 + 0.8, cw - 0.4, 0.35, { fontSize: 16, bold: true, color: WHITE });
      // body cell
      box(s, x, y0 + 1.25, cw, 2.35, { fill: WHITE });
      cards[i][2].forEach((item, j) => {
        const yy = y0 + 1.25 + 0.22 + j * 0.5;
        s.addShape(pres.ShapeType.rect, { x: x + 0.2, y: yy + 0.11, w: 0.12, h: 0.12, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Marker" });
        txt(s, item, x + 0.42, yy, cw - 0.5, 0.35, { fontSize: 14, valign: "middle" });
      });
    }
  }

  // ===== 4. Skills =====
  pres.addSection({ title: "역량" });
  {
    const s = content("역량", "03  SKILLS", "Skills & Expertise");
    const groups = [
      ["재무 / 회계", [["세무 관리", 75], ["월말 결산", 70], ["전표 처리", 80], ["재무제표 분석", 65]]],
      ["IR / 투자", [["주주간계약(SHA)", 70], ["스톡옵션 설계", 65], ["TTC 설계", 60], ["투자 라운드", 60]]],
      ["HR / 노무", [["인사 관리", 75], ["근로계약 관리", 70], ["노동법 이해", 65], ["급여 체계 설계", 60]]],
      ["정부사업 / 기술", [["정부 보조금", 70], ["컴플라이언스", 65], ["Excel (고급)", 85], ["데이터 분석", 60]]],
    ];
    const bw = 4.4, bh = 1.75, gx = 0.2, gy = 0.1, x0 = 0.5, y0 = 1.5;
    for (let i = 0; i < 4; i++) {
      const x = x0 + (i % 2) * (bw + gx), y = y0 + Math.floor(i / 2) * (bh + gy);
      box(s, x, y, bw, bh, { fill: WHITE });
      label(s, groups[i][0], x + 0.15, y + 0.12, 1.15, 0.26);
      txt(s, "숙련도 (%)", x + bw - 1.2, y + 0.12, 1.05, 0.26, { fontSize: 9, color: MUTED, align: "right", valign: "middle" });
      const data = [{ name: groups[i][0], labels: groups[i][1].map((d) => d[0]).reverse(), values: groups[i][1].map((d) => d[1]).reverse() }];
      s.addChart(pres.ChartType.bar, data, {
        x: x + 0.1, y: y + 0.4, w: bw - 0.2, h: bh - 0.45,
        barDir: "bar", barGapWidthPct: 45,
        chartColors: [P.navy],
        showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9, dataLabelColor: P.navy, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0",
        catAxisLabelFontSize: 10, catAxisLabelColor: P.text, catAxisLabelFontFace: "+mn-lt",
        catAxisLineShow: false, catGridLine: { style: "none" }, valGridLine: { style: "none" },
        valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valAxisLabelFontFace: "+mn-lt",
        showLegend: false, showTitle: false,
        objectName: groups[i][0] + " 숙련도 차트",
      });
    }
  }

  // ===== 5. Excel Mastery =====
  pres.addSection({ title: "Excel" });
  {
    const s = content("Excel", "04  EXCEL", "Excel Mastery");
    txt(s, "실무에서 활용하는 엑셀 고급 함수와 자동화 기법", 0.5, 1.42, 8, 0.3, { fontSize: 12, color: MUTED });
    const rows = [
      ["XLOOKUP", "FaSearch", "급여 대장 자동 매칭 / 거래처별 매출 추적 / 계정과목 자동 분류", ["VLOOKUP 대체", "양방향 검색", "에러 처리 내장"]],
      ["피벗테이블", "FaTable", "부서별 예산 집행 분석 / 월별 비용 추이 리포트", ["크로스탭 분석", "슬라이서 필터", "자동 갱신"]],
      ["조건부 서식", "FaFillDrip", "예산 초과 자동 알림 (달성률 100% 초과 빨강, 90% 미만 노랑)", ["데이터 막대", "아이콘 세트", "색조 규칙"]],
      ["자동화 매크로(VBA)", "FaCogs", "월말 결산 자동화 / 급여명세서 자동 발송", ["VBA 매크로", "반복 업무 제거", "3시간→15분"]],
    ];
    const rh = 0.78, gap = 0.08, y0 = 1.85;
    rows.forEach((r, i) => {
      const y = y0 + i * (rh + gap);
      box(s, 0.5, y, 9.0, rh, { fill: i % 2 === 0 ? PALE : WHITE });
      s.addShape(pres.ShapeType.rect, { x: 0.62, y: y + 0.12, w: 0.54, h: 0.54, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Icon tile" });
    });
    for (let i = 0; i < rows.length; i++) {
      const y = y0 + i * (rh + gap);
      s.addImage({ data: await navyIcon(rows[i][1]), x: 0.74, y: y + 0.24, w: 0.3, h: 0.3, objectName: rows[i][0] + " 아이콘" });
      txt(s, rows[i][0], 1.35, y + 0.1, 2.2, 0.3, { fontSize: 15, bold: true, color: NAVY });
      txt(s, rows[i][2], 1.35, y + 0.42, 4.9, 0.3, { fontSize: 11, color: TEXT });
      rows[i][3].forEach((b, j) => chip(s, b, 6.35 + j * 1.05, y + 0.26, 0.98, 0.26, { fontSize: 9 }));
    }
  }

  // ===== 6. 자동화 성과 =====
  pres.addSection({ title: "성과" });
  {
    const s = content("성과", "05  RESULTS", "자동화 성과");
    const stats = [
      [["3시간 → ", "15분"], "월말 결산 보고서 생성", "전표 정리 → 집계 → 보고서, 버튼 하나"],
      [["3일 → ", "1.2일"], "월말 결산 전체 소요 시간", "약 60% 단축"],
      [["", "3단계"], "자동화 프로세스", "전표 정리 · 계정과목별 집계 · 보고서 출력(PDF 저장·메일 발송)"],
    ];
    const sw = 2.87, gap = 0.195, y0 = 1.5, sh = 1.55;
    stats.forEach((st, i) => {
      const x = 0.5 + i * (sw + gap);
      s.addShape(pres.ShapeType.rect, { x, y: y0, w: sw, h: sh, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Stat tile" });
      s.addText(
        [
          { text: st[0][0], options: { fontSize: 16, color: SKY, bold: true } },
          { text: st[0][1], options: { fontSize: 30, color: WHITE, bold: true } },
        ],
        { x: x + 0.2, y: y0 + 0.12, w: sw - 0.3, h: 0.68, margin: 0, valign: "bottom", isTextBox: true, objectName: "Stat " + st[1] }
      );
      txt(s, st[1], x + 0.2, y0 + 0.82, sw - 0.4, 0.3, { fontSize: 12, bold: true, color: SKY });
      txt(s, st[2], x + 0.2, y0 + 1.1, sw - 0.4, 0.42, { fontSize: 10, color: WHITE });
    });
    // pivot table
    label(s, "예시 피벗 표 · 부서별 예산 집행", 0.5, 3.25, 2.3);
    txt(s, "단위: 천원", 7.5, 3.25, 2.0, 0.26, { fontSize: 9, color: MUTED, align: "right", valign: "middle" });
    const head = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } });
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ color: TEXT, align: "right" }, o) });
    const delta = (t, up) => ({ text: t, options: { color: up ? NAVY : RED, bold: true, align: "right" } });
    const tbl = [
      [head("부서"), head("Q1"), head("Q2"), head("증감률")],
      [cell("개발팀", { align: "left" }), cell("45,200"), cell("52,800"), delta("+16.8%", true)],
      [cell("마케팅", { align: "left" }), cell("32,100"), cell("28,500"), delta("-11.2%", false)],
      [cell("경영지원", { align: "left" }), cell("18,300"), cell("19,100"), delta("+4.4%", true)],
      [cell("합계", { align: "left", bold: true, fill: { color: PALE } }), cell("95,600", { bold: true, fill: { color: PALE } }), cell("100,400", { bold: true, fill: { color: PALE } }), delta("+5.0%", true)],
    ];
    tbl[4][3].options.fill = { color: PALE };
    s.addTable(tbl, {
      x: 0.5, y: 3.6, w: 9.0, colW: [2.7, 2.1, 2.1, 2.1], rowH: 0.3,
      fontSize: 11, border: { type: "solid", color: P.sky, pt: 0.75 }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1],
      objectName: "부서별 예산 피벗 표",
    });
  }

  // ===== 7-9. Cases =====
  pres.addSection({ title: "Case Study" });
  const cases = [
    {
      num: "01", cat: "IR / 투자", title: "SHA 주주간계약 검토",
      bg: "GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.",
      steps: ["ROFR(우선매수권) 조항의 실무 적용 시나리오 분석", "Drag-Along · Tag-Along 권리의 행사 조건 정리", "희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)", "투자자 동의권(Consent Rights) 범위 및 영향 분석", "법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화"],
      result: "5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.",
      kpis: [["5개", "핵심 조항 분석"], ["1건", "가이드 문서"]],
      learn: "법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다.",
      icon: "FaFileContract",
    },
    {
      num: "02", cat: "재무 / 회계", title: "월말 결산 프로세스 혁신",
      bg: "매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.",
      steps: ["기존 결산 프로세스 전수조사 및 병목 구간 파악", "전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)", "계정과목별 자동 집계 템플릿 구축", "VBA 매크로를 활용한 보고서 자동 생성", "결산 체크리스트 및 SOP 문서 작성"],
      result: "결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.",
      kpis: [["-60%", "소요 시간"], ["SOP", "프로세스 표준화"]],
      learn: "Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다.",
      icon: "FaCalculator",
    },
    {
      num: "03", cat: "HR / 노무", title: "TTC 기반 보상 체계 설계",
      bg: "스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.",
      steps: ["국내 IT 스타트업 보상 시장조사 및 벤치마킹", "TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)", "직급/직무별 Pay Band 설계", "스톡옵션 배분 기준 및 베스팅 스케줄 수립", "경영진 보고 및 의사결정 자료 준비"],
      result: "전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.",
      kpis: [["전사", "적용 범위"], ["TTC", "체계 도입"]],
      learn: "보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다.",
      icon: "FaUsers",
    },
  ];
  for (const cs of cases) {
    const s = content("Case Study", `06  CASE ${cs.num}`, cs.title);
    chip(s, cs.cat, 8.3, 0.84, 1.2, 0.32, { fontSize: 11, bold: true, fill: PALE });
    const BODY = 12;
    // Row A: 배경 + 지표
    box(s, 0.5, 1.5, 5.9, 1.0);
    label(s, "배경", 0.62, 1.62, 0.6);
    txt(s, cs.bg, 1.35, 1.6, 4.9, 0.85, { fontSize: BODY, lineSpacingMultiple: 1.1 });
    for (let i = 0; i < 2; i++) {
      const x = 6.55 + i * 1.5;
      s.addShape(pres.ShapeType.rect, { x, y: 1.5, w: 1.45, h: 1.0, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "KPI tile" });
      txt(s, cs.kpis[i][0], x + 0.12, 1.55, 1.25, 0.55, { fontSize: 24, bold: true, color: WHITE, valign: "middle" });
      txt(s, cs.kpis[i][1], x + 0.12, 2.1, 1.25, 0.3, { fontSize: 10, color: SKY });
    }
    // Row B: 과정
    box(s, 0.5, 2.62, 9.0, 1.4, { fill: WHITE });
    label(s, "과정", 0.62, 2.74, 0.6);
    s.addText(bullets(cs.steps, { numbered: true, space: 3, run: { fontSize: BODY, color: TEXT } }), {
      x: 1.35, y: 2.72, w: 7.5, h: 1.25, margin: 0, valign: "top", isTextBox: true, objectName: "과정 목록",
    });
    s.addImage({ data: await paleIcon(cs.icon), x: 8.85, y: 2.78, w: 0.45, h: 0.45, objectName: "케이스 아이콘" });
    // Row C: 결과 + 배운 점
    box(s, 0.5, 4.14, 4.1, 1.0);
    label(s, "결과", 0.62, 4.26, 0.6);
    txt(s, cs.result, 1.35, 4.24, 3.15, 0.85, { fontSize: BODY, lineSpacingMultiple: 1.1 });
    box(s, 4.75, 4.14, 4.75, 1.0, { fill: WHITE });
    label(s, "배운 점", 4.87, 4.26, 0.7);
    txt(s, cs.learn, 5.7, 4.24, 3.7, 0.85, { fontSize: BODY, lineSpacingMultiple: 1.1 });
  }

  // ===== 10. Career Timeline =====
  pres.addSection({ title: "경력" });
  {
    const s = content("경력", "07  TIMELINE", "Career Timeline");
    const items = [
      ["2026.04", "스톡옵션 세무 체계 정립", "스톡옵션 법률 구조 및 과세 체계 분석 완료. 임직원 보상 설계의 세무 관점 가이드라인 수립."],
      ["2026.04", "ABTC 인증 처리 완료", "APEC 기업인 여행카드(ABTC) 관련 문서 작업 및 인증 프로세스 완료."],
      ["2026.04", "SHA(주주간계약) 분석", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 등 주주간계약의 핵심 조항 분석 및 실무 적용."],
      ["2026.04", "IT 스타트업 런웨이 관리 체계 수립", "현금흐름 기반 런웨이 관리 프레임워크 학습 및 GXC 적용 방안 도출."],
      ["2026 ~", "GXC 경영지원팀 합류", "재무/회계, IR, HR, 정부사업 등 경영지원 전반 업무 담당. 스타트업 운영 기반 구축."],
    ];
    const rh = 0.66, y0 = 1.5;
    // table header
    s.addShape(pres.ShapeType.rect, { x: 0.5, y: y0, w: 9.0, h: 0.3, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Table head" });
    txt(s, "시기", 0.65, y0, 1.2, 0.3, { fontSize: 10, bold: true, color: WHITE, valign: "middle" });
    txt(s, "활동", 2.45, y0, 3, 0.3, { fontSize: 10, bold: true, color: WHITE, valign: "middle" });
    txt(s, "내용", 5.5, y0, 3, 0.3, { fontSize: 10, bold: true, color: WHITE, valign: "middle" });
    txt(s, "▲ 최근  ·  ▼ 과거", 8.0, y0, 1.4, 0.3, { fontSize: 9, color: SKY, valign: "middle", align: "right" });
    // vertical connector line
    s.addShape(pres.ShapeType.line, { x: 2.05, y: y0 + 0.3, w: 0, h: rh * items.length, line: { color: SKY, width: 1 }, objectName: "Timeline line" });
    items.forEach((it, i) => {
      const y = y0 + 0.3 + i * rh;
      box(s, 0.5, y, 9.0, rh, { fill: i % 2 === 0 ? PALE : WHITE, line: false });
      txt(s, it[0], 0.65, y, 1.2, rh, { fontSize: 13, bold: true, color: NAVY, valign: "middle" });
      s.addShape(pres.ShapeType.rect, { x: 1.96, y: y + rh / 2 - 0.09, w: 0.18, h: 0.18, fill: { color: i === items.length - 1 ? RED : NAVY }, line: { color: WHITE, width: 1.5 }, objectName: "Timeline marker" });
      txt(s, it[1], 2.45, y, 2.95, rh, { fontSize: 13, bold: true, color: TEXT, valign: "middle" });
      txt(s, it[2], 5.5, y, 3.85, rh, { fontSize: 10.5, color: MUTED, valign: "middle", lineSpacingMultiple: 1.1 });
    });
  }

  // ===== 11. Certifications & Learning =====
  pres.addSection({ title: "학습" });
  {
    const s = content("학습", "08  LEARNING", "Certifications & Learning");
    // left: certifications
    s.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 0.55, h: 0.55, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Icon tile" });
    s.addImage({ data: await navyIcon("FaCertificate"), x: 0.62, y: 1.62, w: 0.31, h: 0.31, objectName: "자격증 아이콘" });
    txt(s, "자격증 (준비중)", 1.2, 1.5, 3, 0.3, { fontSize: 16, bold: true, color: NAVY });
    txt(s, "한국세무사회 · 2026.06 목표", 1.2, 1.8, 3, 0.25, { fontSize: 10, color: MUTED });
    const H = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "left", fontSize: 10 } });
    const Cl = (t, o = {}) => ({ text: t, options: Object.assign({ color: TEXT, fontSize: 11 }, o) });
    s.addTable(
      [
        [H("자격증"), H("발급기관"), H("목표")],
        [Cl("전산회계 1급", { bold: true }), Cl("한국세무사회"), Cl("2026.06 목표", { color: NAVY })],
        [Cl("전산세무 2급", { bold: true }), Cl("한국세무사회"), Cl("2026.06 목표", { color: NAVY })],
      ],
      { x: 0.5, y: 2.25, w: 4.1, colW: [1.5, 1.4, 1.2], rowH: 0.42, border: { type: "solid", color: P.sky, pt: 0.75 }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1], objectName: "자격증 표" }
    );
    // status chips under table
    box(s, 0.5, 3.75, 4.1, 1.35);
    label(s, "진행 상태", 0.62, 3.87, 0.8);
    txt(s, "두 자격증 모두 2026년 6월 취득을 목표로 준비 중입니다.", 0.62, 4.2, 3.9, 0.5, { fontSize: 12, lineSpacingMultiple: 1.1 });
    chip(s, "전산회계 1급 · 준비중", 0.62, 4.72, 1.85, 0.28, { fontSize: 10, bold: true });
    chip(s, "전산세무 2급 · 준비중", 2.6, 4.72, 1.85, 0.28, { fontSize: 10, bold: true });
    // right: recent learning
    s.addShape(pres.ShapeType.rect, { x: 5.0, y: 1.5, w: 0.55, h: 0.55, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Icon tile" });
    s.addImage({ data: await navyIcon("FaBookOpen"), x: 5.12, y: 1.62, w: 0.31, h: 0.31, objectName: "학습 아이콘" });
    txt(s, "최근 학습", 5.7, 1.5, 3, 0.3, { fontSize: 16, bold: true, color: NAVY });
    txt(s, "2026.04 · IR/투자 · 재무", 5.7, 1.8, 3, 0.25, { fontSize: 10, color: MUTED });
    s.addTable(
      [
        [H("일자"), H("학습 주제"), H("분야")],
        [Cl("2026.04.16", { color: MUTED }), Cl("스톡옵션 과세 체계", { bold: true }), Cl("IR/투자", { color: NAVY })],
        [Cl("2026.04.14", { color: MUTED }), Cl("스톡옵션 법률 구조", { bold: true }), Cl("IR/투자", { color: NAVY })],
        [Cl("2026.04.10", { color: MUTED }), Cl("주주간계약(SHA) 핵심 조항", { bold: true }), Cl("IR/투자", { color: NAVY })],
        [Cl("2026.04.08", { color: MUTED }), Cl("IT 스타트업 런웨이 관리", { bold: true }), Cl("재무", { color: NAVY })],
      ],
      { x: 5.0, y: 2.25, w: 4.5, colW: [1.15, 2.45, 0.9], rowH: 0.42, border: { type: "solid", color: P.sky, pt: 0.75 }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1], objectName: "최근 학습 표" }
    );
  }

  // ===== 12. Project Gallery =====
  pres.addSection({ title: "프로젝트" });
  {
    const s = content("프로젝트", "09  PROJECTS", "Project Gallery");
    const projects = [
      ["IR/투자", "SHA 주주간계약 검토", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 분석 및 투자 계약 실무 적용", "투자 계약 핵심 조항 가이드 문서화"],
      ["IR/투자", "스톡옵션 체계 설계", "법률 구조 분석, 과세 체계 정립, 임직원 보상 설계 가이드라인 수립", "스톡옵션 세무/법률 통합 가이드 완성"],
      ["HR/노무", "급여 체계 설계", "TTC(Total Target Compensation) 기반 급여 체계 설계 및 보상 구조 수립", "전사 보상 체계 프레임워크 구축"],
      ["정부사업", "ABTC 인증 처리", "APEC 기업인 여행카드 관련 서류 준비 및 인증 프로세스 관리", "인증 완료 및 프로세스 SOP 문서화"],
      ["재무/회계", "런웨이 관리 체계", "IT 스타트업 현금흐름 기반 런웨이 관리 프레임워크 구축", "월별 런웨이 모니터링 대시보드 구축"],
      ["재무/회계", "월말 결산 프로세스", "매월 재무제표 마감 및 결산 프로세스 체계화, 전표 처리 자동화", "결산 소요 시간 단축 및 프로세스 표준화"],
    ];
    const cw = 2.9, ch = 1.72, gx = 0.15, gy = 0.15, x0 = 0.5, y0 = 1.5;
    projects.forEach((p, i) => {
      const x = x0 + (i % 3) * (cw + gx), y = y0 + Math.floor(i / 3) * (ch + gy);
      box(s, x, y, cw, ch, { fill: i % 2 === 0 ? PALE : WHITE });
      label(s, p[0], x + 0.15, y + 0.13, 0.85, 0.24, { fontSize: 9 });
      txt(s, String(i + 1).padStart(2, "0"), x + cw - 0.6, y + 0.1, 0.45, 0.3, { fontSize: 12, bold: true, color: SKY, align: "right" });
      txt(s, p[1], x + 0.15, y + 0.42, cw - 0.3, 0.3, { fontSize: 14, bold: true, color: NAVY });
      txt(s, p[2], x + 0.15, y + 0.74, cw - 0.3, 0.5, { fontSize: 10, color: TEXT, lineSpacingMultiple: 1.1 });
      s.addText(
        [
          { text: "성과  ", options: { bold: true, color: NAVY } },
          { text: p[3], options: { color: TEXT } },
        ],
        { x: x + 0.15, y: y + 1.26, w: cw - 0.3, h: 0.4, fontSize: 10, margin: 0, valign: "top", isTextBox: true, objectName: "성과 " + p[1] }
      );
    });
  }

  // ===== 13. Closing =====
  pres.addSection({ title: "마무리" });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "마무리" });
    n++;
    s.addText("10  CONTACT", { x: 0.5, y: 0.34, w: 1.2, h: 0.26, fill: { color: NAVY }, color: WHITE, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Section label", charSpacing: 1 });
    s.addText("Let's Connect", { x: 0.5, y: 1.4, w: 6, h: 1.0, fontSize: 48, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true, objectName: "제목" });
    txt(s, "새로운 기회나 협업에 대해 이야기하고 싶으시다면 연락주세요.", 0.5, 2.45, 6, 0.4, { fontSize: 16, color: TEXT });
    const contacts = [["Email", "FaEnvelope"], ["LinkedIn", "FaLinkedin"]];
    for (let i = 0; i < 2; i++) {
      const x = 0.5 + i * 2.3;
      box(s, x, 3.15, 2.1, 0.7);
      s.addShape(pres.ShapeType.rect, { x: x + 0.12, y: 3.27, w: 0.46, h: 0.46, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "Icon tile" });
      s.addImage({ data: await navyIcon(contacts[i][1]), x: x + 0.22, y: 3.37, w: 0.26, h: 0.26, objectName: contacts[i][0] + " 아이콘" });
      txt(s, contacts[i][0], x + 0.72, 3.15, 1.3, 0.7, { fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    }
    // right decorative grid (motif echo)
    const gx = 6.9, gy = 1.4;
    const cells = ["재무", "IR", "HR", "컴플라이언스"];
    cells.forEach((t, i) => {
      const cx = gx + (i % 2) * 1.3, cy = gy + Math.floor(i / 2) * 1.3;
      s.addShape(pres.ShapeType.rect, { x: cx, y: cy, w: 1.2, h: 1.2, fill: { color: i === 3 ? NAVY : PALE }, line: { color: SKY, width: 0.75 }, objectName: "Grid cell" });
      txt(s, t, cx, cy, 1.2, 1.2, { fontSize: 13, bold: true, color: i === 3 ? WHITE : NAVY, align: "center", valign: "middle" });
    });
    txt(s, "© 2026 양남규", 0.5, 4.9, 4, 0.3, { fontSize: 11, color: MUTED });
    txt(s, "양남규 포트폴리오 · GXC 경영지원팀", 5.5, 4.9, 4.0, 0.3, { fontSize: 9, color: MUTED, align: "right" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT, "slides:", n);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
