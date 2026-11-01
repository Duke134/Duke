// 03-editorial — "Editorial" deck for 양남규 portfolio
// Run: NODE_PATH=<node_modules> node build.js   (from any cwd; paths are relative to this file)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fa = require("react-icons/fa");
const { applyTheme } = require(
  "/root/.claude/skills/synced/7d6f9636-edbe-41a2-aab0-0f84fceac3dd_307ca375-45f7-4235-8c53-4eacd5b231ab/pptx/scripts/apply_theme.js"
);

const OUT = path.join(__dirname, "deck.pptx");

// ---------- palette (brief: 03-editorial) ----------
const WHITE = "FFFFFF";
const BERRY = "6D2E46";
const ROSE = "A26769";
const CREAM = "ECE2D0";
const INK = "1F1A1C";

const THEME = {
  name: "Editorial",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: INK, lt1: WHITE, dk2: BERRY, lt2: CREAM,
    accent1: BERRY, accent2: ROSE, accent3: CREAM, accent4: INK, accent5: ROSE, accent6: BERRY,
    hlink: BERRY, folHlink: ROSE,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "양남규";
pres.title = "양남규 포트폴리오 — Editorial";
const C = pres.SchemeColor;
const INKc = C.text1, BERRYc = C.text2, WHITEc = C.background1, CREAMc = C.background2, ROSEc = C.accent2;

// ---------- geometry ----------
const BLOCK_W = 3.33;            // left 1/3 color block
const BX = 3.7, BW = 5.85;       // right body column
const BY = 1.7, BB = 5.2;        // body top / bottom

// ---------- icons ----------
async function icon(Component, hex, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Component, { color: "#" + hex, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- text width estimate (inches) for pill sizing ----------
function estW(text, pt) {
  const em = pt / 72;
  let w = 0;
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if (c >= 0x1100 && (c <= 0x11ff || (c >= 0x3130 && c <= 0x318f) || (c >= 0xac00 && c <= 0xd7af) || (c >= 0x4e00 && c <= 0x9fff) || (c >= 0x3000 && c <= 0x30ff))) w += 1.0 * em;
    else if (ch === " ") w += 0.3 * em;
    else if (/[A-Z]/.test(ch)) w += 0.68 * em;
    else if (/[0-9]/.test(ch)) w += 0.58 * em;
    else if (/[a-z]/.test(ch)) w += 0.54 * em;
    else w += 0.35 * em;
  }
  return w * 1.1;
}

// ---------- layouts ----------
function contentLayout(name, blockColor) {
  pres.defineSlideMaster({
    title: name,
    background: { color: WHITE },
    objects: [
      { rect: { x: 0, y: 0, w: BLOCK_W, h: 5.625, fill: { color: blockColor } } },
      {
        placeholder: {
          options: {
            name: "title", type: "title", x: BX, y: 0.3, w: BW, h: 1.2,
            fontSize: 36, bold: true, color: INKc, align: "left", valign: "bottom", margin: 0,
          },
          text: "",
        },
      },
    ],
    slideNumber: { x: 8.95, y: 5.28, w: 0.6, h: 0.25, fontSize: 9, color: ROSE, align: "right" },
  });
}
pres.defineSlideMaster({ title: "COVER", background: { color: BERRY } });
pres.defineSlideMaster({ title: "CLOSING", background: { color: BERRY } });
contentLayout("CONTENT_BERRY", BERRY);
contentLayout("CONTENT_CREAM", CREAM);

// ---------- helpers ----------
let shapeN = 0;
function T(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", align: "left", objectName: "text " + (++shapeN) }, o));
}
function R(slide, o) {
  slide.addShape(pres.shapes.RECTANGLE, Object.assign({ line: { color: WHITE, width: 0 }, objectName: "shape " + (++shapeN) }, o));
}
function RR(slide, o) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ rectRadius: 0.08, line: { color: WHITE, width: 0 }, objectName: "card " + (++shapeN) }, o));
}
function OVAL(slide, o) {
  slide.addShape(pres.shapes.OVAL, Object.assign({ line: { color: WHITE, width: 0 }, objectName: "circle " + (++shapeN) }, o));
}
function HLINE(slide, x, y, w, color, width = 0.75) {
  slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width }, objectName: "rule " + (++shapeN) });
}
function pills(slide, items, x, y, maxW, opts) {
  const { pt = 10, h = 0.26, gap = 0.1, fill, color, rowGap = 0.08 } = opts;
  let cx = x, cy = y;
  for (const t of items) {
    const w = estW(t, pt) + 0.3;
    if (cx + w > x + maxW + 0.01) { cx = x; cy += h + rowGap; }
    RR(slide, { x: cx, y: cy, w, h, fill: { color: fill }, rectRadius: h / 2 });
    T(slide, t, { x: cx, y: cy, w, h, fontSize: pt, color, align: "center", valign: "middle" });
    cx += w + gap;
  }
  return cy + h;
}
// The motif: left block with big number + small label; variant decides text colors
function block(slide, num, label, variant) {
  const big = variant === "berry" ? CREAMc : ROSEc;
  const lab = variant === "berry" ? CREAMc : BERRYc;
  T(slide, num, { x: 0.35, y: 0.28, w: 2.7, h: 1.15, fontSize: 72, bold: true, color: big, valign: "top" });
  T(slide, label, { x: 0.38, y: 1.42, w: 2.7, h: 0.25, fontSize: 10, bold: true, charSpacing: 3, color: lab });
}
function header(slide, text, x, y, w) {
  T(slide, text, { x, y, w, h: 0.22, fontSize: 10, bold: true, charSpacing: 2, color: ROSEc });
}
function bigStat(slide, value, caption, x, y, w, colors, sizes) {
  T(slide, value, { x, y, w, h: sizes.h, fontSize: sizes.pt, bold: true, color: colors.v, valign: "bottom" });
  T(slide, caption, { x, y: y + sizes.h + 0.03, w, h: 0.4, fontSize: sizes.cap || 11, color: colors.c, valign: "top" });
}

async function build() {
  // ============ 1. 표지 ============
  pres.addSection({ title: "표지" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
    T(s, "“", { x: 0.55, y: -0.15, w: 2.5, h: 2.2, fontSize: 200, bold: true, color: ROSEc, valign: "top" });
    T(s, "Hello, I'm", { x: 0.8, y: 1.75, w: 5, h: 0.4, fontSize: 18, color: CREAMc });
    T(s, "양남규", { x: 0.8, y: 2.1, w: 6, h: 1.1, fontSize: 64, bold: true, color: WHITEc, valign: "middle" });
    T(s, "Finance Operations Specialist", { x: 0.8, y: 3.25, w: 6.5, h: 0.45, fontSize: 22, bold: true, color: CREAMc });
    T(s, "스타트업의 성장을 뒷받침하는 경영지원 전문가", { x: 0.8, y: 3.75, w: 6.5, h: 0.4, fontSize: 16, color: WHITEc });
    T(s, "재무 · IR · HR · 컴플라이언스  /  GXC 경영지원팀  /  2026", { x: 0.8, y: 4.85, w: 6.5, h: 0.3, fontSize: 11, color: CREAMc, charSpacing: 1 });
    T(s, "2026", { x: 6.2, y: 3.55, w: 3.6, h: 1.6, fontSize: 96, bold: true, color: ROSEc, align: "right", valign: "bottom" });
  }

  // ============ 2. About Me ============
  pres.addSection({ title: "소개" });
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "소개" });
    block(s, "02", "ABOUT", "cream");
    s.addText("About Me", { placeholder: "title" });
    T(s, "“", { x: BX - 0.05, y: BY - 0.3, w: 1.2, h: 0.95, fontSize: 80, bold: true, color: CREAMc });
    T(s, "GXC에서 경영지원 업무를 담당하며, 재무/회계 관리부터 투자 관련 법률 검토, 인사/노무 행정, 정부사업 관리까지 폭넓은 영역을 수행하고 있습니다. 체계적인 학습과 실무 경험을 병행하며 스타트업의 성장에 필요한 오퍼레이션 역량을 꾸준히 쌓아가고 있습니다.",
      { x: BX, y: BY + 0.55, w: BW, h: 1.95, fontSize: 14, color: INKc, lineSpacingMultiple: 1.3 });
    header(s, "KEYWORDS", BX, 4.2, 3);
    pills(s, ["재무/회계", "IR/투자", "HR/노무", "정부사업", "Excel 고급", "데이터분석", "컴플라이언스"], BX, 4.47, BW, { pt: 11, h: 0.3, fill: BERRY, color: WHITEc });
    // big number in block
    T(s, "7", { x: 0.38, y: 2.0, w: 2.6, h: 1.0, fontSize: 60, bold: true, color: BERRYc, valign: "bottom" });
    T(s, "핵심 키워드", { x: 0.4, y: 3.03, w: 2.6, h: 0.3, fontSize: 11, color: INKc });
  }

  // ============ 3. 업무 영역 ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "소개" });
    block(s, "03", "SERVICE AREAS", "berry");
    s.addText("업무 영역", { placeholder: "title" });
    T(s, "4", { x: 0.38, y: 2.0, w: 2.6, h: 1.0, fontSize: 60, bold: true, color: CREAMc, valign: "bottom" });
    T(s, "개 업무 영역", { x: 0.4, y: 3.03, w: 2.6, h: 0.3, fontSize: 11, color: CREAMc });
    const cards = [
      ["재무 / 회계", "세무 관리, 월말 결산, 전표 처리, 재무제표 분석", Fa.FaCalculator],
      ["IR / 투자", "주주간계약(SHA), 스톡옵션 설계, TTC 설계, 투자 라운드", Fa.FaHandshake],
      ["HR / 노무", "인사 관리, 근로계약 관리, 노동법 이해, 급여 체계 설계", Fa.FaUsers],
      ["정부사업 / 기술", "정부 보조금, 컴플라이언스, Excel (고급), 데이터 분석", Fa.FaLandmark],
    ];
    const cw = 2.85, ch = 1.65, gx = 0.15, gy = 0.2;
    for (let i = 0; i < 4; i++) {
      const cx = BX + (i % 2) * (cw + gx), cy = BY + Math.floor(i / 2) * (ch + gy);
      RR(s, { x: cx, y: cy, w: cw, h: ch, fill: { color: CREAM } });
      OVAL(s, { x: cx + 0.22, y: cy + 0.22, w: 0.5, h: 0.5, fill: { color: BERRY } });
      s.addImage({ data: await icon(cards[i][2], WHITE), x: cx + 0.33, y: cy + 0.33, w: 0.28, h: 0.28, objectName: "icon " + cards[i][0] });
      T(s, String(i + 1).padStart(2, "0"), { x: cx + cw - 0.7, y: cy + 0.18, w: 0.5, h: 0.3, fontSize: 12, bold: true, color: ROSEc, align: "right" });
      T(s, cards[i][0], { x: cx + 0.85, y: cy + 0.22, w: 1.75, h: 0.5, fontSize: 16, bold: true, color: INKc, valign: "middle" });
      T(s, cards[i][1], { x: cx + 0.22, y: cy + 0.9, w: cw - 0.44, h: 0.65, fontSize: 11, color: INKc, lineSpacingMultiple: 1.15 });
    }
  }

  // ============ 4. Skills ============
  pres.addSection({ title: "역량" });
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "역량" });
    block(s, "04", "SKILLS", "cream");
    s.addText("Skills & Expertise", { placeholder: "title" });
    T(s, "85%", { x: 0.38, y: 2.0, w: 2.7, h: 1.0, fontSize: 54, bold: true, color: BERRYc, valign: "bottom" });
    T(s, "Excel (고급) · 최고 숙련도", { x: 0.4, y: 3.03, w: 2.6, h: 0.3, fontSize: 11, color: INKc });
    T(s, "16개 항목 · 단위 %", { x: 0.4, y: 3.33, w: 2.6, h: 0.3, fontSize: 10, color: ROSEc });
    const groups = [
      ["재무/회계", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], [75, 70, 80, 65]],
      ["IR/투자", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], [70, 65, 60, 60]],
      ["HR/노무", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], [75, 70, 65, 60]],
      ["정부사업/기술", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], [70, 65, 85, 60]],
    ];
    const cw = 2.85, ch = 1.75, gx = 0.15;
    for (let i = 0; i < 4; i++) {
      const cx = BX + (i % 2) * (cw + gx), cy = BY + Math.floor(i / 2) * ch;
      header(s, groups[i][0], cx, cy, cw);
      s.addChart(pres.charts.BAR, [{ name: groups[i][0], labels: groups[i][1], values: groups[i][2] }], {
        x: cx - 0.05, y: cy + 0.2, w: cw + 0.05, h: ch - 0.25,
        barDir: "bar", barGapWidthPct: 55,
        chartColors: [i % 2 === 0 ? BERRY : ROSE],
        valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100,
        valGridLine: { style: "none" }, catGridLine: { style: "none" },
        catAxisLineShow: false, valAxisLineShow: false,
        catAxisLabelColor: INK, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt",
        catAxisOrientation: "maxMin",
        showValue: true, dataLabelPosition: "outEnd", dataLabelColor: INK, dataLabelFontSize: 10, dataLabelFontFace: "+mn-lt",
        dataLabelFormatCode: '0"%"',
        showLegend: false, showTitle: false,
        objectName: "chart " + groups[i][0],
      });
    }
  }

  // ============ 5. Excel Mastery ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "역량" });
    block(s, "05", "EXCEL", "berry");
    s.addText("Excel Mastery", { placeholder: "title" });
    T(s, "실무에서 활용하는 엑셀 고급 함수와 자동화 기법", { x: 0.4, y: 1.95, w: 2.6, h: 0.9, fontSize: 14, bold: true, color: WHITEc, lineSpacingMultiple: 1.25 });
    T(s, "4", { x: 0.38, y: 3.1, w: 2.6, h: 1.0, fontSize: 60, bold: true, color: CREAMc, valign: "bottom" });
    T(s, "가지 핵심 기법", { x: 0.4, y: 4.13, w: 2.6, h: 0.3, fontSize: 11, color: CREAMc });
    const rows = [
      ["XLOOKUP", "급여 대장 자동 매칭 / 거래처별 매출 추적 / 계정과목 자동 분류", ["VLOOKUP 대체", "양방향 검색", "에러 처리 내장"]],
      ["피벗테이블", "부서별 예산 집행 분석 / 월별 비용 추이 리포트", ["크로스탭 분석", "슬라이서 필터", "자동 갱신"]],
      ["조건부 서식", "예산 초과 자동 알림 (달성률 100% 초과 빨강, 90% 미만 노랑)", ["데이터 막대", "아이콘 세트", "색조 규칙"]],
      ["자동화 매크로(VBA)", "월말 결산 자동화 / 급여명세서 자동 발송", ["VBA 매크로", "반복 업무 제거", "3시간→15분"]],
    ];
    const pitch = 0.875;
    for (let i = 0; i < 4; i++) {
      const y = BY + i * pitch;
      T(s, String(i + 1).padStart(2, "0"), { x: BX, y: y + 0.02, w: 0.55, h: 0.4, fontSize: 22, bold: true, color: ROSEc });
      T(s, rows[i][0], { x: BX + 0.6, y: y + 0.04, w: 1.75, h: 0.55, fontSize: 14, bold: true, color: INKc });
      T(s, rows[i][1], { x: BX + 2.4, y: y + 0.05, w: BW - 2.4, h: 0.42, fontSize: 11, color: INKc, lineSpacingMultiple: 1.1 });
      pills(s, rows[i][2], BX + 0.6, y + 0.52, BW - 0.6, { pt: 10, h: 0.24, gap: 0.08, fill: CREAM, color: BERRYc });
      if (i < 3) HLINE(s, BX, y + pitch - 0.06, BW, CREAM, 1);
    }
  }

  // ============ 6. 자동화 성과 ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "역량" });
    block(s, "06", "AUTOMATION", "cream");
    s.addText("자동화 성과", { placeholder: "title" });
    T(s, "3단계", { x: 0.38, y: 1.95, w: 2.6, h: 0.8, fontSize: 44, bold: true, color: BERRYc, valign: "bottom" });
    const steps = ["전표 정리", "계정과목별 집계", "보고서 출력(PDF 저장·메일 발송)"];
    for (let i = 0; i < 3; i++) {
      const y = 2.95 + i * 0.62;
      OVAL(s, { x: 0.4, y: y + 0.02, w: 0.34, h: 0.34, fill: { color: BERRY } });
      T(s, String(i + 1), { x: 0.4, y: y + 0.02, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: WHITEc, align: "center", valign: "middle" });
      T(s, steps[i], { x: 0.88, y, w: 2.2, h: 0.5, fontSize: 11, color: INKc, valign: "top", lineSpacingMultiple: 1.1 });
    }
    // big stats
    bigStat(s, "3시간 → 15분", "월말 결산 보고서 생성 (전표 정리 → 집계 → 보고서, 버튼 하나)", BX, BY, 2.75, { v: BERRYc, c: INKc }, { pt: 26, h: 0.55 });
    bigStat(s, "3일 → 1.2일", "월말 결산 전체 소요 시간, 약 60% 단축", BX + 3.1, BY, 2.75, { v: BERRYc, c: INKc }, { pt: 26, h: 0.55 });
    // pivot table
    header(s, "피벗 집계 예시 — 부서 / Q1 / Q2 / 증감률", BX, 3.2, BW);
    const hdr = (t, al) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: BERRY }, align: al, fontSize: 10 } });
    const cell = (t, al, bold, color) => ({ text: t, options: { align: al, fontSize: 10, color: color || INK, bold: !!bold } });
    const tbl = [
      [hdr("부서", "left"), hdr("Q1", "right"), hdr("Q2", "right"), hdr("증감률", "right")],
      [cell("개발팀", "left"), cell("45,200", "right"), cell("52,800", "right"), cell("+16.8%", "right", true, BERRY)],
      [cell("마케팅", "left"), cell("32,100", "right"), cell("28,500", "right"), cell("-11.2%", "right", true, ROSE)],
      [cell("경영지원", "left"), cell("18,300", "right"), cell("19,100", "right"), cell("+4.4%", "right", true, BERRY)],
      [cell("합계", "left", true), cell("95,600", "right", true), cell("100,400", "right", true), cell("+5.0%", "right", true, BERRY)],
    ];
    s.addTable(tbl, {
      x: BX, y: 3.48, w: BW, colW: [1.95, 1.3, 1.3, 1.3], rowH: 0.33,
      border: { type: "solid", color: CREAM, pt: 1 }, margin: [0.03, 0.1, 0.03, 0.1], valign: "middle",
      objectName: "pivot table",
    });
  }

  // ============ 7-9. Cases ============
  pres.addSection({ title: "케이스" });
  const cases = [
    {
      n: "07", variant: "berry", label: "CASE 01 · IR/투자", title: "SHA 주주간계약 검토",
      bg: "GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.",
      steps: ["ROFR(우선매수권) 조항의 실무 적용 시나리오 분석", "Drag-Along · Tag-Along 권리의 행사 조건 정리", "희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)", "투자자 동의권(Consent Rights) 범위 및 영향 분석", "법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화"],
      result: "5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.",
      metrics: [["5개", "핵심 조항 분석"], ["1건", "가이드 문서"]],
      lesson: "법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다.",
    },
    {
      n: "08", variant: "cream", label: "CASE 02 · 재무/회계", title: "월말 결산 프로세스 혁신",
      bg: "매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.",
      steps: ["기존 결산 프로세스 전수조사 및 병목 구간 파악", "전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)", "계정과목별 자동 집계 템플릿 구축", "VBA 매크로를 활용한 보고서 자동 생성", "결산 체크리스트 및 SOP 문서 작성"],
      result: "결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.",
      metrics: [["-60%", "소요 시간"], ["SOP", "프로세스 표준화"]],
      lesson: "Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다.",
    },
    {
      n: "09", variant: "berry", label: "CASE 03 · HR/노무", title: "TTC 기반 보상 체계 설계",
      bg: "스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.",
      steps: ["국내 IT 스타트업 보상 시장조사 및 벤치마킹", "TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)", "직급/직무별 Pay Band 설계", "스톡옵션 배분 기준 및 베스팅 스케줄 수립", "경영진 보고 및 의사결정 자료 준비"],
      result: "전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.",
      metrics: [["전사", "적용 범위"], ["TTC", "체계 도입"]],
      lesson: "보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다.",
    },
  ];
  for (const cs of cases) {
    const s = pres.addSlide({ masterName: cs.variant === "berry" ? "CONTENT_BERRY" : "CONTENT_CREAM", sectionTitle: "케이스" });
    block(s, cs.n, cs.label, cs.variant);
    s.addText(cs.title, { placeholder: "title" });
    const v = cs.variant === "berry" ? CREAMc : BERRYc;
    const c = cs.variant === "berry" ? WHITEc : INKc;
    // metrics in block
    for (let i = 0; i < 2; i++) {
      const x = 0.4 + i * 1.35;
      T(s, cs.metrics[i][0], { x, y: 1.95, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: v, valign: "bottom" });
      T(s, cs.metrics[i][1], { x, y: 2.57, w: 1.3, h: 0.3, fontSize: 10, color: c });
    }
    // lesson quote
    T(s, "“", { x: 0.32, y: 3.05, w: 0.6, h: 0.6, fontSize: 48, bold: true, color: ROSEc });
    T(s, "배운 점", { x: 0.95, y: 3.3, w: 1.6, h: 0.25, fontSize: 10, bold: true, charSpacing: 2, color: v });
    T(s, cs.lesson, { x: 0.4, y: 3.65, w: 2.6, h: 1.5, fontSize: 11, color: c, lineSpacingMultiple: 1.2 });
    // body: two columns
    const cA = 2.75, cB = BW - cA - 0.3, xB = BX + cA + 0.3;
    header(s, "배경", BX, BY, cA);
    T(s, cs.bg, { x: BX, y: BY + 0.27, w: cA, h: 1.55, fontSize: 11, color: INKc, lineSpacingMultiple: 1.2 });
    header(s, "결과", BX, 3.6, cA);
    T(s, cs.result, { x: BX, y: 3.87, w: cA, h: 1.25, fontSize: 11, bold: true, color: BERRYc, lineSpacingMultiple: 1.2 });
    header(s, "과정", xB, BY, cB);
    s.addText(cs.steps.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < cs.steps.length - 1 } })),
      { x: xB, y: BY + 0.27, w: cB, h: 3.2, fontSize: 11, color: INKc, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 7, lineSpacingMultiple: 1.15, objectName: "process list " + cs.n });
  }

  // ============ 10. Career Timeline ============
  pres.addSection({ title: "경력" });
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "경력" });
    block(s, "10", "TIMELINE", "cream");
    s.addText("Career Timeline", { placeholder: "title" });
    T(s, "2026", { x: 0.38, y: 2.0, w: 2.6, h: 1.0, fontSize: 60, bold: true, color: BERRYc, valign: "bottom" });
    T(s, "GXC 경영지원팀 합류", { x: 0.4, y: 3.03, w: 2.6, h: 0.3, fontSize: 11, color: INKc });
    T(s, "위가 최근 · 아래가 과거", { x: 0.4, y: 3.33, w: 2.6, h: 0.3, fontSize: 10, color: ROSEc });
    const items = [
      ["2026.04", "스톡옵션 세무 체계 정립", "스톡옵션 법률 구조 및 과세 체계 분석 완료. 임직원 보상 설계의 세무 관점 가이드라인 수립."],
      ["2026.04", "ABTC 인증 처리 완료", "APEC 기업인 여행카드(ABTC) 관련 문서 작업 및 인증 프로세스 완료."],
      ["2026.04", "SHA(주주간계약) 분석", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 등 주주간계약의 핵심 조항 분석 및 실무 적용."],
      ["2026.04", "IT 스타트업 런웨이 관리 체계 수립", "현금흐름 기반 런웨이 관리 프레임워크 학습 및 GXC 적용 방안 도출."],
      ["2026 ~", "GXC 경영지원팀 합류", "재무/회계, IR, HR, 정부사업 등 경영지원 전반 업무 담당. 스타트업 운영 기반 구축."],
    ];
    const pitch = 0.72, lineX = 4.78;
    s.addShape(pres.shapes.LINE, { x: lineX, y: BY + 0.1, w: 0, h: pitch * 4 + 0.05, line: { color: CREAM, width: 1.5 }, objectName: "timeline" });
    for (let i = 0; i < items.length; i++) {
      const y = BY + i * pitch;
      T(s, items[i][0], { x: BX, y: y + 0.02, w: 0.85, h: 0.25, fontSize: 11, bold: true, color: BERRYc });
      OVAL(s, { x: lineX - 0.08, y: y + 0.07, w: 0.16, h: 0.16, fill: { color: i === items.length - 1 ? ROSE : BERRY } });
      T(s, items[i][1], { x: 5.05, y: y - 0.01, w: 4.5, h: 0.26, fontSize: 13, bold: true, color: INKc });
      T(s, items[i][2], { x: 5.05, y: y + 0.26, w: 4.5, h: 0.38, fontSize: 10, color: INKc, lineSpacingMultiple: 1.05 });
    }
  }

  // ============ 11. Certifications & Learning ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "경력" });
    block(s, "11", "LEARNING", "berry");
    s.addText("Certifications & Learning", { placeholder: "title" });
    T(s, "2026.06", { x: 0.38, y: 2.0, w: 2.7, h: 1.0, fontSize: 44, bold: true, color: CREAMc, valign: "bottom" });
    T(s, "자격증 취득 목표", { x: 0.4, y: 3.03, w: 2.6, h: 0.3, fontSize: 11, color: CREAMc });
    const cA = 2.7, xB = BX + cA + 0.3, cB = BW - cA - 0.3;
    header(s, "자격증(준비중)", BX, BY, cA);
    const certs = [["전산회계 1급", "한국세무사회", "2026.06 목표"], ["전산세무 2급", "한국세무사회", "2026.06 목표"]];
    const certIcon = await icon(Fa.FaCertificate, WHITE);
    for (let i = 0; i < 2; i++) {
      const y = BY + 0.3 + i * 1.45;
      RR(s, { x: BX, y, w: cA, h: 1.3, fill: { color: CREAM } });
      OVAL(s, { x: BX + 0.2, y: y + 0.22, w: 0.46, h: 0.46, fill: { color: BERRY } });
      s.addImage({ data: certIcon, x: BX + 0.31, y: y + 0.33, w: 0.24, h: 0.24, objectName: "cert icon " + (i + 1) });
      T(s, certs[i][0], { x: BX + 0.8, y: y + 0.2, w: 1.8, h: 0.5, fontSize: 15, bold: true, color: INKc, valign: "middle" });
      T(s, certs[i][1], { x: BX + 0.2, y: y + 0.82, w: 1.3, h: 0.3, fontSize: 11, color: INKc });
      T(s, certs[i][2], { x: BX + 1.4, y: y + 0.82, w: 1.1, h: 0.3, fontSize: 11, bold: true, color: BERRYc, align: "right" });
    }
    header(s, "최근 학습", xB, BY, cB);
    const learn = [["2026.04.16", "스톡옵션 과세 체계", "IR/투자"], ["2026.04.14", "스톡옵션 법률 구조", "IR/투자"], ["2026.04.10", "주주간계약(SHA) 핵심 조항", "IR/투자"], ["2026.04.08", "IT 스타트업 런웨이 관리", "재무"]];
    for (let i = 0; i < 4; i++) {
      const y = BY + 0.32 + i * 0.78;
      T(s, [{ text: learn[i][0], options: { bold: true, color: BERRYc } }, { text: "   " + learn[i][2], options: { color: ROSEc } }], { x: xB, y, w: cB, h: 0.22, fontSize: 10, color: INKc });
      T(s, learn[i][1], { x: xB, y: y + 0.24, w: cB, h: 0.3, fontSize: 13, bold: true, color: INKc });
      if (i < 3) HLINE(s, xB, y + 0.68, cB, CREAM, 1);
    }
  }

  // ============ 12. Project Gallery ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "경력" });
    block(s, "12", "PROJECTS", "cream");
    s.addText("Project Gallery", { placeholder: "title" });
    T(s, "6", { x: 0.38, y: 2.0, w: 2.6, h: 1.0, fontSize: 60, bold: true, color: BERRYc, valign: "bottom" });
    T(s, "개 프로젝트", { x: 0.4, y: 3.03, w: 2.6, h: 0.3, fontSize: 11, color: INKc });
    T(s, "IR/투자 · HR/노무 · 정부사업 · 재무/회계", { x: 0.4, y: 3.33, w: 2.6, h: 0.5, fontSize: 10, color: ROSEc });
    const projects = [
      ["IR/투자", "SHA 주주간계약 검토", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 분석 및 투자 계약 실무 적용", "투자 계약 핵심 조항 가이드 문서화"],
      ["IR/투자", "스톡옵션 체계 설계", "법률 구조 분석, 과세 체계 정립, 임직원 보상 설계 가이드라인 수립", "스톡옵션 세무/법률 통합 가이드 완성"],
      ["HR/노무", "급여 체계 설계", "TTC(Total Target Compensation) 기반 급여 체계 설계 및 보상 구조 수립", "전사 보상 체계 프레임워크 구축"],
      ["정부사업", "ABTC 인증 처리", "APEC 기업인 여행카드 관련 서류 준비 및 인증 프로세스 관리", "인증 완료 및 프로세스 SOP 문서화"],
      ["재무/회계", "런웨이 관리 체계", "IT 스타트업 현금흐름 기반 런웨이 관리 프레임워크 구축", "월별 런웨이 모니터링 대시보드 구축"],
      ["재무/회계", "월말 결산 프로세스", "매월 재무제표 마감 및 결산 프로세스 체계화, 전표 처리 자동화", "결산 소요 시간 단축 및 프로세스 표준화"],
    ];
    const cw = 2.85, ch = 1.12, gx = 0.15, gy = 0.07;
    for (let i = 0; i < 6; i++) {
      const cx = BX + (i % 2) * (cw + gx), cy = BY + Math.floor(i / 2) * (ch + gy);
      RR(s, { x: cx, y: cy, w: cw, h: ch, fill: { color: WHITE }, line: { color: CREAM, width: 1.25 } });
      T(s, projects[i][0], { x: cx + cw - 1.15, y: cy + 0.12, w: 1.0, h: 0.24, fontSize: 10, bold: true, color: ROSEc, align: "right" });
      T(s, projects[i][1], { x: cx + 0.15, y: cy + 0.1, w: 1.8, h: 0.26, fontSize: 12, bold: true, color: INKc });
      T(s, [{ text: projects[i][2], options: { color: INKc } }, { text: "  →  " + projects[i][3], options: { bold: true, color: BERRYc } }],
        { x: cx + 0.15, y: cy + 0.4, w: cw - 0.3, h: 0.68, fontSize: 10, lineSpacingMultiple: 1.15 });
    }
  }

  // ============ 13. Let's Connect ============
  pres.addSection({ title: "마무리" });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "마무리" });
    T(s, "“", { x: 0.55, y: -0.15, w: 2.5, h: 2.2, fontSize: 200, bold: true, color: ROSEc });
    T(s, "Let's Connect", { x: 0.8, y: 1.85, w: 8, h: 1.0, fontSize: 56, bold: true, color: WHITEc, valign: "middle" });
    T(s, "새로운 기회나 협업에 대해 이야기하고 싶으시다면 연락주세요.", { x: 0.8, y: 2.95, w: 8, h: 0.4, fontSize: 16, color: CREAMc });
    const contacts = [["Email", Fa.FaEnvelope], ["LinkedIn", Fa.FaLinkedinIn]];
    for (let i = 0; i < 2; i++) {
      const x = 0.8 + i * 2.1;
      OVAL(s, { x, y: 3.75, w: 0.5, h: 0.5, fill: { color: CREAM } });
      s.addImage({ data: await icon(contacts[i][1], BERRY), x: x + 0.12, y: 3.87, w: 0.26, h: 0.26, objectName: "icon " + contacts[i][0] });
      T(s, contacts[i][0], { x: x + 0.62, y: 3.75, w: 1.4, h: 0.5, fontSize: 15, bold: true, color: WHITEc, valign: "middle" });
    }
    T(s, "© 2026 양남규", { x: 0.8, y: 4.95, w: 5, h: 0.3, fontSize: 10, color: CREAMc });
    T(s, "13", { x: 6.2, y: 3.55, w: 3.6, h: 1.6, fontSize: 96, bold: true, color: ROSEc, align: "right", valign: "bottom" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

build().catch((e) => { console.error(e); process.exit(1); });
