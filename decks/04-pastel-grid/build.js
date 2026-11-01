// Pastel Grid — 양남규 포트폴리오 덱 (13장)
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

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.author = "양남규";
  pres.title = "양남규 포트폴리오 — Pastel Grid";
  const C = pres.SchemeColor;

  // Pre-render icons (dark gray on pastel badges; sage/white variants for cover)
  const names = [
    "FiUser", "FiDollarSign", "FiTrendingUp", "FiUsers", "FiShield", "FiBriefcase",
    "FiSearch", "FiGrid", "FiAlertCircle", "FiZap", "FiClock", "FiCalendar",
    "FiFileText", "FiAward", "FiBookOpen", "FiMail", "FiLinkedin", "FiCheckCircle",
    "FiTag", "FiBarChart2", "FiLayers", "FiTarget", "FiFlag", "FiPieChart",
    "FiCreditCard", "FiActivity", "FiRepeat", "FiHash",
  ];
  const ICON = {};
  for (const n of names) {
    ICON[n] = await iconPng(Fi[n], H.dk1);
  }
  const ICON_SAGE = {};
  for (const n of ["FiDollarSign", "FiTrendingUp", "FiUsers", "FiShield", "FiMail", "FiLinkedin", "FiUser"]) {
    ICON_SAGE[n] = await iconPng(Fi[n], H.accent1);
  }

  // ---------- Layouts ----------
  const TITLE_Y = 0.3, TITLE_H = 0.65;
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
            name: "title", type: "title", x: 0.5, y: TITLE_Y, w: 6.6, h: TITLE_H,
            fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0,
          },
          text: "",
        },
      },
      {
        text: {
          text: "양남규 · Finance Operations Specialist",
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

  const PASTEL = [C.accent1, C.accent2, C.accent3];
  const TINT = [C.accent4, C.accent5, C.accent6];
  const TINT_HEX = [H.accent4, H.accent5, H.accent6];

  // =============== 1. 표지 ===============
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    // left big white tile
    tile(s, 0.5, 0.55, 5.35, 4.5, C.background1, { r: 0.2 });
    text(s, "Hello, I'm", { x: 0.9, y: 0.95, w: 4.6, h: 0.4, fontSize: 16, color: C.text2 });
    text(s, "양남규", { x: 0.9, y: 1.35, w: 4.6, h: 1.0, fontSize: 54, bold: true, color: C.text1 });
    text(s, "Finance Operations Specialist", { x: 0.9, y: 2.4, w: 4.7, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
    text(s, "스타트업의 성장을 뒷받침하는 경영지원 전문가", { x: 0.9, y: 2.9, w: 4.7, h: 0.4, fontSize: 15, color: C.text1 });
    // sub chips
    const subs = ["재무 · IR · HR · 컴플라이언스", "GXC 경영지원팀", "2026"];
    let cx = 0.9;
    const cw = [2.55, 1.35, 0.7];
    subs.forEach((t, i) => {
      chip(s, t, cx, 4.25, cw[i], 0.36, TINT[i], C.text1, 10);
      cx += cw[i] + 0.12;
    });
    // right 2x2 white tiles with badges
    const grid = [
      ["FiDollarSign", "재무"], ["FiTrendingUp", "IR"], ["FiUsers", "HR"], ["FiShield", "컴플라이언스"],
    ];
    const gx = 6.15, gy = 0.55, tw = 1.6, th = 2.15, gap = 0.15;
    grid.forEach(([ic, lab], i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      tile(s, x, y, tw, th, C.background1, { r: 0.2 });
      badge(s, x + (tw - 0.8) / 2, y + 0.45, 0.8, TINT[i % 3], ic);
      text(s, lab, { x: x + 0.1, y: y + 1.4, w: tw - 0.2, h: 0.4, fontSize: 13, bold: true, color: C.text1, align: "center" });
    });
    s.addNotes("표지 — 양남규, Finance Operations Specialist");
  }

  // =============== 2. About Me ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intro" });
    s.addText("About Me", { placeholder: "title" });
    tile(s, 0.5, 1.1, 5.55, 4.05, C.accent4);
    badge(s, 0.85, 1.4, 0.7, C.accent1, "FiUser");
    text(s, "GXC 경영지원팀", { x: 1.7, y: 1.45, w: 3.5, h: 0.3, fontSize: 14, bold: true, color: C.text1 });
    text(s, "Finance Operations Specialist", { x: 1.7, y: 1.75, w: 3.8, h: 0.3, fontSize: 11, color: C.text2 });
    text(
      s,
      "GXC에서 경영지원 업무를 담당하며, 재무/회계 관리부터 투자 관련 법률 검토, 인사/노무 행정, 정부사업 관리까지 폭넓은 영역을 수행하고 있습니다. 체계적인 학습과 실무 경험을 병행하며 스타트업의 성장에 필요한 오퍼레이션 역량을 꾸준히 쌓아가고 있습니다.",
      { x: 0.85, y: 2.35, w: 4.85, h: 2.6, fontSize: 16, color: C.text1, valign: "top", lineSpacingMultiple: 1.35 }
    );
    // right: tag tiles 2x4
    const tags = ["재무/회계", "IR/투자", "HR/노무", "정부사업", "Excel 고급", "데이터분석", "컴플라이언스"];
    const gx = 6.3, gy = 1.1, tw = 1.525, th = 0.9, gap = 0.15;
    tags.forEach((t, i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      tile(s, x, y, tw, th, TINT[i % 3]);
      text(s, t, { x: x + 0.1, y, w: tw - 0.2, h: th, fontSize: 13, bold: true, color: C.text1, align: "center", valign: "middle" });
    });
    // 8th cell: badge
    {
      const x = gx + tw + gap, y = gy + 3 * (th + gap);
      tile(s, x, y, tw, th, C.text1);
      badge(s, x + (tw - 0.5) / 2, y + 0.2, 0.5, C.accent3, "FiTag");
    }
    s.addNotes("About Me");
  }

  // =============== 3. 업무 영역 ===============
  pres.addSection({ title: "Work" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Work" });
    s.addText("업무 영역", { placeholder: "title" });
    const items = [
      ["FiDollarSign", "재무 / 회계", "세무 관리, 월말 결산, 전표 처리, 재무제표 분석"],
      ["FiTrendingUp", "IR / 투자", "주주간계약(SHA), 스톡옵션 설계, TTC 설계, 투자 라운드"],
      ["FiUsers", "HR / 노무", "인사 관리, 근로계약 관리, 노동법 이해, 급여 체계 설계"],
      ["FiShield", "정부사업 / 기술", "정부 보조금, 컴플라이언스, Excel (고급), 데이터 분석"],
    ];
    const gx = 0.5, gy = 1.1, tw = 4.4, th = 1.925, gap = 0.2;
    items.forEach(([ic, h, d], i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      tile(s, x, y, tw, th, TINT[i % 3]);
      badge(s, x + 0.3, y + 0.3, 0.75, PASTEL[i % 3], ic);
      text(s, h, { x: x + 1.25, y: y + 0.32, w: tw - 1.5, h: 0.4, fontSize: 20, bold: true, color: C.text1 });
      text(s, d, { x: x + 1.25, y: y + 0.82, w: tw - 1.5, h: 0.9, fontSize: 14, color: C.text2, valign: "top", lineSpacingMultiple: 1.2 });
    });
    s.addNotes("업무 영역 4개");
  }

  // =============== 4. Skills & Expertise ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Work" });
    s.addText("Skills & Expertise", { placeholder: "title" });
    text(s, "숙련도 (단위: %)", { x: 7.2, y: TITLE_Y, w: 2.3, h: TITLE_H, fontSize: 11, color: C.text2, align: "right", valign: "middle" });
    const groups = [
      ["FiDollarSign", "재무/회계", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], [75, 70, 80, 65], H.accent1],
      ["FiTrendingUp", "IR/투자", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], [70, 65, 60, 60], H.accent2],
      ["FiUsers", "HR/노무", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], [75, 70, 65, 60], H.accent3],
      ["FiShield", "정부사업/기술", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], [70, 65, 85, 60], H.dk2],
    ];
    const gx = 0.5, gy = 1.1, tw = 4.4, th = 1.925, gap = 0.2;
    groups.forEach(([ic, h, labels, vals, barHex], i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      const tintHex = TINT_HEX[i % 3];
      tile(s, x, y, tw, th, TINT[i % 3]);
      badge(s, x + 0.25, y + 0.2, 0.45, PASTEL[i % 3], ic);
      text(s, h, { x: x + 0.8, y: y + 0.2, w: 2.5, h: 0.45, fontSize: 15, bold: true, color: C.text1, valign: "middle" });
      s.addChart(
        pres.charts.BAR,
        [{ name: h, labels, values: vals }],
        {
          x: x + 0.15, y: y + 0.68, w: tw - 0.3, h: th - 0.78,
          barDir: "bar", barGapWidthPct: 55,
          chartColors: [barHex],
          chartArea: { fill: { color: tintHex } },
          plotArea: { fill: { color: tintHex } },
          catAxisOrientation: "maxMin",
          catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 10, catAxisLabelColor: H.dk1,
          catAxisLineShow: false, catGridLine: { style: "none" },
          valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: "none" }, valAxisLineShow: false,
          valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: 10, valAxisLabelColor: H.dk2,
          showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
          dataLabelFontFace: "+mn-lt", dataLabelFontSize: 10, dataLabelColor: H.dk1,
          showLegend: false, showTitle: false,
          objectName: nm("Chart"), altText: `${h} 숙련도 가로 막대 차트`,
        }
      );
    });
    s.addNotes("Skills & Expertise — 16개 숙련도, 최고치 Excel 85%");
  }

  // =============== 5. Excel Mastery ===============
  pres.addSection({ title: "Excel" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Excel" });
    s.addText("Excel Mastery", { placeholder: "title" });
    text(s, "실무에서 활용하는 엑셀 고급 함수와 자동화 기법", { x: 0.5, y: 1.0, w: 8, h: 0.3, fontSize: 13, color: C.text2 });
    const items = [
      ["FiSearch", "XLOOKUP", "급여 대장 자동 매칭 / 거래처별 매출 추적 / 계정과목 자동 분류", ["VLOOKUP 대체", "양방향 검색", "에러 처리 내장"]],
      ["FiGrid", "피벗테이블", "부서별 예산 집행 분석 / 월별 비용 추이 리포트", ["크로스탭 분석", "슬라이서 필터", "자동 갱신"]],
      ["FiAlertCircle", "조건부 서식", "예산 초과 자동 알림 (달성률 100% 초과 빨강, 90% 미만 노랑)", ["데이터 막대", "아이콘 세트", "색조 규칙"]],
      ["FiZap", "자동화 매크로(VBA)", "월말 결산 자동화 / 급여명세서 자동 발송", ["VBA 매크로", "반복 업무 제거", "3시간→15분"]],
    ];
    const gx = 0.5, gy = 1.45, tw = 4.4, th = 1.775, gap = 0.2;
    items.forEach(([ic, h, d, chips], i) => {
      const x = gx + (i % 2) * (tw + gap), y = gy + Math.floor(i / 2) * (th + gap);
      tile(s, x, y, tw, th, TINT[i % 3]);
      badge(s, x + 0.25, y + 0.18, 0.5, PASTEL[i % 3], ic);
      text(s, h, { x: x + 0.9, y: y + 0.18, w: tw - 1.15, h: 0.5, fontSize: 17, bold: true, color: C.text1, valign: "middle" });
      text(s, d, { x: x + 0.25, y: y + 0.72, w: tw - 0.5, h: 0.55, fontSize: 13, color: C.text2, valign: "top", lineSpacingMultiple: 1.12 });
      let cx = x + 0.25;
      chips.forEach((c) => {
        const w = Math.max(0.9, c.length * 0.1 + 0.3);
        chip(s, c, cx, y + th - 0.42, w, 0.28, C.background1, C.text1, 9);
        cx += w + 0.1;
      });
    });
    s.addNotes("Excel Mastery 4기법");
  }

  // =============== 6. 자동화 성과 ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Excel" });
    s.addText("자동화 성과", { placeholder: "title" });
    // two stat tiles
    const stats = [
      ["FiClock", "3시간 → 15분", "월말 결산 보고서 생성", "전표 정리 → 집계 → 보고서, 버튼 하나"],
      ["FiCalendar", "3일 → 1.2일", "월말 결산 전체 소요 시간", "약 60% 단축"],
    ];
    stats.forEach(([ic, big, lab, cap], i) => {
      const x = 0.5 + i * 4.6, y = 1.1, w = 4.4, h = 1.55;
      tile(s, x, y, w, h, TINT[i]);
      badge(s, x + w - 0.85, y + 0.25, 0.6, PASTEL[i], ic);
      text(s, big, { x: x + 0.3, y: y + 0.2, w: w - 1.3, h: 0.6, fontSize: 30, bold: true, color: C.text1, valign: "middle" });
      text(s, lab, { x: x + 0.3, y: y + 0.85, w: w - 0.6, h: 0.3, fontSize: 14, bold: true, color: C.text1 });
      text(s, cap, { x: x + 0.3, y: y + 1.15, w: w - 0.6, h: 0.3, fontSize: 11, color: C.text2 });
    });
    // table tile
    {
      const x = 0.5, y = 2.85, w = 5.6, h = 2.3;
      tile(s, x, y, w, h, C.accent6);
      label(s, "예시 피벗 표", x + 0.3, y + 0.2, 3);
      const hdr = (t, al) => ({ text: t, options: { bold: true, color: H.lt1, fill: { color: H.dk1 }, align: al, fontSize: 11 } });
      const cell = (t, al, bold, color) => ({ text: t, options: { align: al, fontSize: 11, bold: !!bold, color: color || H.dk1, fill: { color: H.lt1 } } });
      const rows = [
        [hdr("부서", "left"), hdr("Q1", "right"), hdr("Q2", "right"), hdr("증감률", "right")],
        [cell("개발팀", "left"), cell("45,200", "right"), cell("52,800", "right"), cell("+16.8%", "right", true, H.dk2)],
        [cell("마케팅", "left"), cell("32,100", "right"), cell("28,500", "right"), cell("-11.2%", "right", true, H.dk2)],
        [cell("경영지원", "left"), cell("18,300", "right"), cell("19,100", "right"), cell("+4.4%", "right", true, H.dk2)],
        [cell("합계", "left", true), cell("95,600", "right", true), cell("100,400", "right", true), cell("+5.0%", "right", true)],
      ];
      s.addTable(rows, {
        x: x + 0.3, y: y + 0.5, w: w - 0.6, colW: [1.4, 1.2, 1.2, 1.2], rowH: 0.3,
        border: { type: "solid", color: H.accent6, pt: 1 }, margin: 0.06, valign: "middle",
        objectName: nm("Table"),
      });
    }
    // 3단계 tile
    {
      const x = 6.35, y = 2.85, w = 3.15, h = 2.3;
      tile(s, x, y, w, h, C.accent4);
      text(s, "3단계", { x: x + 0.3, y: y + 0.15, w: 2.0, h: 0.5, fontSize: 26, bold: true, color: C.text1, valign: "middle" });
      badge(s, x + w - 0.75, y + 0.2, 0.5, C.accent1, "FiLayers");
      const steps = ["전표 정리", "계정과목별 집계", "보고서 출력 (PDF 저장·메일 발송)"];
      steps.forEach((t, i) => {
        const yy = y + 0.78 + i * 0.47;
        s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: yy, w: 0.34, h: 0.34, fill: { color: C.background1 }, line: { color: C.background1, width: 0 }, objectName: nm("Step") });
        text(s, String(i + 1), { x: x + 0.3, y: yy, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: C.text1, align: "center", valign: "middle" });
        text(s, t, { x: x + 0.75, y: yy - 0.03, w: w - 1.0, h: 0.4, fontSize: 12, color: C.text1, valign: "middle" });
      });
    }
    s.addNotes("자동화 성과 — 3시간→15분, 3일→1.2일, 3단계");
  }

  // =============== 7-9. Cases ===============
  pres.addSection({ title: "Cases" });
  const cases = [
    {
      no: "Case 01", cat: "IR/투자", title: "SHA 주주간계약 검토", tint: 1,
      bg: "GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.",
      steps: ["ROFR(우선매수권) 조항의 실무 적용 시나리오 분석", "Drag-Along · Tag-Along 권리의 행사 조건 정리", "희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)", "투자자 동의권(Consent Rights) 범위 및 영향 분석", "법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화"],
      result: "5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.",
      kpis: [["FiFileText", "5개", "핵심 조항 분석"], ["FiBookOpen", "1건", "가이드 문서"]],
      learn: "법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다.",
    },
    {
      no: "Case 02", cat: "재무/회계", title: "월말 결산 프로세스 혁신", tint: 0,
      bg: "매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.",
      steps: ["기존 결산 프로세스 전수조사 및 병목 구간 파악", "전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)", "계정과목별 자동 집계 템플릿 구축", "VBA 매크로를 활용한 보고서 자동 생성", "결산 체크리스트 및 SOP 문서 작성"],
      result: "결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.",
      kpis: [["FiClock", "-60%", "소요 시간"], ["FiCheckCircle", "SOP", "프로세스 표준화"]],
      learn: "Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다.",
    },
    {
      no: "Case 03", cat: "HR/노무", title: "TTC 기반 보상 체계 설계", tint: 2,
      bg: "스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.",
      steps: ["국내 IT 스타트업 보상 시장조사 및 벤치마킹", "TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)", "직급/직무별 Pay Band 설계", "스톡옵션 배분 기준 및 베스팅 스케줄 수립", "경영진 보고 및 의사결정 자료 준비"],
      result: "전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.",
      kpis: [["FiUsers", "전사", "적용 범위"], ["FiTarget", "TTC", "체계 도입"]],
      learn: "보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다.",
    },
  ];
  for (const cs of cases) {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Cases" });
    s.addText(cs.title, { placeholder: "title" });
    chip(s, `${cs.no} · ${cs.cat}`, 7.4, 0.42, 2.1, 0.4, PASTEL[cs.tint], C.text1, 11);
    const t = cs.tint;
    // left column
    const LX = 0.5, LW = 5.0, RX = 5.7, RW = 3.8, Y0 = 1.1, YEND = 5.2;
    // 배경
    const bgH = 1.5;
    tile(s, LX, Y0, LW, bgH, TINT[t]);
    label(s, "배경", LX + 0.25, Y0 + 0.18, 2);
    text(s, cs.bg, { x: LX + 0.25, y: Y0 + 0.45, w: LW - 0.5, h: bgH - 0.55, fontSize: 13, color: C.text1, valign: "top", lineSpacingMultiple: 1.15 });
    // 과정
    const pY = Y0 + bgH + 0.15, pH = YEND - pY;
    tile(s, LX, pY, LW, pH, C.background2);
    label(s, "과정", LX + 0.25, pY + 0.18, 2);
    const runs = cs.steps.map((st, i) => ({ text: st, options: { bullet: { indent: 14 }, breakLine: i < cs.steps.length - 1, paraSpaceAfter: 4 } }));
    text(s, runs, { x: LX + 0.25, y: pY + 0.45, w: LW - 0.5, h: pH - 0.55, fontSize: 13, color: C.text1, valign: "top", lineSpacingMultiple: 1.15 });
    // right column: KPI tile
    const kH = 0.8;
    tile(s, RX, Y0, RW, kH, TINT[(t + 1) % 3]);
    cs.kpis.forEach(([ic, big, lab], i) => {
      const kx = RX + 0.2 + i * (RW / 2);
      badge(s, kx, Y0 + 0.15, 0.5, PASTEL[(t + 1) % 3], ic);
      text(s, big, { x: kx + 0.6, y: Y0 + 0.08, w: RW / 2 - 0.8, h: 0.4, fontSize: 20, bold: true, color: C.text1, valign: "middle" });
      text(s, lab, { x: kx + 0.6, y: Y0 + 0.47, w: RW / 2 - 0.8, h: 0.25, fontSize: 10, color: C.text2 });
    });
    // 결과
    const rY = Y0 + kH + 0.15, rH = 1.25;
    tile(s, RX, rY, RW, rH, TINT[(t + 2) % 3]);
    label(s, "결과", RX + 0.25, rY + 0.18, 2);
    text(s, cs.result, { x: RX + 0.25, y: rY + 0.45, w: RW - 0.5, h: rH - 0.55, fontSize: 13, color: C.text1, valign: "top", lineSpacingMultiple: 1.15 });
    // 배운 점
    const lY = rY + rH + 0.15, lH = YEND - lY;
    tile(s, RX, lY, RW, lH, TINT[t]);
    label(s, "배운 점", RX + 0.25, lY + 0.18, 2);
    text(s, cs.learn, { x: RX + 0.25, y: lY + 0.45, w: RW - 0.5, h: lH - 0.5, fontSize: 13, color: C.text1, valign: "top", lineSpacingMultiple: 1.1 });
    s.addNotes(`${cs.no} — ${cs.title} (${cs.cat})`);
  }

  // =============== 10. Career Timeline ===============
  pres.addSection({ title: "Career" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Career" });
    s.addText("Career Timeline", { placeholder: "title" });
    text(s, "아래가 과거 · 위가 최근", { x: 7.2, y: TITLE_Y, w: 2.3, h: TITLE_H, fontSize: 11, color: C.text2, align: "right", valign: "middle" });
    const rows = [
      ["FiCreditCard", "2026.04", "스톡옵션 세무 체계 정립", "스톡옵션 법률 구조 및 과세 체계 분석 완료. 임직원 보상 설계의 세무 관점 가이드라인 수립."],
      ["FiFlag", "2026.04", "ABTC 인증 처리 완료", "APEC 기업인 여행카드(ABTC) 관련 문서 작업 및 인증 프로세스 완료."],
      ["FiFileText", "2026.04", "SHA(주주간계약) 분석", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 등 주주간계약의 핵심 조항 분석 및 실무 적용."],
      ["FiActivity", "2026.04", "IT 스타트업 런웨이 관리 체계 수립", "현금흐름 기반 런웨이 관리 프레임워크 학습 및 GXC 적용 방안 도출."],
      ["FiBriefcase", "2026 ~", "GXC 경영지원팀 합류", "재무/회계, IR, HR, 정부사업 등 경영지원 전반 업무 담당. 스타트업 운영 기반 구축."],
    ];
    const rh = 0.72, gap = 0.1, y0 = 1.1;
    // connector line behind badges
    s.addShape(pres.shapes.LINE, { x: 0.5 + 0.3, y: y0 + 0.36, w: 0, h: (rh + gap) * 4, line: { color: C.accent4, width: 2 }, objectName: nm("Connector") });
    rows.forEach(([ic, date, h, d], i) => {
      const y = y0 + i * (rh + gap);
      badge(s, 0.5, y + 0.06, 0.6, PASTEL[i % 3], ic);
      tile(s, 1.4, y, 8.1, rh, TINT[i % 3]);
      text(s, date, { x: 1.65, y: y + 0.1, w: 0.95, h: 0.52, fontSize: 12, bold: true, color: C.text2, valign: "middle" });
      text(s, h, { x: 2.65, y: y + 0.08, w: 6.7, h: 0.28, fontSize: 14, bold: true, color: C.text1 });
      text(s, d, { x: 2.65, y: y + 0.38, w: 6.7, h: 0.3, fontSize: 11, color: C.text2 });
    });
    s.addNotes("Career Timeline — 시간순, 아래가 과거");
  }

  // =============== 11. Certifications & Learning ===============
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Career" });
    s.addText("Certifications & Learning", { placeholder: "title" });
    // left: 자격증
    {
      const x = 0.5, y = 1.1, w = 3.9, h = 4.05;
      tile(s, x, y, w, h, C.accent5);
      badge(s, x + 0.3, y + 0.3, 0.55, C.accent2, "FiAward");
      text(s, "자격증 (준비중)", { x: x + 1.0, y: y + 0.3, w: 3, h: 0.55, fontSize: 17, bold: true, color: C.text1, valign: "middle" });
      const certs = [["전산회계 1급", "한국세무사회", "2026.06 목표"], ["전산세무 2급", "한국세무사회", "2026.06 목표"]];
      certs.forEach(([n, org, tgt], i) => {
        const cy = y + 1.15 + i * 1.4;
        tile(s, x + 0.3, cy, w - 0.6, 1.2, C.background1);
        text(s, n, { x: x + 0.55, y: cy + 0.18, w: 2.6, h: 0.35, fontSize: 16, bold: true, color: C.text1 });
        text(s, org, { x: x + 0.55, y: cy + 0.58, w: 2, h: 0.3, fontSize: 12, color: C.text2 });
        chip(s, tgt, x + w - 0.55 - 1.3, cy + 0.6, 1.3, 0.32, C.accent6, C.text1, 10);
      });
    }
    // right: 최근 학습
    {
      const x = 4.6, y = 1.1, w = 4.9, h = 4.05;
      tile(s, x, y, w, h, C.accent4);
      badge(s, x + 0.3, y + 0.3, 0.55, C.accent1, "FiBookOpen");
      text(s, "최근 학습", { x: x + 1.0, y: y + 0.3, w: 3, h: 0.55, fontSize: 17, bold: true, color: C.text1, valign: "middle" });
      const learns = [
        ["2026.04.16", "스톡옵션 과세 체계", "IR/투자"],
        ["2026.04.14", "스톡옵션 법률 구조", "IR/투자"],
        ["2026.04.10", "주주간계약(SHA) 핵심 조항", "IR/투자"],
        ["2026.04.08", "IT 스타트업 런웨이 관리", "재무"],
      ];
      learns.forEach(([d, t, cat], i) => {
        const ly = y + 1.15 + i * 0.7;
        tile(s, x + 0.3, ly, w - 0.6, 0.6, C.background1);
        text(s, d, { x: x + 0.5, y: ly + 0.07, w: 1.1, h: 0.24, fontSize: 11, bold: true, color: C.text2, valign: "middle" });
        chip(s, cat, x + 0.5, ly + 0.33, 0.72, 0.21, cat === "재무" ? C.accent6 : C.accent5, C.text1, 8);
        text(s, t, { x: x + 1.6, y: ly, w: w - 1.9, h: 0.6, fontSize: 13, bold: true, color: C.text1, valign: "middle" });
      });
    }
    s.addNotes("Certifications & Learning");
  }

  // =============== 12. Project Gallery ===============
  pres.addSection({ title: "Projects" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Projects" });
    s.addText("Project Gallery", { placeholder: "title" });
    const projects = [
      ["IR/투자", "SHA 주주간계약 검토", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 분석 및 투자 계약 실무 적용", "투자 계약 핵심 조항 가이드 문서화"],
      ["IR/투자", "스톡옵션 체계 설계", "법률 구조 분석, 과세 체계 정립, 임직원 보상 설계 가이드라인 수립", "스톡옵션 세무/법률 통합 가이드 완성"],
      ["HR/노무", "급여 체계 설계", "TTC(Total Target Compensation) 기반 급여 체계 설계 및 보상 구조 수립", "전사 보상 체계 프레임워크 구축"],
      ["정부사업", "ABTC 인증 처리", "APEC 기업인 여행카드 관련 서류 준비 및 인증 프로세스 관리", "인증 완료 및 프로세스 SOP 문서화"],
      ["재무/회계", "런웨이 관리 체계", "IT 스타트업 현금흐름 기반 런웨이 관리 프레임워크 구축", "월별 런웨이 모니터링 대시보드 구축"],
      ["재무/회계", "월말 결산 프로세스", "매월 재무제표 마감 및 결산 프로세스 체계화, 전표 처리 자동화", "결산 소요 시간 단축 및 프로세스 표준화"],
    ];
    const catColor = { "IR/투자": C.accent2, "HR/노무": C.accent3, "정부사업": C.accent1, "재무/회계": C.accent1 };
    const gx = 0.5, gy = 1.1, tw = 2.87, th = 1.95, gap = 0.195;
    projects.forEach(([cat, h, d, res], i) => {
      const x = gx + (i % 3) * (tw + gap), y = gy + Math.floor(i / 3) * (th + gap);
      tile(s, x, y, tw, th, TINT[i % 3]);
      chip(s, cat, x + 0.2, y + 0.15, 0.95, 0.26, catColor[cat], C.text1, 9);
      text(s, h, { x: x + 0.2, y: y + 0.46, w: tw - 0.4, h: 0.3, fontSize: 14, bold: true, color: C.text1 });
      text(s, d, { x: x + 0.2, y: y + 0.77, w: tw - 0.4, h: 0.6, fontSize: 10.5, color: C.text2, valign: "top", lineSpacingMultiple: 1.08 });
      s.addImage({ data: ICON.FiCheckCircle, x: x + 0.2, y: y + th - 0.42, w: 0.2, h: 0.2, objectName: nm("Icon"), altText: "성과" });
      text(s, res, { x: x + 0.47, y: y + th - 0.5, w: tw - 0.65, h: 0.36, fontSize: 10.5, bold: true, color: C.text1, valign: "middle", lineSpacingMultiple: 1.05 });
    });
    s.addNotes("Project Gallery 6개");
  }

  // =============== 13. 마무리 ===============
  pres.addSection({ title: "Outro" });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Outro" });
    tile(s, 0.5, 0.55, 5.35, 4.5, C.background1, { r: 0.2 });
    text(s, "Let's Connect", { x: 0.9, y: 1.0, w: 4.6, h: 0.9, fontSize: 44, bold: true, color: C.text1 });
    text(s, "새로운 기회나 협업에 대해 이야기하고 싶으시다면 연락주세요.", { x: 0.9, y: 2.05, w: 4.5, h: 0.8, fontSize: 15, color: C.text2, valign: "top", lineSpacingMultiple: 1.25 });
    text(s, "양남규", { x: 0.9, y: 3.6, w: 2.5, h: 0.45, fontSize: 20, bold: true, color: C.text1 });
    text(s, "Finance Operations Specialist", { x: 0.9, y: 4.05, w: 4.4, h: 0.3, fontSize: 12, color: C.accent1, bold: true });
    text(s, "© 2026 양남규", { x: 0.9, y: 4.55, w: 4.4, h: 0.3, fontSize: 10, color: C.text2 });
    // right: 2 white tiles Email / LinkedIn + 2 small
    const grid = [["FiMail", "Email"], ["FiLinkedin", "LinkedIn"]];
    const gx = 6.15, tw = 3.35, th = 2.15, gap = 0.2;
    grid.forEach(([ic, lab], i) => {
      const y = 0.55 + i * (th + gap);
      tile(s, gx, y, tw, th, C.background1, { r: 0.2 });
      badge(s, gx + 0.35, y + (th - 0.9) / 2, 0.9, TINT[i + 1], ic);
      text(s, lab, { x: gx + 1.5, y: y + (th - 0.5) / 2, w: 1.7, h: 0.5, fontSize: 22, bold: true, color: C.text1, valign: "middle" });
    });
    s.addNotes("마무리 — Let's Connect");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
