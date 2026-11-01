// Mono Minimal — 양남규 포트폴리오 덱 (13장)
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
pres.title = "양남규 포트폴리오 — Mono Minimal";
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
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  objects: [
    { text: { text: "양남규 · Portfolio 2026", options: { x: 0.6, y: 5.22, w: 3, h: 0.25, fontSize: 9, color: HEX.dk2, margin: 0, isTextBox: true } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 8.8, h: 0.75, fontSize: 36, bold: true, color: HEX.dk1, margin: 0, valign: "top", align: "left" } } },
  ],
  slideNumber: { x: 8.9, y: 5.22, w: 0.5, h: 0.25, fontSize: 9, color: HEX.dk2, align: "right", margin: 0 },
});

// ---------- Helpers ----------
const T = (s, o) => Object.assign({ isTextBox: true, margin: 0, color: C.text1 }, o);

function contentSlide(num, label, title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  // Motif: giant light-gray number behind content
  s.addText(num, T(null, { x: 0.3, y: -0.75, w: 5.5, h: 3.3, fontSize: 250, bold: true, color: C.background2, valign: "top", align: "left", objectName: "Giant number " + num }));
  s.addText(label, T(null, { x: 0.6, y: 0.36, w: 4, h: 0.25, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, objectName: "Section label" }));
  s.addText(title, { placeholder: "title" });
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

(async () => {
  const W = HEX.lt1, D = HEX.dk1;
  const ic = {
    fin: await icon(Fi.FiDollarSign, W), ir: await icon(Fi.FiTrendingUp, W), hr: await icon(Fi.FiUsers, W), gov: await icon(Fi.FiBriefcase, W),
    xl: await icon(Fi.FiSearch, W), pv: await icon(Fi.FiGrid, W), cf: await icon(Fi.FiAlertCircle, W), vba: await icon(Fi.FiZap, W),
    award: await icon(Fi.FiAward, W), book: await icon(Fi.FiBookOpen, W),
    mail: await icon(Fi.FiMail, D), linkedin: await icon(Fi.FiLinkedin, D),
  };

  // ===== 1. 표지 =====
  pres.addSection({ title: "Intro" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    s.addText("Hello, I'm", T(null, { x: 0.8, y: 1.35, w: 5, h: 0.4, fontSize: 18, color: C.text2, objectName: "Greeting" }));
    s.addText([{ text: "양남규" }, { text: ".", options: { color: C.accent1 } }], T(null, { x: 0.8, y: 1.75, w: 7, h: 1.4, fontSize: 88, bold: true, objectName: "Name" }));
    s.addText("Finance Operations Specialist", T(null, { x: 0.8, y: 3.25, w: 8, h: 0.45, fontSize: 22, bold: true, color: C.text2, objectName: "Role" }));
    s.addText("스타트업의 성장을 뒷받침하는 경영지원 전문가", T(null, { x: 0.8, y: 3.75, w: 8, h: 0.4, fontSize: 16, objectName: "Tagline" }));
    s.addText("재무 · IR · HR · 컴플라이언스  /  GXC 경영지원팀  /  2026", T(null, { x: 0.8, y: 4.75, w: 8, h: 0.3, fontSize: 11, color: C.text2, objectName: "Sub" }));
  }

  // ===== 2. About Me =====
  pres.addSection({ title: "About" });
  {
    const s = contentSlide("02", "ABOUT", "About Me", "About");
    s.addText("GXC에서 경영지원 업무를 담당하며, 재무/회계 관리부터 투자 관련 법률 검토, 인사/노무 행정, 정부사업 관리까지 폭넓은 영역을 수행하고 있습니다. 체계적인 학습과 실무 경험을 병행하며 스타트업의 성장에 필요한 오퍼레이션 역량을 꾸준히 쌓아가고 있습니다.",
      T(null, { x: 0.6, y: 1.75, w: 8.4, h: 1.5, fontSize: 16, lineSpacingMultiple: 1.35, valign: "top", objectName: "About body" }));
    const tags = ["재무/회계", "IR/투자", "HR/노무", "정부사업", "Excel 고급", "데이터분석", "컴플라이언스"];
    tags.forEach((t, i) => {
      const row = i < 4 ? 0 : 1, col = i < 4 ? i : i - 4;
      pill(s, t, 0.6 + col * 2.15, 3.7 + row * 0.6, 1.95, 0.44, { fontSize: 12 });
    });
  }

  // ===== 3. 업무 영역 =====
  pres.addSection({ title: "Work" });
  {
    const s = contentSlide("03", "SERVICES", "업무 영역", "Work");
    const cards = [
      ["재무 / 회계", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], ic.fin],
      ["IR / 투자", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], ic.ir],
      ["HR / 노무", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], ic.hr],
      ["정부사업 / 기술", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], ic.gov],
    ];
    cards.forEach(([t, items, img], i) => {
      const x = 0.6 + i * 2.2;
      iconCircle(s, img, x, 1.75, 0.62, t);
      s.addText(t, T(null, { x, y: 2.55, w: 2.0, h: 0.4, fontSize: 17, bold: true, objectName: "Card title " + t }));
      s.addText(items.map((it, j) => ({ text: it, options: { bullet: { indent: 12 }, breakLine: j < items.length - 1 } })),
        T(null, { x, y: 3.05, w: 2.0, h: 1.8, fontSize: 14, color: C.text2, paraSpaceAfter: 6, valign: "top", objectName: "Card items " + t }));
    });
  }

  // ===== 4. Skills =====
  {
    const s = contentSlide("04", "SKILLS", "Skills & Expertise", "Work");
    const groups = [
      ["재무/회계", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], [75, 70, 80, 65]],
      ["IR/투자", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], [70, 65, 60, 60]],
      ["HR/노무", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], [75, 70, 65, 60]],
      ["정부사업/기술", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], [70, 65, 85, 60]],
    ];
    groups.forEach(([name, labels, vals], i) => {
      const x = 0.6 + (i % 2) * 4.5, y = 1.6 + Math.floor(i / 2) * 1.75;
      heading(s, name, x, y, 4.2);
      const colors = vals.map(v => (v === 85 ? HEX.accent1 : HEX.dk2));
      s.addChart(pres.charts.BAR, [{ name: name, labels, values: vals }], {
        x, y: y + 0.22, w: 4.2, h: 1.5, barDir: "bar", barGapWidthPct: 55,
        chartColors: colors,
        showLegend: false, showTitle: false,
        valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: "none" }, catGridLine: { style: "none" },
        catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 10, catAxisLabelColor: HEX.dk1, catAxisLineShow: false, catAxisOrientation: "maxMin",
        showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: "+mn-lt", dataLabelFontSize: 10, dataLabelColor: HEX.dk2, dataLabelFormatCode: '0"%"',
        objectName: "Skills chart " + name,
      });
    });
  }

  // ===== 5. Excel Mastery =====
  {
    const s = contentSlide("05", "EXCEL", "Excel Mastery", "Work");
    s.addText("실무에서 활용하는 엑셀 고급 함수와 자동화 기법", T(null, { x: 0.6, y: 1.32, w: 8, h: 0.3, fontSize: 13, color: C.text2, objectName: "Subtitle" }));
    const items = [
      ["XLOOKUP", "급여 대장 자동 매칭 / 거래처별 매출 추적 / 계정과목 자동 분류", ["VLOOKUP 대체", "양방향 검색", "에러 처리 내장"], ic.xl],
      ["피벗테이블", "부서별 예산 집행 분석 / 월별 비용 추이 리포트", ["크로스탭 분석", "슬라이서 필터", "자동 갱신"], ic.pv],
      ["조건부 서식", "예산 초과 자동 알림 (달성률 100% 초과 빨강, 90% 미만 노랑)", ["데이터 막대", "아이콘 세트", "색조 규칙"], ic.cf],
      ["자동화 매크로(VBA)", "월말 결산 자동화 / 급여명세서 자동 발송", ["VBA 매크로", "반복 업무 제거", "3시간→15분"], ic.vba],
    ];
    items.forEach(([t, uses, badges, img], i) => {
      const x = 0.6 + (i % 2) * 4.5, y = 1.85 + Math.floor(i / 2) * 1.65;
      iconCircle(s, img, x, y, 0.5, t);
      s.addText(t, T(null, { x: x + 0.7, y: y + 0.02, w: 3.5, h: 0.4, fontSize: 17, bold: true, objectName: "Excel title " + t }));
      s.addText(uses, T(null, { x: x + 0.7, y: y + 0.46, w: 3.5, h: 0.5, fontSize: 12, color: C.text2, valign: "top", objectName: "Excel uses " + t }));
      badges.forEach((b, j) => pill(s, b, x + 0.7 + j * 1.2, y + 1.05, 1.12, 0.3, { fontSize: 10 }));
    });
  }

  // ===== 6. 자동화 성과 =====
  {
    const s = contentSlide("06", "RESULTS", "자동화 성과", "Work");
    const stats = [
      ["3시간 →", "15분", "월말 결산 보고서 생성", "전표 정리 → 집계 → 보고서, 버튼 하나"],
      ["3일 →", "1.2일", "월말 결산 전체 소요 시간", "약 60% 단축"],
      ["", "3단계", "전표 정리 · 계정과목별 집계 · 보고서 출력", "PDF 저장 · 메일 발송"],
    ];
    stats.forEach(([from, big, l1, l2], i) => {
      const x = 0.6 + i * 2.95;
      s.addText(from || " ", T(null, { x, y: 1.55, w: 2.7, h: 0.3, fontSize: 14, color: C.text2, objectName: "Stat from " + big }));
      s.addText(big, T(null, { x, y: 1.82, w: 2.7, h: 0.8, fontSize: 48, bold: true, color: i === 2 ? C.text1 : C.accent1, objectName: "Stat big " + big }));
      s.addText(l1, T(null, { x, y: 2.64, w: 2.7, h: 0.5, fontSize: 13, bold: true, valign: "top", objectName: "Stat label " + big }));
      s.addText(l2, T(null, { x, y: 3.1, w: 2.7, h: 0.3, fontSize: 11, color: C.text2, objectName: "Stat sub " + big }));
    });
    const hdr = { bold: true, color: C.text1, fill: { color: C.background2 } };
    const rows = [
      [{ text: "부서", options: hdr }, { text: "Q1", options: Object.assign({ align: "right" }, hdr) }, { text: "Q2", options: Object.assign({ align: "right" }, hdr) }, { text: "증감률", options: Object.assign({ align: "right" }, hdr) }],
      ["개발팀", "45,200", "52,800", "+16.8%"],
      ["마케팅", "32,100", "28,500", "-11.2%"],
      ["경영지원", "18,300", "19,100", "+4.4%"],
      [{ text: "합계", options: { bold: true } }, { text: "95,600", options: { bold: true, align: "right" } }, { text: "100,400", options: { bold: true, align: "right" } }, { text: "+5.0%", options: { bold: true, align: "right", color: C.accent1 } }],
    ].map(r => r.map((c, ci) => (typeof c === "string" ? { text: c, options: { align: ci === 0 ? "left" : "right" } } : c)));
    heading(s, "예시 피벗 표", 0.6, 3.48, 6);
    s.addTable(rows, { x: 0.6, y: 3.75, w: 8.8, colW: [2.8, 2, 2, 2], rowH: 0.26, fontSize: 11, color: C.text1, fontFace: "Malgun Gothic",
      border: { type: "solid", pt: 1, color: HEX.lt1 }, margin: [0.02, 0.1, 0.02, 0.1], valign: "middle", objectName: "Pivot table" });
  }

  // ===== 7-9. Cases =====
  pres.addSection({ title: "Cases" });
  const cases = [
    { n: "07", label: "CASE 01 · IR/투자", title: "SHA 주주간계약 검토",
      bg: "GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.",
      steps: ["ROFR(우선매수권) 조항의 실무 적용 시나리오 분석", "Drag-Along · Tag-Along 권리의 행사 조건 정리", "희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)", "투자자 동의권(Consent Rights) 범위 및 영향 분석", "법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화"],
      result: "5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.",
      kpis: [["5개", "핵심 조항 분석"], ["1건", "가이드 문서"]],
      learn: "법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다." },
    { n: "08", label: "CASE 02 · 재무/회계", title: "월말 결산 프로세스 혁신",
      bg: "매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.",
      steps: ["기존 결산 프로세스 전수조사 및 병목 구간 파악", "전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)", "계정과목별 자동 집계 템플릿 구축", "VBA 매크로를 활용한 보고서 자동 생성", "결산 체크리스트 및 SOP 문서 작성"],
      result: "결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.",
      kpis: [["-60%", "소요 시간"], ["SOP", "프로세스 표준화"]],
      learn: "Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다." },
    { n: "09", label: "CASE 03 · HR/노무", title: "TTC 기반 보상 체계 설계",
      bg: "스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.",
      steps: ["국내 IT 스타트업 보상 시장조사 및 벤치마킹", "TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)", "직급/직무별 Pay Band 설계", "스톡옵션 배분 기준 및 베스팅 스케줄 수립", "경영진 보고 및 의사결정 자료 준비"],
      result: "전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.",
      kpis: [["전사", "적용 범위"], ["TTC", "체계 도입"]],
      learn: "보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다." },
  ];
  for (const c of cases) {
    const s = contentSlide(c.n, c.label, c.title, "Cases");
    const LX = 0.6, LW = 4.7, RX = 5.7, RW = 3.7;
    heading(s, "배경", LX, 1.55, LW);
    s.addText(c.bg, T(null, { x: LX, y: 1.83, w: LW, h: 1.2, fontSize: 13, lineSpacingMultiple: 1.25, valign: "top", objectName: "Background" }));
    heading(s, "과정", LX, 3.15, LW);
    s.addText(c.steps.map((st, j) => ({ text: st, options: { bullet: { indent: 12 }, breakLine: j < c.steps.length - 1 } })),
      T(null, { x: LX, y: 3.43, w: LW, h: 1.7, fontSize: 13, paraSpaceAfter: 4, valign: "top", objectName: "Steps" }));
    heading(s, "지표", RX, 1.55, RW);
    c.kpis.forEach(([big, lab], j) => {
      const x = RX + j * 1.85;
      s.addText(big, T(null, { x, y: 1.8, w: 1.8, h: 0.58, fontSize: 36, bold: true, color: C.accent1, objectName: "KPI " + big }));
      s.addText(lab, T(null, { x, y: 2.38, w: 1.8, h: 0.25, fontSize: 11, color: C.text2, objectName: "KPI label " + lab }));
    });
    heading(s, "결과", RX, 2.72, RW);
    s.addText(c.result, T(null, { x: RX, y: 3.0, w: RW, h: 0.8, fontSize: 13, lineSpacingMultiple: 1.25, valign: "top", objectName: "Result" }));
    heading(s, "배운 점", RX, 3.85, RW);
    s.addText(c.learn, T(null, { x: RX, y: 4.13, w: RW, h: 0.95, fontSize: 12, color: C.text2, lineSpacingMultiple: 1.25, valign: "top", objectName: "Learned" }));
  }

  // ===== 10. Career Timeline =====
  pres.addSection({ title: "Career" });
  {
    const s = contentSlide("10", "TIMELINE", "Career Timeline", "Career");
    const items = [
      ["2026.04", "스톡옵션 세무 체계 정립", "스톡옵션 법률 구조 및 과세 체계 분석 완료. 임직원 보상 설계의 세무 관점 가이드라인 수립."],
      ["2026.04", "ABTC 인증 처리 완료", "APEC 기업인 여행카드(ABTC) 관련 문서 작업 및 인증 프로세스 완료."],
      ["2026.04", "SHA(주주간계약) 분석", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 등 주주간계약의 핵심 조항 분석 및 실무 적용."],
      ["2026.04", "IT 스타트업 런웨이 관리 체계 수립", "현금흐름 기반 런웨이 관리 프레임워크 학습 및 GXC 적용 방안 도출."],
      ["2026 ~", "GXC 경영지원팀 합류", "재무/회계, IR, HR, 정부사업 등 경영지원 전반 업무 담당. 스타트업 운영 기반 구축."],
    ];
    items.forEach(([date, t, d], i) => {
      const y = 1.65 + i * 0.7;
      s.addText(date, T(null, { x: 0.6, y: y + 0.02, w: 1.3, h: 0.3, fontSize: 14, bold: true, color: C.text2, objectName: "Date " + i }));
      s.addShape(pres.shapes.OVAL, { x: 2.05, y: y + 0.08, w: 0.16, h: 0.16, fill: { color: i === 0 ? C.accent1 : C.text2 }, line: { color: i === 0 ? C.accent1 : C.text2, width: 0 }, objectName: "Dot " + i });
      s.addText(t, T(null, { x: 2.45, y, w: 6.9, h: 0.3, fontSize: 15, bold: true, objectName: "Event " + i }));
      s.addText(d, T(null, { x: 2.45, y: y + 0.31, w: 6.9, h: 0.36, fontSize: 11, color: C.text2, valign: "top", objectName: "Event desc " + i }));
    });
  }

  // ===== 11. Certifications & Learning =====
  {
    const s = contentSlide("11", "LEARNING", "Certifications & Learning", "Career");
    heading(s, "자격증 (준비중)", 0.6, 1.6, 4);
    const certs = [["전산회계 1급", "한국세무사회", "2026.06 목표"], ["전산세무 2급", "한국세무사회", "2026.06 목표"]];
    certs.forEach(([n, org, goal], i) => {
      const y = 2.0 + i * 1.35;
      iconCircle(s, ic.award, 0.6, y, 0.6, n);
      s.addText(n, T(null, { x: 1.4, y: y - 0.02, w: 3, h: 0.4, fontSize: 18, bold: true, objectName: "Cert " + n }));
      s.addText(org, T(null, { x: 1.4, y: y + 0.38, w: 3, h: 0.3, fontSize: 13, color: C.text2, objectName: "Cert org " + n }));
      s.addText(goal, T(null, { x: 1.4, y: y + 0.72, w: 2.5, h: 0.28, fontSize: 11, bold: true, color: C.accent1, objectName: "Cert goal " + n }));
    });
    heading(s, "최근 학습", 4.85, 1.6, 4.5);
    const learn = [["2026.04.16", "스톡옵션 과세 체계", "IR/투자"], ["2026.04.14", "스톡옵션 법률 구조", "IR/투자"], ["2026.04.10", "주주간계약(SHA) 핵심 조항", "IR/투자"], ["2026.04.08", "IT 스타트업 런웨이 관리", "재무"]];
    learn.forEach(([d, t, cat], i) => {
      const y = 2.0 + i * 0.72;
      s.addText(d, T(null, { x: 4.85, y: y + 0.02, w: 1.0, h: 0.3, fontSize: 11, color: C.text2, objectName: "Learn date " + i }));
      s.addText(t, T(null, { x: 5.85, y, w: 2.75, h: 0.5, fontSize: 14, bold: true, valign: "top", objectName: "Learn topic " + i }));
      pill(s, cat, 8.65, y + 0.02, 0.75, 0.28, { fontSize: 10 });
    });
  }

  // ===== 12. Project Gallery =====
  pres.addSection({ title: "Projects" });
  {
    const s = contentSlide("12", "PROJECTS", "Project Gallery", "Projects");
    const projects = [
      ["IR/투자", "SHA 주주간계약 검토", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 분석 및 투자 계약 실무 적용", "투자 계약 핵심 조항 가이드 문서화"],
      ["IR/투자", "스톡옵션 체계 설계", "법률 구조 분석, 과세 체계 정립, 임직원 보상 설계 가이드라인 수립", "스톡옵션 세무/법률 통합 가이드 완성"],
      ["HR/노무", "급여 체계 설계", "TTC(Total Target Compensation) 기반 급여 체계 설계 및 보상 구조 수립", "전사 보상 체계 프레임워크 구축"],
      ["정부사업", "ABTC 인증 처리", "APEC 기업인 여행카드 관련 서류 준비 및 인증 프로세스 관리", "인증 완료 및 프로세스 SOP 문서화"],
      ["재무/회계", "런웨이 관리 체계", "IT 스타트업 현금흐름 기반 런웨이 관리 프레임워크 구축", "월별 런웨이 모니터링 대시보드 구축"],
      ["재무/회계", "월말 결산 프로세스", "매월 재무제표 마감 및 결산 프로세스 체계화, 전표 처리 자동화", "결산 소요 시간 단축 및 프로세스 표준화"],
    ];
    projects.forEach(([cat, t, d, res], i) => {
      const x = 0.6 + (i % 3) * 3.0, y = 1.6 + Math.floor(i / 3) * 1.8;
      pill(s, cat, x, y, 0.95, 0.26, { fontSize: 10, fill: { color: C.text2 }, color: C.background1 });
      s.addText(t, T(null, { x, y: y + 0.34, w: 2.8, h: 0.32, fontSize: 15, bold: true, objectName: "Project " + t }));
      s.addText(d, T(null, { x, y: y + 0.68, w: 2.75, h: 0.6, fontSize: 10.5, color: C.text2, valign: "top", lineSpacingMultiple: 1.15, objectName: "Project desc " + t }));
      s.addText([{ text: "성과  ", options: { bold: true, color: C.accent1 } }, { text: res }],
        T(null, { x, y: y + 1.3, w: 2.75, h: 0.4, fontSize: 10.5, valign: "top", objectName: "Project result " + t }));
    });
  }

  // ===== 13. 마무리 =====
  pres.addSection({ title: "Contact" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Contact" });
    s.addText("LET'S CONNECT", T(null, { x: 0.8, y: 1.35, w: 5, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 3, objectName: "Connect label" }));
    s.addText([{ text: "양남규" }, { text: ".", options: { color: C.accent1 } }], T(null, { x: 0.8, y: 1.7, w: 7, h: 1.3, fontSize: 80, bold: true, objectName: "Name" }));
    s.addText("새로운 기회나 협업에 대해 이야기하고 싶으시다면 연락주세요.", T(null, { x: 0.8, y: 3.15, w: 8, h: 0.4, fontSize: 16, objectName: "Message" }));
    [["Email", ic.mail], ["LinkedIn", ic.linkedin]].forEach(([lab, img], i) => {
      const x = 0.8 + i * 1.9;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.85, w: 1.7, h: 0.5, rectRadius: 0.25, fill: { color: C.background2 }, line: { color: C.background2, width: 0 }, objectName: "Contact pill " + lab });
      s.addImage({ data: img, x: x + 0.22, y: 3.98, w: 0.24, h: 0.24, objectName: "Contact icon " + lab });
      s.addText(lab, T(null, { x: x + 0.55, y: 3.85, w: 1.1, h: 0.5, fontSize: 13, bold: true, valign: "middle", objectName: "Contact " + lab }));
    });
    s.addText("© 2026 양남규", T(null, { x: 0.8, y: 4.95, w: 4, h: 0.25, fontSize: 10, color: C.text2, objectName: "Copyright" }));
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch(e => { console.error(e); process.exit(1); });
