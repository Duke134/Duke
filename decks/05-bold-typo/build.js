// Bold Typo — 양남규 포트폴리오 덱 (13장)
// Run: NODE_PATH=<node_modules> node build.js   (writes deck.pptx next to this file)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const Fa = require("react-icons/fa");

const OUT = path.join(__dirname, "deck.pptx");

// ---------- THEME ----------
const THEME = {
  name: "Bold Typo",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "1B1F3B", // 본문
    lt1: "FFFFFF", // 흰
    dk2: "2F3C7E", // 네이비
    lt2: "F9E795", // 골드
    accent1: "F96167", // 코랄
    accent2: "2F3C7E", // 네이비
    accent3: "F9E795", // 골드
    accent4: "1B1F3B", // 본문
    accent5: "F3F4F9", // 아주 연한 네이비 틴트(카드 바탕)
    accent6: "FFFFFF",
    hlink: "2F3C7E",
    folHlink: "F96167",
  },
};
const HEX = THEME.colors;

// ---------- applyTheme (inlined from the pptx skill's apply_theme.js) ----------
function loadJSZip() {
  return require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
}
const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
async function applyTheme(deckPath, theme) {
  const zip = await loadJSZip().loadAsync(fs.readFileSync(deckPath));
  const part = "ppt/theme/theme1.xml";
  const name = String(theme.name).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const scheme = `<a:clrScheme name="${name}">` + SLOTS.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k].toUpperCase()}"/></a:${k}>`).join("") + "</a:clrScheme>";
  const xml = await zip.file(part).async("string");
  const out = xml
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${name}"`);
  zip.file(part, out);
  for (const n of Object.keys(zip.files)) {
    if (!n.endsWith(".xml")) continue;
    const bad = (await zip.file(n).async("string")).match(/<a:srgbClr val="((?![0-9A-Fa-f]{6}")[^"]*)"/);
    if (bad) throw new Error(`${n}: srgbClr val "${bad[1]}" is not hex — a scheme color was passed to a hex-only option.`);
  }
  fs.writeFileSync(deckPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

// ---------- icons ----------
async function iconData(Icon, hex) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.author = "양남규";
  pres.title = "양남규 포트폴리오 — Bold Typo";
  const C = pres.SchemeColor;

  const W = 10, H = 5.625;
  const NAVY = C.text2, CORAL = C.accent1, GOLD = C.background2, WHITE = C.background1, INK = C.text1, TINT = C.accent5;

  // Block color roles: which text colors sit on each block fill
  const ROLE = {
    navy: { fill: NAVY, label: GOLD, kw: GOLD, sub: WHITE, body: WHITE, soft: GOLD },
    coral: { fill: CORAL, label: NAVY, kw: NAVY, sub: INK, body: INK, soft: NAVY },
    gold: { fill: GOLD, label: CORAL, kw: NAVY, sub: INK, body: INK, soft: NAVY },
  };
  // Motif geometry
  const LB_W = 3.3; // left block width (left/right split)
  const RX = 3.65, RW = 5.85; // right content region
  const TB_H = 1.6; // top block height (top/bottom split)
  const BY = 1.85; // bottom content start

  // ---------- Layouts ----------
  pres.defineSlideMaster({ title: "COVER", background: { color: HEX.accent1 } });
  pres.defineSlideMaster({ title: "CLOSING", background: { color: HEX.accent1 } });
  for (const k of Object.keys(ROLE)) {
    const r = ROLE[k];
    // left/right split
    pres.defineSlideMaster({
      title: `SL_${k}`,
      background: { color: HEX.lt1 },
      objects: [
        { rect: { x: 0, y: 0, w: LB_W, h: H, fill: { color: r.fill } } },
        { placeholder: { options: { name: "label", type: "body", x: 0.35, y: 0.4, w: 2.9, h: 0.3, fontSize: 11, bold: true, color: r.label, margin: 0, charSpacing: 2 }, text: "" } },
        { placeholder: { options: { name: "kw", type: "title", x: 0.3, y: 0.7, w: 3.0, h: 1.05, fontSize: 60, bold: true, color: r.kw, margin: 0, align: "left", valign: "top" }, text: "" } },
        { placeholder: { options: { name: "sub", type: "body", x: 0.35, y: 1.8, w: 2.7, h: 0.75, fontSize: 20, bold: true, color: r.sub, margin: 0, valign: "top" }, text: "" } },
      ],
    });
    // top/bottom split
    pres.defineSlideMaster({
      title: `ST_${k}`,
      background: { color: HEX.lt1 },
      objects: [
        { rect: { x: 0, y: 0, w: W, h: TB_H, fill: { color: r.fill } } },
        { placeholder: { options: { name: "label", type: "body", x: 0.5, y: 0.25, w: 3, h: 0.3, fontSize: 11, bold: true, color: r.label, margin: 0, charSpacing: 3 }, text: "" } },
        { placeholder: { options: { name: "kw", type: "title", x: 0.45, y: 0.5, w: 6.0, h: 1.0, fontSize: 60, bold: true, color: r.kw, margin: 0, align: "left", valign: "top" }, text: "" } },
        { placeholder: { options: { name: "sub", type: "body", x: 5.5, y: 0.5, w: 4.0, h: 0.95, fontSize: 20, bold: true, color: r.sub, margin: 0, align: "right", valign: "bottom" }, text: "" } },
      ],
    });
  }

  // ---------- helpers ----------
  const total = 13;
  let n = 0;
  function split(kind, role, section, label, kw, sub) {
    n += 1;
    const s = pres.addSlide({ masterName: `${kind}_${role}`, sectionTitle: section });
    s.addText(`${String(n).padStart(2, "0")}   ${label}`, { placeholder: "label" });
    s.addText(kw, { placeholder: "kw" });
    s.addText(sub, { placeholder: "sub" });
    return s;
  }
  function chip(s, x, y, w, h, text, fill, color, size = 10) {
    s.addText(text, { x, y, w, h, shape: pres.ShapeType.roundRect, rectRadius: h / 2, fill: { color: fill }, color, fontSize: size, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `chip ${text}` });
  }
  async function circleIcon(s, x, y, d, Icon, circleFill, iconHex, name) {
    s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: circleFill }, line: { color: circleFill, width: 0 }, objectName: `${name} circle` });
    const pad = d * 0.26;
    s.addImage({ data: await iconData(Icon, iconHex), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name });
  }
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", align: "left" }, o));

  // =====================================================================
  // 1. 표지 — 코랄 풀블리드, 네이비 타이포
  // =====================================================================
  pres.addSection({ title: "Intro" });
  n += 1;
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    txt(s, "Hello, I'm", { x: 0.6, y: 0.75, w: 6, h: 0.5, fontSize: 24, color: NAVY, objectName: "greeting" });
    txt(s, "양남규", { x: 0.55, y: 1.15, w: 8.9, h: 1.75, fontSize: 112, bold: true, color: NAVY, objectName: "name" });
    txt(s, "Finance Operations Specialist", { x: 0.6, y: 3.0, w: 8.8, h: 0.55, fontSize: 28, bold: true, color: INK, objectName: "role" });
    txt(s, "스타트업의 성장을 뒷받침하는 경영지원 전문가", { x: 0.6, y: 3.6, w: 8.8, h: 0.45, fontSize: 18, color: INK, objectName: "tagline" });
    // 보조 — 네이비 블록 위 골드 텍스트 (색면 분할의 예고)
    s.addShape(pres.ShapeType.rect, { x: 0, y: 4.75, w: W, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "footer block" });
    txt(s, "재무 · IR · HR · 컴플라이언스   /   GXC 경영지원팀   /   2026", { x: 0.6, y: 4.75, w: 8.8, h: 0.85, fontSize: 14, bold: true, color: GOLD, valign: "middle", objectName: "meta" });
    s.addNotes("표지: 양남규, Finance Operations Specialist");
  }

  // =====================================================================
  // 2. About Me — 좌 네이비 블록
  // =====================================================================
  {
    const s = split("SL", "navy", "Intro", "ABOUT", "소개", "About Me");
    txt(s, "GXC에서 경영지원 업무를 담당하며, 재무/회계 관리부터 투자 관련 법률 검토, 인사/노무 행정, 정부사업 관리까지 폭넓은 영역을 수행하고 있습니다. 체계적인 학습과 실무 경험을 병행하며 스타트업의 성장에 필요한 오퍼레이션 역량을 꾸준히 쌓아가고 있습니다.", {
      x: RX, y: 0.55, w: RW, h: 2.3, fontSize: 16, color: INK, lineSpacingMultiple: 1.25, objectName: "about body",
    });
    txt(s, "KEYWORDS", { x: RX, y: 2.75, w: 3, h: 0.3, fontSize: 11, bold: true, color: CORAL, charSpacing: 3, objectName: "keywords label" });
    const tags = ["재무/회계", "IR/투자", "HR/노무", "정부사업", "Excel 고급", "데이터분석", "컴플라이언스"];
    const fills = [[CORAL, WHITE], [NAVY, WHITE], [GOLD, NAVY]];
    tags.forEach((t, i) => {
      const col = i % 4, row = Math.floor(i / 4);
      const [f, c] = fills[i % 3];
      chip(s, RX + col * 1.5, 3.1 + row * 0.6, 1.35, 0.44, t, f, c, 12);
    });
    s.addNotes("About Me");
  }

  // =====================================================================
  // 3. 업무 영역 — 상 코랄 블록, 4개 아이콘 카드
  // =====================================================================
  pres.addSection({ title: "Expertise" });
  {
    const s = split("ST", "coral", "Expertise", "AREAS", "WORK", "업무 영역");
    const cards = [
      ["재무 / 회계", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], Fa.FaCalculator],
      ["IR / 투자", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], Fa.FaHandshake],
      ["HR / 노무", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], Fa.FaUsers],
      ["정부사업 / 기술", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], Fa.FaLandmark],
    ];
    const cw = 2.1, gap = 0.2;
    for (let i = 0; i < 4; i++) {
      const [head, items, Icon] = cards[i];
      const x = 0.5 + i * (cw + gap);
      const fill = i % 2 === 0 ? NAVY : CORAL;
      const iconHex = i % 2 === 0 ? HEX.lt2 : HEX.lt1;
      await circleIcon(s, x, BY + 0.1, 0.75, Icon, fill, iconHex, head);
      txt(s, `0${i + 1}`, { x: x + 0.95, y: BY + 0.33, w: 1.1, h: 0.3, fontSize: 11, bold: true, color: CORAL, charSpacing: 2, objectName: `card ${i + 1} index` });
      txt(s, head, { x, y: BY + 1.0, w: cw, h: 0.4, fontSize: 16, bold: true, color: NAVY, objectName: `card ${i + 1} head` });
      txt(s, items.map((t, j) => ({ text: t, options: { breakLine: j < items.length - 1, bullet: { indent: 14 }, paraSpaceAfter: 6 } })), {
        x, y: BY + 1.5, w: cw, h: 1.9, fontSize: 14, color: INK, objectName: `card ${i + 1} items`,
      });
    }
    s.addNotes("업무 영역 4개");
  }

  // =====================================================================
  // 4. Skills — 상 골드 블록, 네이티브 가로 막대 차트 4개
  // =====================================================================
  {
    const s = split("ST", "gold", "Expertise", "PROFICIENCY", "SKILLS", "Skills & Expertise");
    const groups = [
      ["재무/회계", ["세무 관리", "월말 결산", "전표 처리", "재무제표 분석"], [75, 70, 80, 65], HEX.accent1],
      ["IR/투자", ["주주간계약(SHA)", "스톡옵션 설계", "TTC 설계", "투자 라운드"], [70, 65, 60, 60], HEX.dk2],
      ["HR/노무", ["인사 관리", "근로계약 관리", "노동법 이해", "급여 체계 설계"], [75, 70, 65, 60], HEX.dk2],
      ["정부사업/기술", ["정부 보조금", "컴플라이언스", "Excel (고급)", "데이터 분석"], [70, 65, 85, 60], HEX.accent1],
    ];
    const cw = 4.4, ch = 1.72;
    groups.forEach((g, i) => {
      const [name, labels, values, color] = g;
      const x = 0.5 + (i % 2) * (cw + 0.2);
      const y = BY + Math.floor(i / 2) * (ch + 0.06);
      s.addChart(pres.ChartType.bar, [{ name, labels, values }], {
        x, y, w: cw, h: ch,
        barDir: "bar", barGapWidthPct: 45,
        chartColors: [color],
        showTitle: true, title: name, titleFontSize: 12, titleColor: HEX.dk2, titleFontFace: "+mn-lt", titleAlign: "left", titlePos: { x: 0, y: 0 },
        showLegend: false,
        showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: '0"%"',
        valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: "none" },
        catAxisLabelFontSize: 10, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin", catGridLine: { style: "none" },
        catAxisLineShow: false, valAxisLineShow: false,
        objectName: `skills chart ${name}`,
      });
    });
    s.addNotes("Skills & Expertise — 단위 %, 최고치 Excel 85");
  }

  // =====================================================================
  // 5. Excel Mastery — 상 네이비 블록, 4기법
  // =====================================================================
  {
    const s = split("ST", "navy", "Expertise", "TECHNIQUES", "EXCEL", "Excel Mastery");
    txt(s, "실무에서 활용하는 엑셀 고급 함수와 자동화 기법", { x: 0.5, y: BY - 0.05, w: 9, h: 0.3, fontSize: 12, color: CORAL, bold: true, objectName: "excel subtitle" });
    const tech = [
      ["XLOOKUP", ["급여 대장 자동 매칭", "거래처별 매출 추적", "계정과목 자동 분류"], ["VLOOKUP 대체", "양방향 검색", "에러 처리 내장"], Fa.FaSearch],
      ["피벗테이블", ["부서별 예산 집행 분석", "월별 비용 추이 리포트"], ["크로스탭 분석", "슬라이서 필터", "자동 갱신"], Fa.FaTable],
      ["조건부 서식", ["예산 초과 자동 알림 (달성률 100% 초과 빨강, 90% 미만 노랑)"], ["데이터 막대", "아이콘 세트", "색조 규칙"], Fa.FaPalette],
      ["자동화 매크로(VBA)", ["월말 결산 자동화", "급여명세서 자동 발송"], ["VBA 매크로", "반복 업무 제거", "3시간→15분"], Fa.FaCode],
    ];
    const cw = 2.15, gap = (9 - 4 * cw) / 3;
    for (let i = 0; i < 4; i++) {
      const [name, uses, badges, Icon] = tech[i];
      const x = 0.5 + i * (cw + gap);
      const y0 = BY + 0.4;
      await circleIcon(s, x, y0, 0.55, Icon, CORAL, HEX.lt1, name);
      txt(s, name, { x: x + 0.7, y: y0 - 0.05, w: cw - 0.7, h: 0.65, fontSize: 15, bold: true, color: NAVY, valign: "middle", objectName: `tech ${i + 1} name` });
      txt(s, uses.map((t, j) => ({ text: t, options: { breakLine: j < uses.length - 1, paraSpaceAfter: 3 } })), {
        x, y: y0 + 0.72, w: cw, h: 0.95, fontSize: 14, color: INK, objectName: `tech ${i + 1} uses`,
      });
      badges.forEach((b, j) => chip(s, x, y0 + 1.78 + j * 0.37, 1.55, 0.3, b, GOLD, NAVY, 10));
    }
    s.addNotes("Excel Mastery 4기법");
  }

  // =====================================================================
  // 6. 자동화 성과 — 좌 코랄 블록, 큰 숫자 콜아웃
  // =====================================================================
  {
    const s = split("SL", "coral", "Expertise", "AUTOMATION", "자동화", "자동화 성과");
    txt(s, "60%", { x: 0.3, y: 3.5, w: 3.0, h: 1.0, fontSize: 66, bold: true, color: NAVY, objectName: "sixty" });
    txt(s, "월말 결산 전체 소요 시간, 약 60% 단축", { x: 0.35, y: 4.55, w: 2.8, h: 0.6, fontSize: 12, bold: true, color: INK, objectName: "sixty label" });
    const rows = [
      ["3시간 → 15분", "월말 결산 보고서 생성 (전표 정리 → 집계 → 보고서, 버튼 하나)"],
      ["3일 → 1.2일", "월말 결산 전체 소요 시간, 약 60% 단축"],
      ["3단계", "전표 정리 · 계정과목별 집계 · 보고서 출력(PDF 저장·메일 발송)"],
    ];
    rows.forEach(([big, cap], i) => {
      const y = 0.5 + i * 1.55;
      txt(s, big, { x: RX, y, w: RW, h: 0.8, fontSize: 44, bold: true, color: i === 1 ? NAVY : CORAL, objectName: `stat ${i + 1}` });
      txt(s, cap, { x: RX, y: y + 0.85, w: RW, h: 0.6, fontSize: 14, color: INK, objectName: `stat ${i + 1} caption` });
    });
    s.addNotes("자동화 성과");
  }

  // =====================================================================
  // 7-9. Case 01~03 — 좌 블록 (네이비/코랄/골드)
  // =====================================================================
  pres.addSection({ title: "Cases" });
  const cases = [
    {
      role: "navy", num: "01", cat: "IR/투자", title: "SHA 주주간계약 검토",
      bg: "GXC의 투자 라운드 진행 과정에서 주주간계약(SHA)의 핵심 조항에 대한 체계적 분석이 필요했습니다. 법률 자문 내용을 실무 관점에서 정리하고, 향후 투자 계약 시 빠르게 참조할 수 있는 가이드를 만드는 것이 목표였습니다.",
      steps: ["ROFR(우선매수권) 조항의 실무 적용 시나리오 분석", "Drag-Along · Tag-Along 권리의 행사 조건 정리", "희석방지(Anti-dilution) 조항의 유형별 비교 (Full Ratchet vs Weighted Average)", "투자자 동의권(Consent Rights) 범위 및 영향 분석", "법률 자문사 미팅 내용 정리 및 실무 Q&A 문서화"],
      result: "5개 핵심 조항에 대한 실무 가이드 문서 완성. 향후 투자 계약 검토 시 체크리스트로 활용 가능한 레퍼런스 구축.",
      metrics: [["5개", "핵심 조항 분석"], ["1건", "가이드 문서"]],
      learned: "법률 용어를 실무 관점으로 번역하는 과정에서 투자 구조에 대한 깊은 이해를 얻었습니다. 특히 희석방지 조항이 후속 투자 라운드에 미치는 영향을 구체적으로 파악할 수 있었습니다.",
    },
    {
      role: "coral", num: "02", cat: "재무/회계", title: "월말 결산 프로세스 혁신",
      bg: "매월 결산 시 수작업 위주의 프로세스로 인해 3일 이상 소요되었고, 휴먼 에러 리스크가 높았습니다. 체계적인 결산 프로세스 구축과 Excel 자동화를 통한 효율화가 필요했습니다.",
      steps: ["기존 결산 프로세스 전수조사 및 병목 구간 파악", "전표 데이터 정리 자동화 (XLOOKUP + 피벗테이블 활용)", "계정과목별 자동 집계 템플릿 구축", "VBA 매크로를 활용한 보고서 자동 생성", "결산 체크리스트 및 SOP 문서 작성"],
      result: "결산 소요 시간 3일 → 1.2일로 약 60% 단축. 표준화된 SOP로 인한 담당자 변경 시에도 일관된 결산 품질 유지.",
      metrics: [["-60%", "소요 시간"], ["SOP", "프로세스 표준화"]],
      learned: "Excel 고급 기능만으로도 상당한 업무 자동화가 가능하다는 것을 체감했습니다. 프로세스를 먼저 정리하고 자동화하는 순서가 중요합니다.",
    },
    {
      role: "gold", num: "03", cat: "HR/노무", title: "TTC 기반 보상 체계 설계",
      bg: "스타트업 성장에 따라 직급/직무별 보상 기준이 필요했습니다. 시장 데이터 기반의 TTC(Total Target Compensation) 체계를 도입하여 공정하고 경쟁력 있는 보상 구조를 만들고자 했습니다.",
      steps: ["국내 IT 스타트업 보상 시장조사 및 벤치마킹", "TTC 구성요소 정의 (기본급 + 성과급 + 스톡옵션)", "직급/직무별 Pay Band 설계", "스톡옵션 배분 기준 및 베스팅 스케줄 수립", "경영진 보고 및 의사결정 자료 준비"],
      result: "전사 적용 가능한 보상 체계 프레임워크 구축 완료. 채용 시 보상 제안의 일관성과 투명성 확보.",
      metrics: [["전사", "적용 범위"], ["TTC", "체계 도입"]],
      learned: "보상 설계는 재무(비용 관리)와 HR(인재 확보)의 교차점에 있어, 크로스펑셔널 역량이 핵심이라는 것을 깊이 느꼈습니다.",
    },
  ];
  for (const c of cases) {
    const r = ROLE[c.role];
    const s = split("SL", c.role, "Cases", `CASE · ${c.cat}`, c.num, c.title);
    // 좌 블록: 지표
    c.metrics.forEach(([v, l], i) => {
      const y = 3.2 + i * 1.05;
      txt(s, v, { x: 0.35, y, w: 2.7, h: 0.65, fontSize: 36, bold: true, color: r.kw, objectName: `metric ${i + 1}` });
      txt(s, l, { x: 0.35, y: y + 0.62, w: 2.7, h: 0.3, fontSize: 12, bold: true, color: r.sub, objectName: `metric ${i + 1} label` });
    });
    // 우측 본문
    const L = (t) => ({ text: t + "   ", options: { bold: true, color: CORAL } });
    txt(s, [L("배경"), { text: c.bg, options: {} }], { x: RX, y: 0.45, w: RW, h: 1.0, fontSize: 14, color: INK, objectName: "background" });
    txt(s, [
      { text: "과정", options: { bold: true, color: CORAL, breakLine: true, paraSpaceAfter: 2 } },
      ...c.steps.map((t, j) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < c.steps.length - 1, paraSpaceAfter: 2 } })),
    ], { x: RX, y: 1.55, w: RW, h: 1.95, fontSize: 14, color: INK, objectName: "process" });
    txt(s, [L("결과"), { text: c.result, options: {} }], { x: RX, y: 3.6, w: RW, h: 0.7, fontSize: 14, color: INK, objectName: "result" });
    txt(s, [L("배운 점"), { text: c.learned, options: {} }], { x: RX, y: 4.35, w: RW, h: 0.95, fontSize: 14, color: INK, objectName: "learned" });
    s.addNotes(`Case ${c.num} — ${c.title} (${c.cat})`);
  }

  // =====================================================================
  // 10. Career Timeline — 좌 네이비 블록, 세로 타임라인
  // =====================================================================
  pres.addSection({ title: "Journey" });
  {
    const s = split("SL", "navy", "Journey", "TIMELINE", "2026", "Career Timeline");
    const ev = [
      ["2026.04", "스톡옵션 세무 체계 정립", "스톡옵션 법률 구조 및 과세 체계 분석 완료. 임직원 보상 설계의 세무 관점 가이드라인 수립."],
      ["2026.04", "ABTC 인증 처리 완료", "APEC 기업인 여행카드(ABTC) 관련 문서 작업 및 인증 프로세스 완료."],
      ["2026.04", "SHA(주주간계약) 분석", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 등 주주간계약의 핵심 조항 분석 및 실무 적용."],
      ["2026.04", "IT 스타트업 런웨이 관리 체계 수립", "현금흐름 기반 런웨이 관리 프레임워크 학습 및 GXC 적용 방안 도출."],
      ["2026 ~", "GXC 경영지원팀 합류", "재무/회계, IR, HR, 정부사업 등 경영지원 전반 업무 담당. 스타트업 운영 기반 구축."],
    ];
    const rowH = 0.93, y0 = 0.5, lineX = 4.78;
    s.addShape(pres.ShapeType.line, { x: lineX, y: y0 + 0.05, w: 0, h: rowH * 4 + 0.1, line: { color: NAVY, width: 1.5 }, objectName: "timeline line" });
    ev.forEach(([date, title, desc], i) => {
      const y = y0 + i * rowH;
      txt(s, date, { x: RX, y: y, w: 1.0, h: 0.3, fontSize: 12, bold: true, color: CORAL, objectName: `event ${i + 1} date` });
      s.addShape(pres.ShapeType.ellipse, { x: lineX - 0.09, y: y + 0.045, w: 0.18, h: 0.18, fill: { color: i === 0 ? CORAL : NAVY }, line: { color: WHITE, width: 1.5 }, objectName: `event ${i + 1} dot` });
      txt(s, title, { x: 5.05, y: y - 0.02, w: 4.45, h: 0.3, fontSize: 14, bold: true, color: NAVY, objectName: `event ${i + 1} title` });
      txt(s, desc, { x: 5.05, y: y + 0.3, w: 4.45, h: 0.6, fontSize: 11, color: INK, objectName: `event ${i + 1} desc` });
    });
    s.addNotes("Career Timeline — 아래가 과거");
  }

  // =====================================================================
  // 11. Certifications & Learning — 상 코랄 블록
  // =====================================================================
  {
    const s = split("ST", "coral", "Journey", "GROWTH", "LEARNING", "Certifications & Learning");
    // 좌: 자격증
    await circleIcon(s, 0.5, BY + 0.1, 0.5, Fa.FaCertificate, NAVY, HEX.lt2, "자격증");
    txt(s, "자격증 (준비중)", { x: 1.15, y: BY + 0.1, w: 3.2, h: 0.5, fontSize: 16, bold: true, color: NAVY, valign: "middle", objectName: "cert head" });
    const certs = [["전산회계 1급", "한국세무사회 · 2026.06 목표"], ["전산세무 2급", "한국세무사회 · 2026.06 목표"]];
    certs.forEach(([name, meta], i) => {
      const y = BY + 0.8 + i * 1.25;
      s.addShape(pres.ShapeType.roundRect, { x: 0.5, y, w: 4.1, h: 1.05, rectRadius: 0.12, fill: { color: GOLD }, line: { color: GOLD, width: 0 }, objectName: `cert card ${i + 1}` });
      txt(s, name, { x: 0.75, y: y + 0.15, w: 3.6, h: 0.45, fontSize: 20, bold: true, color: NAVY, objectName: `cert ${i + 1} name` });
      txt(s, meta, { x: 0.75, y: y + 0.6, w: 3.6, h: 0.3, fontSize: 12, color: INK, objectName: `cert ${i + 1} meta` });
    });
    // 우: 최근 학습
    await circleIcon(s, 5.1, BY + 0.1, 0.5, Fa.FaBookOpen, CORAL, HEX.lt1, "최근 학습");
    txt(s, "최근 학습", { x: 5.75, y: BY + 0.1, w: 3.5, h: 0.5, fontSize: 16, bold: true, color: NAVY, valign: "middle", objectName: "learning head" });
    const learn = [["2026.04.16", "스톡옵션 과세 체계", "IR/투자"], ["2026.04.14", "스톡옵션 법률 구조", "IR/투자"], ["2026.04.10", "주주간계약(SHA) 핵심 조항", "IR/투자"], ["2026.04.08", "IT 스타트업 런웨이 관리", "재무"]];
    learn.forEach(([d, t, tag], i) => {
      const y = BY + 0.85 + i * 0.66;
      txt(s, d, { x: 5.1, y: y + 0.07, w: 1.15, h: 0.3, fontSize: 12, bold: true, color: CORAL, objectName: `learn ${i + 1} date` });
      txt(s, t, { x: 6.3, y: y + 0.02, w: 2.4, h: 0.45, fontSize: 14, color: INK, objectName: `learn ${i + 1} topic` });
      chip(s, 8.7, y + 0.06, 0.8, 0.3, tag, NAVY, WHITE, 9);
    });
    s.addNotes("Certifications & Learning");
  }

  // =====================================================================
  // 12. Project Gallery — 상 골드 블록, 3×2 카드
  // =====================================================================
  {
    const s = split("ST", "gold", "Journey", "GALLERY", "PROJECTS", "Project Gallery");
    const projects = [
      ["IR/투자", "SHA 주주간계약 검토", "ROFR, Drag-Along, Tag-Along, 희석방지 조항 분석 및 투자 계약 실무 적용", "투자 계약 핵심 조항 가이드 문서화"],
      ["IR/투자", "스톡옵션 체계 설계", "법률 구조 분석, 과세 체계 정립, 임직원 보상 설계 가이드라인 수립", "스톡옵션 세무/법률 통합 가이드 완성"],
      ["HR/노무", "급여 체계 설계", "TTC(Total Target Compensation) 기반 급여 체계 설계 및 보상 구조 수립", "전사 보상 체계 프레임워크 구축"],
      ["정부사업", "ABTC 인증 처리", "APEC 기업인 여행카드 관련 서류 준비 및 인증 프로세스 관리", "인증 완료 및 프로세스 SOP 문서화"],
      ["재무/회계", "런웨이 관리 체계", "IT 스타트업 현금흐름 기반 런웨이 관리 프레임워크 구축", "월별 런웨이 모니터링 대시보드 구축"],
      ["재무/회계", "월말 결산 프로세스", "매월 재무제표 마감 및 결산 프로세스 체계화, 전표 처리 자동화", "결산 소요 시간 단축 및 프로세스 표준화"],
    ];
    const cw = 2.9, ch = 1.65, gx = 0.15, gy = 0.15;
    const tagFill = { "IR/투자": CORAL, "HR/노무": NAVY, "정부사업": NAVY, "재무/회계": CORAL };
    projects.forEach(([tag, title, desc, outcome], i) => {
      const x = 0.5 + (i % 3) * (cw + gx), y = BY + Math.floor(i / 3) * (ch + gy);
      s.addShape(pres.ShapeType.rect, { x, y, w: cw, h: ch, fill: { color: TINT }, line: { color: TINT, width: 0 }, objectName: `project card ${i + 1}` });
      chip(s, x + 0.15, y + 0.13, 0.85, 0.26, tag, tagFill[tag], WHITE, 9);
      txt(s, title, { x: x + 0.15, y: y + 0.45, w: cw - 0.3, h: 0.3, fontSize: 14, bold: true, color: NAVY, objectName: `project ${i + 1} title` });
      txt(s, desc, { x: x + 0.15, y: y + 0.75, w: cw - 0.3, h: 0.5, fontSize: 10, color: INK, objectName: `project ${i + 1} desc` });
      txt(s, [{ text: "성과  ", options: { bold: true, color: CORAL } }, { text: outcome, options: { bold: true } }], { x: x + 0.15, y: y + 1.25, w: cw - 0.3, h: 0.36, fontSize: 10, color: NAVY, objectName: `project ${i + 1} outcome` });
    });
    s.addNotes("Project Gallery 6개");
  }

  // =====================================================================
  // 13. 마무리 — 코랄 풀블리드, 네이비 타이포
  // =====================================================================
  pres.addSection({ title: "Closing" });
  n += 1;
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Closing" });
    txt(s, `${String(n).padStart(2, "0")}   CONTACT`, { x: 0.6, y: 0.5, w: 4, h: 0.3, fontSize: 11, bold: true, color: NAVY, charSpacing: 3, objectName: "closing label" });
    txt(s, "Let's Connect", { x: 0.55, y: 1.0, w: 8.9, h: 1.4, fontSize: 84, bold: true, color: NAVY, objectName: "closing title" });
    txt(s, "새로운 기회나 협업에 대해 이야기하고 싶으시다면 연락주세요.", { x: 0.6, y: 2.55, w: 8.8, h: 0.5, fontSize: 18, color: INK, objectName: "closing body" });
    const links = [["Email", Fa.FaEnvelope], ["LinkedIn", Fa.FaLinkedin]];
    for (let i = 0; i < links.length; i++) {
      const [name, Icon] = links[i];
      const x = 0.6 + i * 2.6;
      await circleIcon(s, x, 3.35, 0.6, Icon, NAVY, HEX.accent3, name);
      txt(s, name, { x: x + 0.75, y: 3.35, w: 1.7, h: 0.6, fontSize: 18, bold: true, color: NAVY, valign: "middle", objectName: `${name} label` });
    }
    s.addShape(pres.ShapeType.rect, { x: 0, y: 4.75, w: W, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "footer block" });
    txt(s, "© 2026 양남규", { x: 0.6, y: 4.75, w: 8.8, h: 0.85, fontSize: 12, bold: true, color: GOLD, valign: "middle", objectName: "copyright" });
    s.addNotes("마무리 — Let's Connect");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT, "slides:", n);
})().catch((e) => { console.error(e); process.exit(1); });
