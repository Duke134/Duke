// Bold Typo — 양남규 경영지원 프로필 덱 v2 (14장)
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
  pres.title = "양남규 경영지원 프로필 — Bold Typo";
  const C = pres.SchemeColor;

  const W = 10, H = 5.625;
  const NAVY = C.text2, CORAL = C.accent1, GOLD = C.background2, WHITE = C.background1, INK = C.text1, TINT = C.accent5;

  // Block color roles: which text colors sit on each block fill
  const ROLE = {
    navy: { fill: NAVY, label: GOLD, kw: GOLD, sub: WHITE, body: WHITE, soft: GOLD, iconFill: GOLD, iconHex: HEX.dk2 },
    coral: { fill: CORAL, label: NAVY, kw: NAVY, sub: INK, body: INK, soft: NAVY, iconFill: NAVY, iconHex: HEX.lt2 },
    gold: { fill: GOLD, label: CORAL, kw: NAVY, sub: INK, body: INK, soft: NAVY, iconFill: NAVY, iconHex: HEX.lt2 },
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
    // left/right split — case variant (long title under the number)
    pres.defineSlideMaster({
      title: `SLC_${k}`,
      background: { color: HEX.lt1 },
      objects: [
        { rect: { x: 0, y: 0, w: LB_W, h: H, fill: { color: r.fill } } },
        { placeholder: { options: { name: "label", type: "body", x: 0.35, y: 0.4, w: 2.9, h: 0.3, fontSize: 11, bold: true, color: r.label, margin: 0, charSpacing: 2 }, text: "" } },
        { placeholder: { options: { name: "kw", type: "title", x: 0.3, y: 0.7, w: 3.0, h: 1.05, fontSize: 60, bold: true, color: r.kw, margin: 0, align: "left", valign: "top" }, text: "" } },
        { placeholder: { options: { name: "sub", type: "body", x: 0.35, y: 1.8, w: 2.7, h: 1.7, fontSize: 18, bold: true, color: r.sub, margin: 0, valign: "top", lineSpacingMultiple: 1.15 }, text: "" } },
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
        { placeholder: { options: { name: "sub", type: "body", x: 4.7, y: 0.5, w: 4.8, h: 0.95, fontSize: 20, bold: true, color: r.sub, margin: 0, align: "right", valign: "bottom" }, text: "" } },
      ],
    });
  }

  // ---------- helpers ----------
  let n = 0;
  function split(kind, role, section, label, kw, sub) {
    n += 1;
    const s = pres.addSlide({ masterName: `${kind}_${role}`, sectionTitle: section });
    s.addText(`${String(n).padStart(2, "0")}   ${label}`, { placeholder: "label" });
    s.addText(kw, { placeholder: "kw" });
    s.addText(sub, { placeholder: "sub" });
    return s;
  }
  async function circleIcon(s, x, y, d, Icon, circleFill, iconHex, name) {
    s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: circleFill }, line: { color: circleFill, width: 0 }, objectName: `${name} circle` });
    const pad = d * 0.26;
    s.addImage({ data: await iconData(Icon, iconHex), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name });
  }
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", align: "left" }, o));
  // step flow: chevrons in a row (first is a homePlate)
  function flow(s, x, y, w, h, steps, name, size = 12) {
    const gap = 0.08;
    const sw = (w - gap * (steps.length - 1)) / steps.length;
    steps.forEach((t, i) => {
      const fill = i % 2 === 0 ? NAVY : CORAL;
      s.addText(t, {
        x: x + i * (sw + gap), y, w: sw, h,
        shape: i === 0 ? pres.ShapeType.homePlate : pres.ShapeType.chevron,
        fill: { color: fill }, line: { color: fill, width: 0 },
        color: WHITE, fontSize: size, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true,
        objectName: `${name} step ${i + 1}`,
      });
    });
  }
  const labelOpts = (color) => ({ fontSize: 11, bold: true, color, charSpacing: 2 });

  // =====================================================================
  // 1. 표지 — 코랄 풀블리드, 네이비 타이포
  // =====================================================================
  pres.addSection({ title: "Intro" });
  n += 1;
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "Intro" });
    txt(s, "경영지원 매니저", { x: 0.6, y: 0.65, w: 6, h: 0.5, fontSize: 24, bold: true, color: NAVY, objectName: "kicker" });
    txt(s, "양남규", { x: 0.55, y: 1.05, w: 8.9, h: 1.75, fontSize: 112, bold: true, color: NAVY, objectName: "name" });
    txt(s, "Duke Yang", { x: 0.6, y: 2.85, w: 8.8, h: 0.5, fontSize: 28, bold: true, color: INK, objectName: "english name" });
    txt(s, "IT 스타트업에서 재무·인사·IR·정부사업·해외업무·총무를 혼자 맡아 운영하며, 그 일을 AI와 지식관리 도구로 시스템화하는 경영지원 매니저입니다.", {
      x: 0.6, y: 3.45, w: 8.6, h: 1.0, fontSize: 16, color: INK, lineSpacingMultiple: 1.2, objectName: "tagline",
    });
    s.addShape(pres.ShapeType.rect, { x: 0, y: 4.75, w: W, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "footer block" });
    txt(s, "기준일 2026-10-27", { x: 0.6, y: 4.75, w: 8.8, h: 0.85, fontSize: 14, bold: true, color: GOLD, valign: "middle", objectName: "meta" });
    s.addNotes("표지: 양남규 (Duke Yang), 경영지원 매니저");
  }

  // =====================================================================
  // 2. 일하는 원칙 — 좌 네이비 블록, 큰 키워드 3개 + 문장
  // =====================================================================
  {
    const s = split("SL", "navy", "Intro", "PRINCIPLES", "원칙", "일하는 원칙");
    const rows = [
      ["시스템", "사람의 기억이 아닌 시스템에 남긴다.", Fa.FaDatabase, CORAL, HEX.lt1],
      ["원문 확인", "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보낸다.", Fa.FaSearch, NAVY, HEX.lt2],
      ["스킬 · SOP", "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다.", Fa.FaCogs, GOLD, HEX.dk2],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [kw, sentence, Icon, fill, iconHex] = rows[i];
      const y = 0.55 + i * 1.55;
      await circleIcon(s, RX, y + 0.05, 0.7, Icon, fill, iconHex, `principle ${i + 1}`);
      txt(s, `0${i + 1}`, { x: RX + 0.95, y: y - 0.02, w: 1.0, h: 0.25, ...labelOpts(CORAL), objectName: `principle ${i + 1} index` });
      txt(s, kw, { x: RX + 0.95, y: y + 0.18, w: RW - 0.95, h: 0.55, fontSize: 30, bold: true, color: NAVY, objectName: `principle ${i + 1} keyword` });
      txt(s, sentence, { x: RX + 0.95, y: y + 0.78, w: RW - 0.95, h: 0.65, fontSize: 15, color: INK, lineSpacingMultiple: 1.15, objectName: `principle ${i + 1} sentence` });
    }
    s.addNotes("일하는 원칙 3가지");
  }

  // =====================================================================
  // 3. 소개 — 상 코랄 블록, 2×2 문단 + 아이콘
  // =====================================================================
  {
    const s = split("ST", "coral", "Intro", "ABOUT", "소개", "About Me");
    const paras = [
      ["IT 스타트업 경영지원팀 매니저로 일하며 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다.", Fa.FaBriefcase],
      ["국제통상학을 전공했고, 미국 현지 유통기업에서 근무한 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다.", Fa.FaGlobeAsia],
      ["여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해 옵시디언 기반 업무 운영체계와 AI 자동화 워크플로를 직접 설계해 쓰고 있습니다.", Fa.FaSitemap],
      ["받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다.", Fa.FaCheckDouble],
    ];
    for (let i = 0; i < paras.length; i++) {
      const [p, Icon] = paras[i];
      const x = 0.5 + (i % 2) * 4.6, y = BY + 0.15 + Math.floor(i / 2) * 1.7;
      const fill = i % 2 === 0 ? NAVY : CORAL;
      await circleIcon(s, x, y, 0.55, Icon, fill, i % 2 === 0 ? HEX.lt2 : HEX.lt1, `about ${i + 1}`);
      txt(s, p, { x: x + 0.75, y: y - 0.03, w: 3.65, h: 1.5, fontSize: 14, color: INK, lineSpacingMultiple: 1.2, objectName: `about paragraph ${i + 1}` });
    }
    s.addNotes("소개 4문단");
  }

  // =====================================================================
  // 4. 담당 영역 — 좌 골드 블록, 아이콘 카드 6개
  // =====================================================================
  pres.addSection({ title: "Areas" });
  {
    const s = split("SL", "gold", "Areas", "AREAS", "SIX", "담당 영역");
    txt(s, "여섯 영역을 한 사람이 맡아 운영합니다", { x: 0.35, y: 2.6, w: 2.7, h: 0.8, fontSize: 15, color: INK, lineSpacingMultiple: 1.2, objectName: "areas lead" });
    const areas = [
      ["재무·회계", "자금 운용, 세무 전 주기, 결산·외부감사", Fa.FaCalculator],
      ["인사·노무", "급여·4대보험, 채용, 해외 원격 인력 계약", Fa.FaUsers],
      ["IR·투자", "주총·이사회, 등기, 스톡옵션, 투자 계약", Fa.FaHandshake],
      ["정부사업", "공고 분석부터 정산까지 전 주기", Fa.FaLandmark],
      ["해외업무", "해외 자회사 증빙, 해외 행사·출장", Fa.FaPlane],
      ["총무", "법인 행정, 기업 인증, 사무실·자산, 대외행사", Fa.FaBuilding],
    ];
    const cw = 2.85, ch = 1.45, gx = 0.15, gy = 0.13;
    for (let i = 0; i < areas.length; i++) {
      const [head, sum, Icon] = areas[i];
      const x = RX + (i % 2) * (cw + gx), y = 0.45 + Math.floor(i / 2) * (ch + gy);
      s.addShape(pres.ShapeType.rect, { x, y, w: cw, h: ch, fill: { color: TINT }, line: { color: TINT, width: 0 }, objectName: `area card ${i + 1}` });
      const fill = i % 2 === 0 ? NAVY : CORAL;
      await circleIcon(s, x + 0.18, y + 0.17, 0.52, Icon, fill, i % 2 === 0 ? HEX.lt2 : HEX.lt1, head);
      txt(s, `0${i + 1}`, { x: x + cw - 0.6, y: y + 0.15, w: 0.45, h: 0.25, ...labelOpts(CORAL), align: "right", objectName: `area ${i + 1} index` });
      txt(s, head, { x: x + 0.85, y: y + 0.17, w: cw - 1.45, h: 0.52, fontSize: 16, bold: true, color: NAVY, valign: "middle", objectName: `area ${i + 1} head` });
      txt(s, sum, { x: x + 0.18, y: y + 0.82, w: cw - 0.36, h: 0.55, fontSize: 13, color: INK, lineSpacingMultiple: 1.15, objectName: `area ${i + 1} summary` });
    }
    s.addNotes("담당 영역 6개");
  }

  // =====================================================================
  // 5. 담당 영역 상세 1 — 상 네이비 블록, 3열
  // =====================================================================
  {
    const s = split("ST", "navy", "Areas", "DETAIL 01", "WORK 01", "재무·회계 / 인사·노무 / IR·투자");
    const cols = [
      ["재무·회계", ["자금 운용과 송금·결제를 관리하고, 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응합니다.", "결산과 외부감사 대응, 별도·연결 재무제표 관리를 맡고 외부 회계법인과의 커뮤니케이션을 담당합니다."], Fa.FaCalculator],
      ["인사·노무", ["급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다.", "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다."], Fa.FaUsers],
      ["IR·투자", ["정기주주총회·이사회 운영, 정관 개정, 등기, 주주명부와 스톡옵션 관리를 맡고 있습니다.", "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다."], Fa.FaHandshake],
    ];
    const cw = 2.85, gap = (9 - 3 * cw) / 2;
    for (let i = 0; i < cols.length; i++) {
      const [head, paras, Icon] = cols[i];
      const x = 0.5 + i * (cw + gap);
      const fill = i === 1 ? CORAL : NAVY;
      await circleIcon(s, x, BY + 0.05, 0.5, Icon, fill, i === 1 ? HEX.lt1 : HEX.lt2, head);
      txt(s, head, { x: x + 0.65, y: BY + 0.05, w: cw - 0.65, h: 0.5, fontSize: 18, bold: true, color: NAVY, valign: "middle", objectName: `detail ${i + 1} head` });
      txt(s, paras.map((t, j) => ({ text: t, options: { breakLine: j < paras.length - 1, paraSpaceAfter: 8 } })), {
        x, y: BY + 0.75, w: cw, h: 2.9, fontSize: 13, color: INK, lineSpacingMultiple: 1.15, objectName: `detail ${i + 1} body`,
      });
    }
    s.addNotes("담당 영역 상세 1 — 재무·회계 / 인사·노무 / IR·투자");
  }

  // =====================================================================
  // 6. 담당 영역 상세 2 — 상 코랄 블록, 5단계 흐름 + 2열
  // =====================================================================
  {
    const s = split("ST", "coral", "Areas", "DETAIL 02", "WORK 02", "정부사업 / 해외업무 / 총무");
    // 정부사업 — 5단계 흐름
    await circleIcon(s, 0.5, BY + 0.05, 0.45, Fa.FaLandmark, NAVY, HEX.lt2, "정부사업");
    txt(s, "정부사업", { x: 1.1, y: BY + 0.05, w: 4, h: 0.45, fontSize: 18, bold: true, color: NAVY, valign: "middle", objectName: "gov head" });
    flow(s, 0.5, BY + 0.62, 9.0, 0.48, ["공고 분석", "사업계획서 작성", "신청", "협약", "정산"], "gov");
    txt(s, "공고 분석 → 사업계획서 작성 → 신청 → 협약 → 정산까지 전 주기를 실무로 수행합니다. 입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다.", {
      x: 0.5, y: BY + 1.2, w: 9.0, h: 0.6, fontSize: 13, color: INK, lineSpacingMultiple: 1.15, objectName: "gov body",
    });
    // 해외업무 / 총무 — 2열
    const cols = [
      ["해외업무", "해외 자회사의 인보이스와 용역계약 증빙을 관리하고, 해외 사업 행사 참가와 해외 출장 서류 실무를 지원합니다.", Fa.FaPlane, CORAL, HEX.lt1],
      ["총무", "본점 이전 같은 법인 행정 절차와 기업 인증 관리, 사무실·자산 관리를 담당합니다. 대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다.", Fa.FaBuilding, NAVY, HEX.lt2],
    ];
    for (let i = 0; i < cols.length; i++) {
      const [head, body, Icon, fill, iconHex] = cols[i];
      const x = 0.5 + i * 4.6, y = BY + 1.95;
      await circleIcon(s, x, y, 0.45, Icon, fill, iconHex, head);
      txt(s, head, { x: x + 0.6, y, w: 3.8, h: 0.45, fontSize: 18, bold: true, color: NAVY, valign: "middle", objectName: `${head} head` });
      txt(s, body, { x, y: y + 0.55, w: 4.4, h: 1.2, fontSize: 13, color: INK, lineSpacingMultiple: 1.15, objectName: `${head} body` });
    }
    s.addNotes("담당 영역 상세 2 — 정부사업(5단계 흐름) / 해외업무 / 총무");
  }

  // =====================================================================
  // 7-11. 사례 1~5 — 좌 블록 (네이비/코랄/골드 순환), 배경 / 한 일 / 결과
  // =====================================================================
  pres.addSection({ title: "Cases" });
  const cases = [
    {
      role: "navy", num: "01", title: "1인 경영지원을 \"인계 가능한 시스템\"으로 바꾸기", Icon: Fa.FaSitemap,
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 커스텀 AI 스킬로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다.",
    },
    {
      role: "coral", num: "02", title: "단일 통화 자금 시트를 의사결정 도구로 재설계", Icon: Fa.FaCoins,
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리하고 있어, 경영진의 자금 판단 질문에 바로 답하기 어려웠습니다.",
      did: "대시보드·통화별 입출금·정기 비용·미지급 항목으로 시트를 나누고, 다중 통화 환산과 \"부족액 식별 → 송금액 계산 → 송금 후 잔액\" 시뮬레이션을 자동화했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.",
      flow: ["부족액 식별", "송금액 계산", "송금 후 잔액"],
      result: "경영진의 자금 의사결정 질문에 한 화면에서 바로 답할 수 있는 상태가 됐습니다.",
    },
    {
      role: "gold", num: "03", title: "법률 1차 검토를 법령 원문 기반으로 내재화", Icon: Fa.FaBalanceScale,
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했습니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 실무 상식만으로는 놓칠 수 있는 별도 조항을 미리 잡게 됐습니다.",
    },
    {
      role: "navy", num: "04", title: "외부로 나가기 전 \"받은 산출물\"을 역검증하는 습관", Icon: Fa.FaCheckDouble,
      bg: "회계법인·기관·플랫폼에서 받은 자료나 내부 문서가 확인 없이 외부로 나가면 세금·등기·평가에 바로 영향을 줍니다.",
      did: "회계법인 산출물, 기관·플랫폼 자료, 계약서와 내부 문서를 원장·법령 원문과 대조한 뒤 제출하는 절차를 분기 내내 운영했습니다.",
      result: "기간 오류, 견적 오기, 정관과 계약의 충돌 같은 문제를 외부 제출 전에 찾아냈습니다.",
    },
    {
      role: "coral", num: "05", title: "예산 대비 실적 관리 워크북 구축", Icon: Fa.FaTable,
      bg: "월별 지출 추적이 여러 엑셀 파일에 흩어져 있어 예산을 통제하기 어려웠습니다.",
      did: "실제 계정원장을 기준으로 계정과목을 정리해 카테고리로 묶었습니다. 월별 예산 수립 → 실적 입력 → 차이 분석 흐름을 자동화했고, 이후 중앙 가정 시트와 연동되도록 확장했습니다.",
      flow: ["예산 수립", "실적 입력", "차이 분석"],
      result: "예산 대비 실적을 한 체계에서 관리하는 관리회계 기반을 만들었습니다.",
    },
  ];
  for (const c of cases) {
    const r = ROLE[c.role];
    const s = split("SLC", c.role, "Cases", `CASE ${c.num}`, c.num, c.title);
    // 좌 블록 하단: 사례 아이콘
    await circleIcon(s, 0.35, 4.2, 0.95, c.Icon, r.iconFill, r.iconHex, `case ${c.num} icon`);
    // 우측: 배경 / 한 일 / 결과 — 본문 길이에 맞춰 세로 위치를 계산
    const LH = 0.27, lines = (t, cpl) => Math.ceil(t.length / cpl);
    const bgL = lines(c.bg, 36), didL = lines(c.did, 36), resL = lines(c.result, 33);
    let y = 0.45;
    txt(s, "배경", { x: RX, y, w: 2, h: 0.25, ...labelOpts(CORAL), objectName: "background label" });
    txt(s, c.bg, { x: RX, y: y + 0.27, w: RW, h: bgL * LH + 0.08, fontSize: 14, color: INK, lineSpacingMultiple: 1.15, objectName: "background" });
    y += 0.27 + bgL * LH + 0.32;
    txt(s, "한 일", { x: RX, y, w: 2, h: 0.25, ...labelOpts(CORAL), objectName: "did label" });
    txt(s, c.did, { x: RX, y: y + 0.27, w: RW, h: didL * LH + 0.08, fontSize: 14, color: INK, lineSpacingMultiple: 1.15, objectName: "did" });
    y += 0.27 + didL * LH + 0.08;
    if (c.flow) { flow(s, RX, y + 0.22, RW, 0.42, c.flow, `case ${c.num}`, 12); y += 0.22 + 0.42; }
    // 결과 — 골드 틴트 박스
    const boxH = resL * LH + 0.72;
    const boxY = Math.min(y + 0.32, 5.3 - boxH);
    s.addShape(pres.ShapeType.roundRect, { x: RX, y: boxY, w: RW, h: boxH, rectRadius: 0.1, fill: { color: GOLD }, line: { color: GOLD, width: 0 }, objectName: "result box" });
    txt(s, "결과", { x: RX + 0.25, y: boxY + 0.17, w: 2, h: 0.25, ...labelOpts(NAVY), objectName: "result label" });
    txt(s, c.result, { x: RX + 0.25, y: boxY + 0.45, w: RW - 0.5, h: resL * LH + 0.08, fontSize: 14, bold: true, color: INK, lineSpacingMultiple: 1.15, objectName: "result" });
    s.addNotes(`사례 ${c.num} — ${c.title}`);
  }

  // =====================================================================
  // 12. 스킬·도구 — 상 골드 블록, 2열 리스트 8행
  // =====================================================================
  pres.addSection({ title: "Skills & Journey" });
  {
    const s = split("ST", "gold", "Skills & Journey", "SKILLS", "SKILLS", "스킬 · 도구");
    const skills = [
      ["엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·자금 시뮬레이션 설계", Fa.FaFileExcel],
      ["옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인", Fa.FaBook],
      ["AI 업무 자동화", "Claude 커스텀 스킬 설계·운영, 예약 작업 자동화, MCP 연동(법령 원문 조회, 옵시디언 연동)", Fa.FaRobot],
      ["협업 도구", "Gmail, Google Drive, Slack", Fa.FaComments],
      ["ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)", Fa.FaServer],
      ["문서 표준화", "공문 서식, docx·pptx 폰트 표준, 인보이스·인사·급여 템플릿, SOP", Fa.FaFileAlt],
      ["실무 지식", "결산·세무 신고 대응, 주총·이사회·등기, 정부사업 전 주기", Fa.FaBalanceScale],
      ["언어", "비즈니스 영어(미국 현지 근무 경험)", Fa.FaLanguage],
    ];
    const colW = 4.4, pitch = 0.9;
    for (let i = 0; i < skills.length; i++) {
      const [name, desc, Icon] = skills[i];
      const x = 0.5 + Math.floor(i / 4) * 4.6, y = BY + (i % 4) * pitch;
      const fill = i % 2 === 0 ? NAVY : CORAL;
      await circleIcon(s, x, y, 0.36, Icon, fill, i % 2 === 0 ? HEX.lt2 : HEX.lt1, name);
      txt(s, name, { x: x + 0.55, y: y - 0.02, w: colW - 0.55, h: 0.34, fontSize: 14, bold: true, color: NAVY, valign: "middle", objectName: `skill ${i + 1} name` });
      txt(s, desc, { x: x + 0.55, y: y + 0.35, w: colW - 0.55, h: 0.42, fontSize: 11, color: INK, lineSpacingMultiple: 1.1, objectName: `skill ${i + 1} desc` });
    }
    s.addNotes("스킬·도구 8행");
  }

  // =====================================================================
  // 13. 타임라인 — 좌 네이비 블록, 세로 타임라인 (최신이 위)
  // =====================================================================
  {
    const s = split("SL", "navy", "Skills & Journey", "TIMELINE", "경력", "타임라인 2013 – 현재");
    await circleIcon(s, 0.35, 4.2, 0.95, Fa.FaRoute, GOLD, HEX.dk2, "timeline icon");
    const ev = [
      ["2022 – 현재", "IT 스타트업", "경영지원팀 매니저"],
      ["2019 – 2022", "미국 소재 유통기업", "재고관리 인턴 → 구매(Purchasing) 담당"],
      ["2018 – 2019", "미국 비즈니스 취업연수 과정 수료", ""],
      ["2017", "중국 교환학생", ""],
      ["2013 – 2019", "학부", "국제통상학 전공 (중어중문학 복수전공)"],
    ];
    const rowH = 0.95, y0 = 0.55, lineX = 5.0;
    s.addShape(pres.ShapeType.line, { x: lineX, y: y0 + 0.1, w: 0, h: rowH * 4, line: { color: NAVY, width: 1.5 }, objectName: "timeline line" });
    ev.forEach(([date, title, desc], i) => {
      const y = y0 + i * rowH;
      txt(s, date, { x: RX, y: y + 0.02, w: 1.25, h: 0.3, fontSize: 12, bold: true, color: CORAL, objectName: `event ${i + 1} date` });
      s.addShape(pres.ShapeType.ellipse, { x: lineX - 0.1, y: y + 0.01, w: 0.2, h: 0.2, fill: { color: i === 0 ? CORAL : NAVY }, line: { color: WHITE, width: 1.5 }, objectName: `event ${i + 1} dot` });
      txt(s, title, { x: 5.3, y: y - 0.04, w: 4.2, h: 0.35, fontSize: 16, bold: true, color: NAVY, objectName: `event ${i + 1} title` });
      if (desc) txt(s, desc, { x: 5.3, y: y + 0.33, w: 4.2, h: 0.35, fontSize: 13, color: INK, objectName: `event ${i + 1} desc` });
    });
    s.addNotes("타임라인 — 최신이 위");
  }

  // =====================================================================
  // 14. 연락 — 코랄 풀블리드, 네이비 타이포
  // =====================================================================
  pres.addSection({ title: "Closing" });
  n += 1;
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "Closing" });
    txt(s, `${String(n).padStart(2, "0")}   CONTACT`, { x: 0.6, y: 0.5, w: 4, h: 0.3, fontSize: 11, bold: true, color: NAVY, charSpacing: 3, objectName: "closing label" });
    txt(s, "연락", { x: 0.55, y: 0.85, w: 8.9, h: 1.5, fontSize: 96, bold: true, color: NAVY, objectName: "closing title" });
    txt(s, "양남규  ·  Duke Yang  ·  경영지원 매니저", { x: 0.6, y: 2.5, w: 8.8, h: 0.5, fontSize: 20, bold: true, color: INK, objectName: "closing name" });
    await circleIcon(s, 0.6, 3.3, 0.65, Fa.FaEnvelope, NAVY, HEX.accent3, "이메일");
    txt(s, "namgyu1412@naver.com", { x: 1.45, y: 3.3, w: 7.5, h: 0.65, fontSize: 24, bold: true, color: NAVY, valign: "middle", objectName: "email" });
    s.addShape(pres.ShapeType.rect, { x: 0, y: 4.75, w: W, h: 0.85, fill: { color: NAVY }, line: { color: NAVY, width: 0 }, objectName: "footer block" });
    txt(s, "© 2026 Duke Yang", { x: 0.6, y: 4.75, w: 8.8, h: 0.85, fontSize: 12, bold: true, color: GOLD, valign: "middle", objectName: "copyright" });
    s.addNotes("연락 — namgyu1412@naver.com");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT, "slides:", n);
})().catch((e) => { console.error(e); process.exit(1); });
