// 03-editorial — "Editorial" deck for 양남규 portfolio (content v2, 14 slides)
// Run: NODE_PATH=<node_modules> node build.js   (from any cwd; paths are relative to this file)
const path = require("path");
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
pres.title = "양남규 경영지원 프로필 — Editorial";
const C = pres.SchemeColor;
const INKc = C.text1, BERRYc = C.text2, WHITEc = C.background1, CREAMc = C.background2, ROSEc = C.accent2;

// ---------- geometry ----------
const BLOCK_W = 3.33;            // left 1/3 color block (motif)
const BX = 3.7, BW = 5.85;       // right body column
const BY = 1.5, BB = 5.2;        // default body top / bottom
const TITLE_Y = 0.3, TITLE_PT = 36;
const BLK_X = 0.4, BLK_W = 2.55, BLK_Y = 1.95; // text area inside the block

// ---------- icons ----------
async function icon(Component, hex, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Component, { color: "#" + hex, size: px }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- text metrics (inches), tuned for Noto Sans CJK / Malgun Gothic with slack ----------
function charW(ch, pt) {
  const em = pt / 72;
  const c = ch.codePointAt(0);
  if (c >= 0x1100 && (c <= 0x11ff || (c >= 0x3130 && c <= 0x318f) || (c >= 0xac00 && c <= 0xd7af) || (c >= 0x4e00 && c <= 0x9fff) || (c >= 0x3000 && c <= 0x30ff) || c === 0x2192)) return 1.0 * em;
  if (c === 0x201c || c === 0x201d) return 0.22 * em;
  if (ch === " ") return 0.2 * em;
  if (/[A-Z]/.test(ch)) return 0.68 * em;
  if (/[0-9]/.test(ch)) return 0.58 * em;
  if (/[a-z]/.test(ch)) return 0.54 * em;
  return 0.36 * em;
}
function estW(text, pt, slack = 1.04) {
  let w = 0;
  for (const ch of text) w += charW(ch, pt);
  return w * slack;
}
// greedy word wrap estimate (break at spaces, as PowerPoint/LibreOffice do for Korean with spaces)
function lines(text, pt, w, slack = 1.04) {
  let n = 1, cur = 0;
  const sp = charW(" ", pt) * slack;
  for (const word of text.split(" ")) {
    const ww = estW(word, pt, slack);
    if (cur === 0) cur = ww;
    else if (cur + sp + ww <= w) cur += sp + ww;
    else { n++; cur = ww; }
  }
  return n;
}
const LS = 1.15;                                   // body lineSpacingMultiple
const lh = (pt) => (pt / 72) * 1.2 * LS;           // rendered line height
function paraH(text, pt, w) { return lines(text, pt, w) * lh(pt) + 0.04; }

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
            name: "title", type: "title", x: BX, y: TITLE_Y, w: BW, h: 1.15,
            fontSize: TITLE_PT, bold: true, color: INKc, align: "left", valign: "top", margin: 0,
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
function RR(slide, o) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ rectRadius: 0.08, line: { color: WHITE, width: 0 }, objectName: "card " + (++shapeN) }, o));
}
function OVAL(slide, o) {
  slide.addShape(pres.shapes.OVAL, Object.assign({ line: { color: WHITE, width: 0 }, objectName: "circle " + (++shapeN) }, o));
}
function HLINE(slide, x, y, w, color, width = 0.75) {
  slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width }, objectName: "rule " + (++shapeN) });
}
// Variant palettes: text colors inside the block
const V = {
  berry: { big: CREAMc, lab: CREAMc, head: CREAMc, body: WHITEc, quote: ROSEc, name: "CONTENT_BERRY" },
  cream: { big: ROSEc, lab: BERRYc, head: BERRYc, body: INKc, quote: ROSEc, name: "CONTENT_CREAM" },
};
// The motif: left block with big number + small label
function block(slide, num, label, variant) {
  const v = V[variant];
  T(slide, num, { x: 0.35, y: 0.28, w: 2.7, h: 1.15, fontSize: 72, bold: true, color: v.big, valign: "top" });
  T(slide, label, { x: 0.38, y: 1.42, w: 2.7, h: 0.25, fontSize: 10, bold: true, charSpacing: 3, color: v.lab });
}
function bigQuote(slide, x, y, color, pt = 80) {
  T(slide, "“", { x, y, w: pt / 72 * 1.2, h: pt / 72 * 1.3, fontSize: pt, bold: true, color });
}
function header(slide, text, x, y, w, color = ROSEc) {
  T(slide, text, { x, y, w, h: 0.22, fontSize: 11, bold: true, charSpacing: 2, color });
}
// Title in the placeholder; returns the body top for this slide (titles may wrap to 2–3 lines)
function title(slide, text) {
  slide.addText(text, { placeholder: "title" });
  const n = lines(text, TITLE_PT, BW, 1.0);
  return Math.max(BY, TITLE_Y + n * (TITLE_PT / 72) * 1.2 + 0.2);
}
// Header + paragraphs; returns bottom y
function section(slide, head, paras, x, y, w, pt, colors) {
  header(slide, head, x, y, w, colors.head);
  const h = paras.reduce((a, p) => a + paraH(p, pt, w), 0) + (paras.length - 1) * (8 / 72);
  T(slide, paras.map((p, i) => ({ text: p, options: { breakLine: i < paras.length - 1 } })),
    { x, y: y + 0.3, w, h, fontSize: pt, color: colors.body, lineSpacingMultiple: LS, paraSpaceAfter: 8 });
  return y + 0.3 + h;
}
// Step flow: chips joined by arrows; returns bottom y
function flow(slide, steps, x, y, maxW, opts) {
  const { pt = 11, h = 0.4, fill = BERRY, color = WHITEc, arrow = ROSEc, pad = 0.26, gap = 0.3 } = opts || {};
  let widths = steps.map((s) => estW(s, pt) + pad);
  const total = widths.reduce((a, b) => a + b, 0) + gap * (steps.length - 1);
  if (total > maxW) widths = widths.map((w) => (w * (maxW - gap * (steps.length - 1))) / (total - gap * (steps.length - 1)));
  let cx = x;
  steps.forEach((s, i) => {
    RR(slide, { x: cx, y, w: widths[i], h, fill: { color: fill }, rectRadius: h / 2 });
    T(slide, s, { x: cx, y, w: widths[i], h, fontSize: pt, bold: true, color, align: "center", valign: "middle" });
    cx += widths[i];
    if (i < steps.length - 1) {
      T(slide, "→", { x: cx, y, w: gap, h, fontSize: pt + 1, bold: true, color: arrow, align: "center", valign: "middle" });
      cx += gap;
    }
  });
  return y + h;
}

async function build() {
  // ============ 1. 표지 ============
  pres.addSection({ title: "표지" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
    T(s, "“", { x: 0.55, y: -0.15, w: 2.5, h: 2.2, fontSize: 200, bold: true, color: ROSEc, valign: "top" });
    T(s, "경영지원 매니저", { x: 0.8, y: 1.7, w: 5, h: 0.4, fontSize: 18, bold: true, color: CREAMc, charSpacing: 1 });
    T(s, "양남규", { x: 0.8, y: 2.05, w: 6, h: 1.1, fontSize: 64, bold: true, color: WHITEc, valign: "middle" });
    T(s, "Duke Yang", { x: 0.8, y: 3.15, w: 6, h: 0.45, fontSize: 22, bold: true, color: CREAMc });
    T(s, "IT 스타트업에서 재무·인사·IR·정부사업·해외업무·총무를 혼자 맡아 운영하며, 그 일을 AI와 지식관리 도구로 시스템화하는 경영지원 매니저입니다.",
      { x: 0.8, y: 3.7, w: 6.3, h: 0.95, fontSize: 14, color: WHITEc, lineSpacingMultiple: 1.25 });
    T(s, "기준일 2026-10-27", { x: 0.8, y: 4.95, w: 4, h: 0.3, fontSize: 11, color: CREAMc, charSpacing: 1 });
    T(s, "”", { x: 7.4, y: 3.3, w: 2.3, h: 2.4, fontSize: 200, bold: true, color: ROSEc, align: "right", valign: "bottom" });
  }

  // ============ 2. 일하는 원칙 ============
  pres.addSection({ title: "원칙·소개" });
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "원칙·소개" });
    block(s, "02", "PRINCIPLES", "berry");
    const top = title(s, "일하는 원칙");
    bigQuote(s, 0.32, 3.55, ROSEc, 150);
    const items = [
      ["시스템에 남긴다", "사람의 기억이 아닌 시스템에 남긴다."],
      ["확인한 뒤 내보낸다", "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보낸다."],
      ["스킬로, SOP로 만든다", "반복 업무는 스킬로, 재발하는 업무는 SOP로 만든다."],
    ];
    const pitch = 1.3;
    items.forEach(([kw, sent], i) => {
      const y = top + i * pitch;
      T(s, String(i + 1).padStart(2, "0"), { x: BX, y: y - 0.08, w: 0.95, h: 0.7, fontSize: 40, bold: true, color: ROSEc });
      T(s, kw, { x: BX + 1.0, y, w: BW - 1.0, h: 0.42, fontSize: 22, bold: true, color: BERRYc, valign: "top" });
      T(s, sent, { x: BX + 1.0, y: y + 0.48, w: BW - 1.0, h: 0.6, fontSize: 14, color: INKc, lineSpacingMultiple: LS });
    });
  }

  // ============ 3. 소개 ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "원칙·소개" });
    block(s, "03", "ABOUT", "cream");
    const top = title(s, "소개");
    const v = V.cream;
    bigQuote(s, 0.32, 1.75, v.quote, 72);
    const p1 = "IT 스타트업 경영지원팀 매니저로 일하며 국내 법인과 해외 자회사의 재무·인사·법무·정부사업·IR 실무를 단독으로 운영하고 있습니다.";
    T(s, p1, { x: BLK_X, y: 2.75, w: BLK_W, h: paraH(p1, 14, BLK_W), fontSize: 14, bold: true, color: v.body, lineSpacingMultiple: LS });
    const paras = [
      "국제통상학을 전공했고, 미국 현지 유통기업에서 근무한 경험을 바탕으로 해외 조직과 함께 일하는 데 익숙합니다.",
      "여러 영역이 맞물리는 경영지원 업무를 사람의 기억이 아닌 시스템에 남기기 위해 옵시디언 기반 업무 운영체계와 AI 자동화 워크플로를 직접 설계해 쓰고 있습니다.",
      "받은 자료를 그대로 쓰지 않고, 법령 원문과 원자료로 확인한 뒤 내보내는 것을 일하는 원칙으로 삼고 있습니다.",
    ];
    let cy = top + 0.1;
    for (const p of paras) {
      const h = paraH(p, 15, BW);
      T(s, p, { x: BX, y: cy, w: BW, h, fontSize: 15, color: INKc, lineSpacingMultiple: LS });
      cy += h + 0.22;
    }
  }

  // ============ 4. 담당 영역 ============
  pres.addSection({ title: "담당 영역" });
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "담당 영역" });
    block(s, "04", "AREAS", "berry");
    const top = title(s, "담당 영역");
    bigQuote(s, 0.32, 1.8, ROSEc, 72);
    T(s, "여섯 영역을 한 사람이 맡아 운영합니다", { x: BLK_X, y: 2.8, w: BLK_W, h: 1.0, fontSize: 18, bold: true, color: CREAMc, lineSpacingMultiple: 1.25 });
    const cards = [
      ["재무·회계", "자금 운용, 세무 전 주기, 결산·외부감사", Fa.FaCalculator],
      ["인사·노무", "급여·4대보험, 채용, 해외 원격 인력 계약", Fa.FaUsers],
      ["IR·투자", "주총·이사회, 등기, 스톡옵션, 투자 계약", Fa.FaHandshake],
      ["정부사업", "공고 분석부터 정산까지 전 주기", Fa.FaLandmark],
      ["해외업무", "해외 자회사 증빙, 해외 행사·출장", Fa.FaGlobeAsia],
      ["총무", "법인 행정, 기업 인증, 사무실·자산, 대외행사", Fa.FaBuilding],
    ];
    const cw = 2.85, ch = 1.12, gx = 0.15, gy = 0.1;
    for (let i = 0; i < 6; i++) {
      const cx = BX + (i % 2) * (cw + gx), cy = top + Math.floor(i / 2) * (ch + gy);
      RR(s, { x: cx, y: cy, w: cw, h: ch, fill: { color: CREAM } });
      OVAL(s, { x: cx + 0.18, y: cy + 0.16, w: 0.42, h: 0.42, fill: { color: BERRY } });
      s.addImage({ data: await icon(cards[i][2], WHITE), x: cx + 0.28, y: cy + 0.26, w: 0.22, h: 0.22, objectName: "icon " + cards[i][0] });
      T(s, String(i + 1).padStart(2, "0"), { x: cx + cw - 0.65, y: cy + 0.14, w: 0.5, h: 0.28, fontSize: 11, bold: true, color: ROSEc, align: "right" });
      T(s, cards[i][0], { x: cx + 0.72, y: cy + 0.16, w: 1.45, h: 0.42, fontSize: 15, bold: true, color: INKc, valign: "middle" });
      T(s, cards[i][1], { x: cx + 0.18, y: cy + 0.66, w: cw - 0.34, h: 0.42, fontSize: 12, color: INKc, lineSpacingMultiple: 1.1 });
    }
  }

  // ============ 5. 담당 영역 상세 1 ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "담당 영역" });
    block(s, "05", "DETAIL 1", "cream");
    const top = title(s, "담당 영역 상세 1");
    const v = V.cream;
    section(s, "재무·회계", [
      "자금 운용과 송금·결제를 관리하고, 부가세·법인세·원천세·지방세 등 세무 전 주기를 대응합니다.",
      "결산과 외부감사 대응, 별도·연결 재무제표 관리를 맡고 외부 회계법인과의 커뮤니케이션을 담당합니다.",
    ], BLK_X, BLK_Y, BLK_W, 13, v);
    let cy = section(s, "인사·노무", [
      "급여·4대보험·연말정산·퇴직금 산정 등 인사 운영 실무를 처리합니다.",
      "채용 공고 기획부터 면접, 입사 조율까지 진행하고, 해외 원격 인력의 계약·온보딩·오프보딩 절차를 표준화했습니다.",
    ], BX, top, BW, 14, { head: ROSEc, body: INKc });
    section(s, "IR·투자", [
      "정기주주총회·이사회 운영, 정관 개정, 등기, 주주명부와 스톡옵션 관리를 맡고 있습니다.",
      "투자 계약 절차를 상법 조문 기준 일정표로 설계하고, 투자자 대상 분기 실적보고서(국문·영문)를 작성합니다.",
    ], BX, cy + 0.22, BW, 14, { head: ROSEc, body: INKc });
  }

  // ============ 6. 담당 영역 상세 2 ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "담당 영역" });
    block(s, "06", "DETAIL 2", "berry");
    const top = title(s, "담당 영역 상세 2");
    const v = V.berry;
    section(s, "해외업무", [
      "해외 자회사의 인보이스와 용역계약 증빙을 관리하고, 해외 사업 행사 참가와 해외 출장 서류 실무를 지원합니다.",
    ], BLK_X, BLK_Y, BLK_W, 14, v);
    header(s, "정부사업", BX, top, BW);
    let cy = flow(s, ["공고 분석", "사업계획서 작성", "신청", "협약", "정산"], BX, top + 0.3, BW, { pt: 11, h: 0.4 });
    cy += 0.14;
    {
      const ps = [
        "공고 분석 → 사업계획서 작성 → 신청 → 협약 → 정산까지 전 주기를 실무로 수행합니다.",
        "입찰·협약 서류를 반복해서 쓸 수 있는 패키지로 정리해 운영하고 있습니다.",
      ];
      const h = ps.reduce((a, p) => a + paraH(p, 14, BW), 0) + 8 / 72;
      T(s, ps.map((p, i) => ({ text: p, options: { breakLine: i < ps.length - 1 } })),
        { x: BX, y: cy, w: BW, h, fontSize: 14, color: INKc, lineSpacingMultiple: LS, paraSpaceAfter: 8 });
      cy += h;
    }
    section(s, "총무", [
      "본점 이전 같은 법인 행정 절차와 기업 인증 관리, 사무실·자산 관리를 담당합니다.",
      "대외행사 대관·운영을 맡고, 그 과정을 SOP로 남겼습니다.",
    ], BX, cy + 0.22, BW, 14, { head: ROSEc, body: INKc });
  }

  // ============ 7-11. 사례 ============
  pres.addSection({ title: "사례" });
  const cases = [
    {
      n: "07", variant: "cream", label: "CASE 01", title: "1인 경영지원을 “인계 가능한 시스템”으로 바꾸기",
      bg: "경영지원 인력이 한 명이라 업무가 개인 머릿속에만 있으면, 그 사람이 빠질 때 회사 업무가 멈추는 구조였습니다.",
      did: "업무 도구를 노션에서 옵시디언으로 완전히 옮겨 번호형 폴더 구조로 재구성했습니다. 반복 업무는 커스텀 AI 스킬로, 재발하는 업무는 SOP로 만들었고, 회사 기본 사실은 단일 원본 문서 하나로 모았습니다.",
      result: "옵시디언이 경영지원의 단일 운영체제가 되어, 신규 담당자도 볼트만 보고 업무 흐름을 파악할 수 있게 됐습니다.",
    },
    {
      n: "08", variant: "berry", label: "CASE 02", title: "단일 통화 자금 시트를 의사결정 도구로 재설계",
      bg: "국내 법인과 해외 자회사의 자금을 단일 통화 시트로 수동 관리하고 있어, 경영진의 자금 판단 질문에 바로 답하기 어려웠습니다.",
      did: "대시보드·통화별 입출금·정기 비용·미지급 항목으로 시트를 나누고, 다중 통화 환산과 “부족액 식별 → 송금액 계산 → 송금 후 잔액” 시뮬레이션을 자동화했습니다. 매월 정기 보고 SOP도 함께 만들었습니다.",
      flow: ["부족액 식별", "송금액 계산", "송금 후 잔액"],
      result: "경영진의 자금 의사결정 질문에 한 화면에서 바로 답할 수 있는 상태가 됐습니다.",
    },
    {
      n: "09", variant: "cream", label: "CASE 03", title: "법률 1차 검토를 법령 원문 기반으로 내재화",
      bg: "노무·세무·상법 판단이 실무 상식이나 기억에 기대는 경우가 많아, 근거가 약한 보고가 나갈 위험이 있었습니다.",
      did: "국가법령정보 공개 API와 오픈소스 법령 조회 도구를 업무용 AI에 연동했습니다. 판단마다 조문·시행령·판례를 확인하고 확인 여부 라벨을 붙이는 규칙을 고정했습니다.",
      result: "법이 걸린 판단이 조문과 시행일 근거를 갖춘 보고 형식으로 바뀌었고, 실무 상식만으로는 놓칠 수 있는 별도 조항을 미리 잡게 됐습니다.",
    },
    {
      n: "10", variant: "berry", label: "CASE 04", title: "외부로 나가기 전 “받은 산출물”을 역검증하는 습관",
      bg: "회계법인·기관·플랫폼에서 받은 자료나 내부 문서가 확인 없이 외부로 나가면 세금·등기·평가에 바로 영향을 줍니다.",
      did: "회계법인 산출물, 기관·플랫폼 자료, 계약서와 내부 문서를 원장·법령 원문과 대조한 뒤 제출하는 절차를 분기 내내 운영했습니다.",
      result: "기간 오류, 견적 오기, 정관과 계약의 충돌 같은 문제를 외부 제출 전에 찾아냈습니다.",
    },
    {
      n: "11", variant: "cream", label: "CASE 05", title: "예산 대비 실적 관리 워크북 구축",
      bg: "월별 지출 추적이 여러 엑셀 파일에 흩어져 있어 예산을 통제하기 어려웠습니다.",
      did: "실제 계정원장을 기준으로 계정과목을 정리해 카테고리로 묶었습니다. 월별 예산 수립 → 실적 입력 → 차이 분석 흐름을 자동화했고, 이후 중앙 가정 시트와 연동되도록 확장했습니다.",
      flow: ["예산 수립", "실적 입력", "차이 분석"],
      result: "예산 대비 실적을 한 체계에서 관리하는 관리회계 기반을 만들었습니다.",
    },
  ];
  for (const cs of cases) {
    const v = V[cs.variant];
    const s = pres.addSlide({ masterName: v.name, sectionTitle: "사례" });
    block(s, cs.n, cs.label, cs.variant);
    const top = title(s, cs.title);
    // 결과 — stands out inside the colored block with a big quote mark and label
    bigQuote(s, 0.32, 1.85, v.quote, 60);
    header(s, "결과", 1.2, 2.2, 1.6, v.head);
    const rh = paraH(cs.result, 14, BLK_W);
    T(s, cs.result, { x: BLK_X, y: 2.75, w: BLK_W, h: rh, fontSize: 14, bold: true, color: v.body, lineSpacingMultiple: LS });
    // 배경 / 한 일 in the body column; drop to 13pt only when 14pt would not fit
    const build = (pt) => {
      let cy = top;
      header(s, "배경", BX, cy, BW);
      const h1 = paraH(cs.bg, pt, BW);
      T(s, cs.bg, { x: BX, y: cy + 0.3, w: BW, h: h1, fontSize: pt, color: INKc, lineSpacingMultiple: LS });
      cy += 0.3 + h1 + 0.24;
      header(s, "한 일", BX, cy, BW);
      const h2 = paraH(cs.did, pt, BW);
      T(s, cs.did, { x: BX, y: cy + 0.3, w: BW, h: h2, fontSize: pt, color: INKc, lineSpacingMultiple: LS });
      cy += 0.3 + h2;
      if (cs.flow) {
        cy += 0.12;
        cy = flow(s, cs.flow, BX, cy, BW, { pt: 11, h: 0.38, fill: cs.variant === "cream" ? BERRY : CREAM, color: cs.variant === "cream" ? WHITEc : BERRYc });
      }
      return cy;
    };
    const est = (pt) => top + 0.3 + paraH(cs.bg, pt, BW) + 0.24 + 0.3 + paraH(cs.did, pt, BW) + (cs.flow ? 0.5 : 0);
    build(est(14) <= BB ? 14 : 13);
  }

  // ============ 12. 스킬·도구 ============
  pres.addSection({ title: "역량·경력" });
  {
    const s = pres.addSlide({ masterName: "CONTENT_BERRY", sectionTitle: "역량·경력" });
    block(s, "12", "SKILLS & TOOLS", "berry");
    const top = title(s, "스킬·도구");
    const v = V.berry;
    let by = section(s, "실무 지식", ["결산·세무 신고 대응, 주총·이사회·등기, 정부사업 전 주기"], BLK_X, BLK_Y, BLK_W, 14, v);
    section(s, "언어", ["비즈니스 영어(미국 현지 근무 경험)"], BLK_X, by + 0.25, BLK_W, 14, v);
    const rows = [
      ["엑셀", "함수 조합(VLOOKUP·IF), 피벗, 관리회계 워크북·자금 시뮬레이션 설계"],
      ["옵시디언", "번호형 볼트 구조, Templater·Tasks 플러그인"],
      ["AI 업무 자동화", "Claude 커스텀 스킬 설계·운영, 예약 작업 자동화, MCP 연동(법령 원문 조회, 옵시디언 연동)"],
      ["협업 도구", "Gmail, Google Drive, Slack"],
      ["ERP", "클라우드 ERP 1차 셋업(홈택스·은행·법인카드 연동)"],
      ["문서 표준화", "공문 서식, docx·pptx 폰트 표준, 인보이스·인사·급여 템플릿, SOP"],
    ];
    const lw = 1.45, dw = BW - lw - 0.1, pt = 13;
    let cy = top;
    rows.forEach(([lab, desc], i) => {
      const h = Math.max(paraH(desc, pt, dw), 0.3);
      T(s, lab, { x: BX, y: cy, w: lw, h: 0.3, fontSize: 14, bold: true, color: BERRYc });
      T(s, desc, { x: BX + lw + 0.1, y: cy + 0.02, w: dw, h, fontSize: pt, color: INKc, lineSpacingMultiple: LS });
      cy += h + 0.07;
      if (i < rows.length - 1) { HLINE(s, BX, cy, BW, CREAM, 1); cy += 0.1; }
    });
  }

  // ============ 13. 타임라인 ============
  {
    const s = pres.addSlide({ masterName: "CONTENT_CREAM", sectionTitle: "역량·경력" });
    block(s, "13", "TIMELINE", "cream");
    const top = title(s, "타임라인");
    T(s, "2013", { x: BLK_X - 0.02, y: 2.0, w: BLK_W, h: 0.95, fontSize: 60, bold: true, color: BERRYc, valign: "bottom" });
    T(s, "— 현재", { x: BLK_X, y: 3.0, w: BLK_W, h: 0.45, fontSize: 22, bold: true, color: ROSEc });
    T(s, "최신이 위", { x: BLK_X, y: 3.55, w: BLK_W, h: 0.3, fontSize: 11, color: INKc, charSpacing: 1 });
    const items = [
      ["2022 – 현재", "IT 스타트업", "경영지원팀 매니저"],
      ["2019 – 2022", "미국 소재 유통기업", "재고관리 인턴 → 구매(Purchasing) 담당"],
      ["2018 – 2019", "미국 비즈니스 취업연수 과정 수료", ""],
      ["2017", "중국 교환학생", ""],
      ["2013 – 2019", "학부", "국제통상학 전공 (중어중문학 복수전공)"],
    ];
    const pitch = 0.72, lineX = BX + 1.32, textX = lineX + 0.3;
    s.addShape(pres.shapes.LINE, { x: lineX, y: top + 0.12, w: 0, h: pitch * 4, line: { color: CREAM, width: 1.5 }, objectName: "timeline" });
    items.forEach(([when, org, role], i) => {
      const y = top + i * pitch;
      T(s, when, { x: BX, y: y + 0.02, w: 1.2, h: 0.26, fontSize: 13, bold: true, color: BERRYc });
      OVAL(s, { x: lineX - 0.09, y: y + 0.03, w: 0.18, h: 0.18, fill: { color: i === 0 ? BERRY : ROSE } });
      T(s, org, { x: textX, y: y - 0.01, w: BW - (textX - BX), h: 0.3, fontSize: 15, bold: true, color: INKc });
      if (role) T(s, role, { x: textX, y: y + 0.3, w: BW - (textX - BX), h: 0.28, fontSize: 13, color: INKc });
    });
  }

  // ============ 14. 연락 ============
  pres.addSection({ title: "마무리" });
  {
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: "마무리" });
    T(s, "“", { x: 0.55, y: -0.15, w: 2.5, h: 2.2, fontSize: 200, bold: true, color: ROSEc, valign: "top" });
    T(s, "연락", { x: 0.8, y: 1.85, w: 8, h: 1.0, fontSize: 56, bold: true, color: WHITEc, valign: "middle" });
    OVAL(s, { x: 0.8, y: 3.2, w: 0.5, h: 0.5, fill: { color: CREAM } });
    s.addImage({ data: await icon(Fa.FaEnvelope, BERRY), x: 0.92, y: 3.32, w: 0.26, h: 0.26, objectName: "icon email" });
    T(s, "namgyu1412@naver.com", { x: 1.45, y: 3.2, w: 5.5, h: 0.5, fontSize: 20, bold: true, color: WHITEc, valign: "middle" });
    T(s, "© 2026 Duke Yang", { x: 0.8, y: 4.95, w: 5, h: 0.3, fontSize: 11, color: CREAMc, charSpacing: 1 });
    T(s, "”", { x: 7.4, y: 3.3, w: 2.3, h: 2.4, fontSize: 200, bold: true, color: ROSEc, align: "right", valign: "bottom" });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
}

build().catch((e) => { console.error(e); process.exit(1); });
