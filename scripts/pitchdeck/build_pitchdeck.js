// BidBox 피치덱(.pptx) 생성기
// 근거: 사업계획서 v28(02_과제/..._v28.docx)과 AGENTS.md 「현재 사업계획서」의 수치만 쓴다. 새 숫자 없음.
// 구성: 6회차 강의(P·S·S·T, 목차·간지·장 제목 필수, 줄글 대신 숫자·OX·도식).
// 디자인: BidBox 웹 화면 색(primary #0073E6, sidebar #020B1A, slate) + Coinbase식 단일 강조색·어두운 띠 위 제품 화면
//         + IBM Carbon식 상태 3색(초록·주황·빨강).
// 실행: NODE_PATH=<pptxgenjs·react-icons·react·react-dom·sharp 설치 폴더>/node_modules \
//       node scripts/pitchdeck/build_pitchdeck.js <출력.pptx> <pptx 스킬 apply_theme.js 경로>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const ROOT = path.resolve(__dirname, "../..");
const CAP = (f) => path.join(ROOT, "02_과제/캡처_이미지", f);
const OUT = process.argv[2] || path.join(ROOT, "02_과제/피치덱/BidBox_피치덱_v1.pptx");
const APPLY = process.argv[3];

const F = "Malgun Gothic";
const THEME = {
  name: "BidBox",
  headFontFace: F,
  bodyFontFace: F,
  colors: {
    dk1: "0F172A", lt1: "FFFFFF", dk2: "020B1A", lt2: "F9FAFB",
    accent1: "0073E6", accent2: "059669", accent3: "B45309", accent4: "DC2626",
    accent5: "64748B", accent6: "E6E9EF", hlink: "0073E6", folHlink: "005BB3",
  },
};
// 색 토큰(hex). 옅은 바탕색은 테마에 없으니 hex로 둔다.
const K = {
  ink: "0F172A", ink2: "334155", mute: "64748B", line: "E6E9EF", bg: "F9FAFB", white: "FFFFFF",
  blue: "0073E6", blueDeep: "005BB3", blueTint: "E6F1FF", blue2: "9CC9F7", blue3: "4D9DEC",
  navy: "020B1A", navy2: "0F1B2E", navy3: "1E2A3D", onDarkMute: "A8B3C4", skyOnDark: "6CB4FF",
  green: "059669", greenTint: "ECFDF5", amber: "B45309", amberTint: "FFFBEB", red: "DC2626", redTint: "FEF2F2",
  slateTint: "F1F5F9",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: F, bodyFontFace: F };
pres.title = "BidBox 피치덱";
pres.company = "Narani";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6, CW = W - 2 * M;

// ---------- 레이아웃 ----------
const titlePh = (dark) => ({
  placeholder: {
    options: { name: "title", type: "title", x: M, y: 0.72, w: CW, h: 0.8, fontFace: F, fontSize: 30, bold: true,
      color: dark ? C.background1 : C.text1, valign: "top", align: "left", margin: 0 },
    text: "",
  },
});
const eyebrowPh = (dark) => ({
  placeholder: {
    options: { name: "eyebrow", type: "body", x: M, y: 0.38, w: CW, h: 0.34, fontFace: F, fontSize: 13, bold: true,
      color: dark ? C.accent6 : C.accent1, valign: "top", margin: 0 },
    text: "",
  },
});
const footer = (dark) => ({
  text: { text: "BidBox · Narani", options: { x: M, y: 7.02, w: 4, h: 0.3, fontFace: F, fontSize: 10,
    color: dark ? K.onDarkMute : K.mute, margin: 0 } },
});
pres.defineSlideMaster({ title: "LIGHT", background: { color: K.white },
  objects: [eyebrowPh(false), titlePh(false), footer(false)],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: K.mute, align: "right" } });
pres.defineSlideMaster({ title: "DARK", background: { color: K.navy },
  objects: [eyebrowPh(true), titlePh(true), footer(true)],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: K.onDarkMute, align: "right" } });
pres.defineSlideMaster({ title: "DIVIDER", background: { color: K.navy },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: M + 0.1, y: 3.35, w: 9, h: 1.0, fontFace: F,
    fontSize: 44, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } }, footer(true)],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: K.onDarkMute, align: "right" } });
pres.defineSlideMaster({ title: "COVER", background: { color: K.navy }, objects: [] });

// ---------- 도우미 ----------
let objN = 0;
const nm = (p) => `${p}-${++objN}`;
function T(s, text, o) {
  s.addText(text, Object.assign({ fontFace: F, margin: 0, isTextBox: true, valign: "top", objectName: nm("text") }, o));
}
function box(s, o) {
  const shape = o.radius === 0 ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE;
  const opt = { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill || K.white },
    line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
    objectName: nm(o.name || "card") };
  if (shape === pres.shapes.ROUNDED_RECTANGLE) opt.rectRadius = o.radius == null ? 0.14 : o.radius;
  if (o.shadow) opt.shadow = { type: "outer", color: "000000", opacity: 0.12, blur: 8, offset: 2, angle: 90 };
  s.addShape(shape, opt);
}
function chip(s, text, x, y, color, tint, o = {}) {
  const w = o.w || Math.max(0.9, text.length * 0.135 + 0.35);
  s.addText(text, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.16, x, y, w, h: o.h || 0.32,
    fill: { color: tint }, line: { type: "none" }, fontFace: F, fontSize: o.size || 11, bold: true, color,
    align: "center", valign: "middle", margin: 0, objectName: nm("chip") });
  return w;
}
const iconCache = {};
async function icon(name, color, px = 256) {
  const key = name + color;
  if (!iconCache[key]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: px }));
    const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
    iconCache[key] = "image/png;base64," + buf.toString("base64");
  }
  return iconCache[key];
}
// 원 안 아이콘(모티프)
async function iconCircle(s, name, x, y, d, fg, bg) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: nm("icon-bg") });
  const p = d * 0.27;
  s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icon") });
}
function arrow(s, x, y, w, color = K.blue2) {
  s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h: 0.3, fill: { color }, line: { type: "none" }, objectName: nm("arrow") });
}
// 어두운 바탕 위 제품 화면 카드(Coinbase product-ui-card-dark)
function shot(s, file, x, y, w, ratio, label) {
  const h = w / ratio, p = 0.1;
  box(s, { x: x - p, y: y - p, w: w + 2 * p, h: h + 2 * p, fill: K.navy3, radius: 0.16, shadow: true, name: "shot-frame" });
  s.addImage({ path: CAP(file), x, y, w, h, objectName: nm("shot") });
  if (label) chip(s, label, x + 0.15, y + h - 0.47, K.white, K.navy, { size: 12 });
  return h;
}
function lightSlide(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function darkSlide(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function divider(section, num, title, sub, weight) {
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: section });
  T(s, num, { x: M + 0.1, y: 1.45, w: 4, h: 1.8, fontSize: 110, bold: true, color: K.blue });
  s.addText(title, { placeholder: "title" });
  T(s, sub, { x: M + 0.1, y: 4.45, w: 9, h: 0.5, fontSize: 20, color: K.onDarkMute });
  chip(s, `예비창업패키지 평가 비중 ${weight}`, M + 0.1, 5.3, K.skyOnDark, K.navy3, { size: 13, h: 0.42, w: 3.6 });
  return s;
}
function note(s, text, y = 6.62, dark = false) {
  T(s, text, { x: M, y, w: CW, h: 0.32, fontSize: 10, color: dark ? K.onDarkMute : K.mute });
}

(async () => {
  // ===== 1. 표지 =====
  pres.addSection({ title: "표지" });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  T(s, "2026 예비창업패키지 · 피치덱", { x: M, y: 1.0, w: 6, h: 0.4, fontSize: 14, bold: true, color: K.skyOnDark });
  T(s, "공고마다 다른 점수 규칙,\n찾아서 금액 범위까지", { x: M, y: 1.55, w: 6.2, h: 1.9, fontSize: 38, bold: true,
    color: K.white, lineSpacingMultiple: 1.1 });
  T(s, "공공기관 용역 입찰 도우미 웹 서비스", { x: M, y: 3.7, w: 6, h: 0.4, fontSize: 18, color: K.onDarkMute });
  T(s, "BidBox (가칭)", { x: M, y: 4.12, w: 6, h: 0.6, fontSize: 28, bold: true, color: K.blue });
  let cx = M;
  for (const [ic, t] of [["FaBroom", "청소"], ["FaShieldHalved", "경비"], ["FaPeopleGroup", "파견"], ["FaHandshake", "직업소개"]]) {
    await iconCircle(s, ic, cx, 5.05, 0.5, K.skyOnDark, K.navy3);
    T(s, t, { x: cx + 0.58, y: 5.15, w: 1.0, h: 0.3, fontSize: 14, color: K.white });
    cx += t.length > 2 ? 1.75 : 1.35;
  }
  T(s, "Narani", { x: M, y: 6.55, w: 3, h: 0.4, fontSize: 16, bold: true, color: K.white });
  shot(s, "dashboard.jpg", 7.2, 1.05, 5.5, 1530 / 661, "대시보드");
  shot(s, "detail_top.jpg", 6.55, 3.75, 5.65, 1568 / 661, "공고 상세");
  s.addNotes("BidBox는 공공기관 청소·경비·파견·직업소개 용역 입찰에서, 공고에 맞는 점수 계산 규칙을 찾아 써 낼 금액의 범위를 알려 주는 웹 서비스입니다. 오른쪽은 실제 동작하는 시제품 화면입니다.");

  // ===== 2. 목차 =====
  pres.addSection({ title: "목차" });
  s = lightSlide("목차", "목차", "발표 순서");
  const ag = [["P", "01", "문제 인식", "누가, 언제,\n무엇 때문에 불편한가"], ["S", "02", "실현 가능성", "제품, 차별점,\n협약 8개월 계획"],
    ["S", "03", "성장전략", "시장, 수익 구조,\n로드맵, 자금"], ["T", "04", "팀 구성", "대표자 역량,\n채용, 협력 기관"]];
  ag.forEach(([l, n, t, d], i) => {
    const x = M + i * (2.9 + 0.177);
    box(s, { x, y: 1.75, w: 2.9, h: 3.05, fill: K.bg, line: K.line, name: "agenda-card" });
    T(s, l, { x: x + 0.3, y: 1.95, w: 1.2, h: 0.9, fontSize: 48, bold: true, color: K.blue });
    T(s, n, { x: x + 1.9, y: 2.1, w: 0.7, h: 0.4, fontSize: 16, bold: true, color: K.mute, align: "right" });
    T(s, t, { x: x + 0.3, y: 2.95, w: 2.4, h: 0.5, fontSize: 22, bold: true, color: K.ink });
    T(s, d, { x: x + 0.3, y: 3.55, w: 2.4, h: 0.9, fontSize: 14, color: K.mute });
  });
  T(s, "예비창업패키지 발표 평가 비중", { x: M, y: 5.2, w: 6, h: 0.35, fontSize: 14, bold: true, color: K.ink });
  const wts = [[10, "문제 인식 10%", K.blue2, K.ink], [40, "실현 가능성 40%", K.blue, K.white], [35, "성장전략 35%", K.blueDeep, K.white], [15, "팀 15%", K.navy3, K.white]];
  let bx = M;
  for (const [p, t, f, c] of wts) {
    const w = CW * p / 100;
    s.addText(t, { shape: pres.shapes.RECTANGLE, x: bx, y: 5.65, w, h: 0.6, fill: { color: f }, line: { color: K.white, width: 2 },
      fontFace: F, fontSize: p < 12 ? 11 : 14, bold: true, color: c, align: "center", valign: "middle", margin: 0, objectName: nm("weight") });
    bx += w;
  }
  note(s, "강의(6회차) 기준 배점. 장 수도 실현 가능성과 성장전략에 가장 많이 배분", 6.42);
  s.addNotes("예비창업패키지 발표 평가 비중에 맞춰, 실현 가능성과 성장전략에 가장 많은 장을 배분했습니다.");

  // ===== P 간지 =====
  pres.addSection({ title: "P 문제 인식" });
  s = divider("P 문제 인식", "01", "문제 인식", "누가, 언제, 무엇 때문에 불편한가", "10%");
  s.addNotes("첫 번째, 문제 인식입니다.");

  // ===== 4. 고객 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "고객: 공고를 자주 여는 중소 용역업체 입찰 담당자");
  box(s, { x: M, y: 1.75, w: 5.4, h: 4.75, fill: K.bg, line: K.line });
  await iconCircle(s, "FaUserTie", M + 0.35, 2.05, 1.0, K.blue, K.blueTint);
  T(s, "입찰 담당자", { x: M + 1.6, y: 2.12, w: 3.6, h: 0.45, fontSize: 22, bold: true, color: K.ink });
  T(s, "공고를 열고, 얼마를 써 낼지 정하는 사람", { x: M + 1.6, y: 2.62, w: 3.6, h: 0.6, fontSize: 14, color: K.mute });
  T(s, "1차 시장: 네 업종 허가(면허)를 가진 회사", { x: M + 0.35, y: 3.5, w: 4.8, h: 0.35, fontSize: 14, bold: true, color: K.ink2 });
  const ind = [["FaBroom", "청소(위생관리)"], ["FaShieldHalved", "경비"], ["FaPeopleGroup", "파견"], ["FaHandshake", "직업소개"]];
  for (let i = 0; i < 4; i++) {
    const x = M + 0.35 + (i % 2) * 2.5, y = 4.05 + Math.floor(i / 2) * 1.0;
    await iconCircle(s, ind[i][0], x, y, 0.7, K.white, K.blue);
    T(s, ind[i][1], { x: x + 0.85, y: y + 0.18, w: 1.6, h: 0.4, fontSize: 15, bold: true, color: K.ink });
  }
  const who = [["FaUser", "1차 사용자", "네 업종 회사에서 공고를 자주 여는 입찰 담당자", "직무로 특정 · 인터뷰 없음"],
    ["FaStamp", "구매 결정자", "회사 대표 또는 입찰 비용을 결재하는 사람", "지불 의사 미확인"],
    ["FaBriefcase", "2차 사용자", "입찰 대행·교육 기관 실무자 (경쟁 상대이자 첫 판매 통로)", "확정 거래처 없음"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.75 + i * 1.65;
    box(s, { x: 6.35, y, w: 6.38, h: 1.45, fill: K.white, line: K.line });
    await iconCircle(s, who[i][0], 6.6, y + 0.32, 0.8, K.blue, K.blueTint);
    T(s, who[i][1], { x: 7.6, y: y + 0.2, w: 2.5, h: 0.4, fontSize: 16, bold: true, color: K.ink });
    chip(s, who[i][3], 12.48 - (who[i][3].length * 0.135 + 0.35), y + 0.22, K.amber, K.amberTint);
    T(s, who[i][2], { x: 7.6, y: y + 0.68, w: 4.95, h: 0.6, fontSize: 14, color: K.ink2 });
  }
  note(s, "이름이 비슷해도 허가(면허)가 다른 업종은 목록에서 제외");
  s.addNotes("고객은 청소·경비·파견·직업소개 네 업종 중소 용역업체의 입찰 담당자입니다. 직무로 특정했고, 인터뷰와 지불 의사는 시험 사용에서 확인합니다.");

  // ===== 5. 문제 흐름 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "공고 한 건마다 규칙표를 직접 찾아 맞춰야 함");
  const steps = [["FaFileLines", "공고 본문 열기", "본문에는 점수 계산 규칙이 없음"],
    ["FaPaperclip", "첨부 파일 뒤지기", "규칙표(별표)와 기준값은 첨부 파일에만 있음"],
    ["FaCodeBranch", "맞는 판 고르기", "기관과 시행일마다 규칙이 다름"],
    ["FaTriangleExclamation", "탈락 위험", "규칙을 잘못 고르면 금액을 잘못 써 냄"]];
  for (let i = 0; i < 4; i++) {
    const x = M + i * 3.13, red = i === 3;
    box(s, { x, y: 1.8, w: 2.75, h: 2.75, fill: red ? K.redTint : K.bg, line: red ? "FCA5A5" : K.line });
    T(s, `0${i + 1}`, { x: x + 0.3, y: 2.0, w: 1, h: 0.4, fontSize: 16, bold: true, color: red ? K.red : K.blue });
    await iconCircle(s, steps[i][0], x + 1.7, 1.98, 0.75, red ? K.white : K.blue, red ? K.red : K.blueTint);
    T(s, steps[i][1], { x: x + 0.3, y: 2.95, w: 2.3, h: 0.45, fontSize: 18, bold: true, color: red ? K.red : K.ink });
    T(s, steps[i][2], { x: x + 0.3, y: 3.5, w: 2.2, h: 0.9, fontSize: 14, color: K.ink2 });
    if (i < 3) arrow(s, x + 2.82, 3.03, 0.25);
  }
  box(s, { x: M, y: 4.85, w: CW, h: 1.6, fill: K.white, line: K.line });
  T(s, "실제 예", { x: M + 0.35, y: 5.05, w: 2.2, h: 0.35, fontSize: 13, bold: true, color: K.blue });
  T(s, "인천광역시 계약정보", { x: M + 0.35, y: 5.42, w: 2.6, h: 0.45, fontSize: 17, bold: true, color: K.ink });
  for (const [x, t] of [[3.55, "2017-04-11본"], [6.15, "2017-12-13 개정본"]]) {
    box(s, { x, y: 5.1, w: 2.3, h: 1.1, fill: K.bg, line: K.line, name: "doc" });
    s.addImage({ data: await icon("FaFileLines", K.mute), x: x + 0.2, y: 5.35, w: 0.45, h: 0.6, objectName: nm("doc-icon") });
    T(s, t, { x: x + 0.8, y: 5.45, w: 1.45, h: 0.5, fontSize: 13, bold: true, color: K.ink });
  }
  s.addImage({ data: await icon("FaCircleQuestion", K.amber), x: 5.88, y: 5.48, w: 0.24, h: 0.24, objectName: nm("q") });
  T(s, "같은 기준의 옛 판과 고친 판이\n나란히 게시 → 어느 판을 쓸지\n담당자가 직접 판단", { x: 8.85, y: 5.08, w: 3.7, h: 1.2, fontSize: 14, color: K.ink2, valign: "middle" });
  s.addNotes("점수를 계산하려면 공고마다 규칙표와 기준값이 필요한데, 이 내용은 첨부 파일에만 있습니다. 인천광역시처럼 같은 기준의 옛 판과 고친 판이 나란히 올라오기도 합니다.");

  // ===== 6. 문제 크기 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "10건 중 3건꼴은 기준이 없어 자동 계산 불가");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "최근 12개월 용역 적격심사", labels: ["규칙 자동 적용 21,394건", "기준 없음 9,911건"], values: [21394, 9911] }],
    { x: M, y: 1.65, w: 5.4, h: 4.75, holeSize: 66, chartColors: [K.blue, K.red], showLegend: true, legendPos: "b",
      legendFontFace: F, legendFontSize: 13, legendColor: K.ink2, showValue: false, showPercent: false, showLabel: false,
      dataBorder: { pt: 2, color: "FFFFFF" }, objectName: nm("donut") });
  T(s, "31.7%", { x: M + 1.45, y: 3.0, w: 2.5, h: 0.75, fontSize: 40, bold: true, color: K.red, align: "center" });
  T(s, "기준 없음", { x: M + 1.45, y: 3.75, w: 2.5, h: 0.35, fontSize: 14, color: K.mute, align: "center" });
  const st = [["9,911건", "기준이 없어 자동 계산 불가", K.red], ["31,305건", "최근 12개월 용역 적격심사 공고", K.ink], ["55.1조원", "2024년 공공조달 용역 계약 (조달청)", K.blue]];
  st.forEach(([v, l, c], i) => {
    const y = 1.8 + i * 1.55;
    box(s, { x: 6.5, y, w: 6.23, h: 1.35, fill: K.bg, line: K.line });
    T(s, v, { x: 6.85, y: y + 0.2, w: 3.0, h: 0.9, fontSize: 34, bold: true, color: c, valign: "middle" });
    T(s, l, { x: 9.75, y: y + 0.2, w: 2.85, h: 0.9, fontSize: 14, color: K.ink2, valign: "middle" });
  });
  note(s, "규칙 판정 2026-10-19 · 55.1조원은 조달청 보도자료(2025-06-27)의 계약 규모이며 소프트웨어 매출 아님");
  s.addNotes("최근 12개월 용역 적격심사 31,305건 중 9,911건, 31.7%는 기준이 없어 자동 계산이 되지 않습니다. 2024년 공공조달 용역 계약은 55.1조원입니다.");

  // ===== 7. 기존 방법 한계 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "지금 쓰는 방법: 규칙표를 골라 주는 곳이 없음");
  const alt = [["FaUserTie", "사람 자문", "인포21C · 비드큐", [["✕", "써 낼 금액을 대신 정해 줌"], ["✕", "따내면 성공보수 (비드큐 1:1 분석 1.5%)"]]],
    ["FaCalculator", "계산기", "비드프로 · 디마툴즈", [["✕", "담당자가 넣은 숫자로 점수만 계산"], ["✕", "규칙표를 잘못 골라도 잡아 주지 못함"]]],
    ["FaLandmark", "나라장터", "공공조달 시스템", [["△", "공고 검색과 금액 제출 중심"], ["?", "공고별 규칙표 안내 여부는 확인 전"]]]];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 4.14;
    box(s, { x, y: 1.75, w: 3.85, h: 3.9, fill: K.white, line: K.line, shadow: true });
    await iconCircle(s, alt[i][0], x + 0.35, 2.05, 0.85, K.mute, K.slateTint);
    T(s, alt[i][1], { x: x + 1.4, y: 2.1, w: 2.3, h: 0.45, fontSize: 20, bold: true, color: K.ink });
    T(s, alt[i][2], { x: x + 1.4, y: 2.55, w: 2.3, h: 0.35, fontSize: 13, color: K.mute });
    alt[i][3].forEach(([m, t], j) => {
      const y = 3.35 + j * 1.0, col = m === "✕" ? K.red : K.amber;
      s.addText(m, { shape: pres.shapes.OVAL, x: x + 0.35, y, w: 0.42, h: 0.42, fill: { color: m === "✕" ? K.redTint : K.amberTint },
        line: { type: "none" }, fontFace: F, fontSize: 14, bold: true, color: col, align: "center", valign: "middle", margin: 0, objectName: nm("mark") });
      T(s, t, { x: x + 0.95, y: y - 0.02, w: 2.7, h: 0.8, fontSize: 14, color: K.ink2 });
    });
  }
  s.addText([{ text: "빈자리 두 곳   ", options: { bold: true, color: K.blue } },
    { text: "① 공고에 맞는 규칙표 골라 주기     ② 기준이 없을 때 계산 멈추기", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.14, x: M, y: 5.9, w: CW, h: 0.75, fill: { color: K.blueTint },
      line: { type: "none" }, fontFace: F, fontSize: 18, align: "center", valign: "middle", margin: 0, objectName: nm("gap-banner") });
  s.addNotes("사람 자문은 금액을 대신 정해 주고 성공보수를 받습니다. 계산기는 넣은 숫자로만 계산합니다. 규칙표를 골라 주는 일과 기준이 없을 때 멈추는 일은 비어 있습니다.");

  // ===== S1 간지 =====
  pres.addSection({ title: "S 실현 가능성" });
  s = divider("S 실현 가능성", "02", "실현 가능성", "제품과 서비스가 있어야 하는 이유", "40%");
  s.addNotes("두 번째, 실현 가능성입니다.");

  // ===== 9. 불편 → 해결 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "고객의 불편을 한 화면에서 해결");
  T(s, "고객의 불편", { x: M + 0.25, y: 1.62, w: 4, h: 0.35, fontSize: 14, bold: true, color: K.mute });
  T(s, "BidBox의 해결", { x: 7.25, y: 1.62, w: 4, h: 0.35, fontSize: 14, bold: true, color: K.blue });
  const ps = [["기관·시행일마다 규칙이 달라 어느 판을 쓸지 모름", "FaCodeBranch", "공고 날짜에 맞는 판을 골라 자동 적용"],
    ["탈락하지 않는 금액 범위를 직접 계산", "FaCalculator", "최저가·최상가를 거꾸로 계산, 회사 신용등급·신인도 반영"],
    ["지난 입찰에서 얼마에 따냈는지 따로 찾기", "FaChartLine", "AI 예측가와 예상 낙찰률을 금액 범위 옆에 표시"],
    ["매일 맞는 공고를 직접 고르기", "FaMagnifyingGlass", "네 업종이 목록 기본값, 공고 505,271건 색인"]];
  for (let i = 0; i < 4; i++) {
    const y = 2.05 + i * 1.13;
    box(s, { x: M, y, w: 5.75, h: 0.95, fill: K.slateTint, radius: 0.12 });
    T(s, ps[i][0], { x: M + 0.25, y, w: 5.3, h: 0.95, fontSize: 15, color: K.ink2, valign: "middle" });
    arrow(s, 6.47, y + 0.33, 0.45, K.blue);
    box(s, { x: 7.0, y, w: 5.73, h: 0.95, fill: K.blueTint, radius: 0.12 });
    await iconCircle(s, ps[i][1], 7.18, y + 0.17, 0.6, K.white, K.blue);
    T(s, ps[i][2], { x: 7.95, y, w: 4.65, h: 0.95, fontSize: 15, bold: true, color: K.ink, valign: "middle" });
  }
  note(s, "낙찰 보증 없음. 세 가지 금액과 예측 범위는 참고 정보");
  s.addNotes("고객의 네 가지 불편을 BidBox가 한 화면에서 해결합니다. 맞는 판 적용, 금액 범위 계산, AI 예측가 표시, 네 업종 기본 목록입니다.");

  // ===== 10. 제품 화면 (어두운 띠) =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "이미 동작하는 시제품 1식");
  shot(s, "detail_top.jpg", M + 0.1, 1.72, 7.8, 1568 / 661, "공고 상세: 맞는 규칙과 세 가지 금액");
  shot(s, "dashboard.jpg", 8.75, 1.72, 3.88, 1530 / 661, "대시보드");
  shot(s, "bids_list.jpg", 8.75, 3.63, 3.88, 1530 / 661, "네 업종 공고 탐색");
  const flow = ["공고 모으기", "네 업종 검색", "맞는 규칙 찾기", "회사 점수 반영", "세 가지 금액", "근거·상담", "기록으로 보강"];
  flow.forEach((t, i) => {
    const x = M + i * 1.773;
    s.addText(t, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.2, x, y: 5.6, w: 1.5, h: 0.62,
      fill: { color: i === 2 ? K.blue : K.navy3 }, line: { type: "none" }, fontFace: F, fontSize: 13, bold: true,
      color: K.white, align: "center", valign: "middle", margin: 0, objectName: nm("flow") });
    if (i < 6) arrow(s, x + 1.53, 5.76, 0.2, K.skyOnDark);
  });
  note(s, "실제 프로토타입 화면 (2026-10-22 캡처) · 규칙을 찾지 못하면 계산 멈춤", 6.45, true);
  s.addNotes("모두 실제 동작하는 시제품 화면입니다. 공고를 모으고, 네 업종을 검색하고, 맞는 규칙을 찾아 회사 점수를 반영한 세 가지 금액과 근거를 보여 줍니다.");

  // ===== 11. 세 가지 금액 + 모르면 멈춤 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "세 가지 금액을 한 화면에, 모르면 계산 멈춤");
  box(s, { x: M, y: 1.75, w: 6.4, h: 4.7, fill: K.bg, line: K.line });
  T(s, "세 가지 금액", { x: M + 0.35, y: 1.98, w: 4, h: 0.4, fontSize: 18, bold: true, color: K.ink });
  const bx0 = M + 0.45, bw = 5.5;
  box(s, { x: bx0, y: 3.55, w: bw, h: 0.26, fill: K.line, radius: 0.13, name: "range-track" });
  box(s, { x: bx0 + bw * 0.12, y: 3.55, w: bw * 0.76, h: 0.26, fill: K.blue2, radius: 0.13, name: "range-band" });
  const mk = [[0.12, "최저가", "써도 되는\n가장 낮은 금액", K.blueDeep], [0.5, "AI 예측가", "AI가 예상한\n금액", K.blue], [0.88, "최상가", "써도 되는\n가장 높은 금액", K.blueDeep]];
  for (const [p, n, d, c] of mk) {
    const x = bx0 + bw * p;
    s.addShape(pres.shapes.OVAL, { x: x - 0.17, y: 3.51, w: 0.34, h: 0.34, fill: { color: c }, line: { color: K.white, width: 2 }, objectName: nm("marker") });
    T(s, n, { x: x - 0.9, y: 2.95, w: 1.8, h: 0.4, fontSize: 16, bold: true, color: c, align: "center" });
    T(s, d, { x: x - 0.9, y: 4.0, w: 1.8, h: 0.65, fontSize: 12, color: K.ink2, align: "center" });
  }
  await iconCircle(s, "FaBuilding", M + 0.35, 4.95, 0.55, K.blue, K.blueTint);
  T(s, "회사 신용등급·신인도를 넣으면 최저가·최상가 계산\n없으면 AI 예측가만 보여 주고 입력 안내", { x: M + 1.05, y: 4.88, w: 5.2, h: 0.75, fontSize: 13, color: K.ink2, valign: "middle" });
  chip(s, "낙찰 보증 없음 · 참고 정보", M + 0.35, 5.85, K.red, K.redTint, { size: 12, w: 2.6 });
  const rr = 880 / 450, rw = 5.2;
  box(s, { x: 7.42, y: 1.67, w: rw + 0.16, h: rw / rr + 0.16, fill: K.white, line: "FCA5A5", lw: 1.5, radius: 0.1, name: "rule-frame" });
  s.addImage({ path: CAP("detail_rule.jpg"), x: 7.5, y: 1.75, w: rw, h: rw / rr, objectName: nm("rule-shot") });
  T(s, "규칙을 찾지 못한 공고: 틀린 금액 대신 계산 멈춤", { x: 7.5, y: 4.55, w: 5.2, h: 0.35, fontSize: 14, bold: true, color: K.red });
  const stops = ["금액 없는 공고에 0원 추천 없음", "기준 금액 없는 공고(비예가)는 가격점수 계산 안 함", "협상 계약은 평가 비율·낙찰률 참고 분포"];
  for (let i = 0; i < 3; i++) {
    s.addImage({ data: await icon("FaCircleStop", K.red), x: 7.5, y: 5.05 + i * 0.47, w: 0.26, h: 0.26, objectName: nm("stop") });
    T(s, stops[i], { x: 7.9, y: 5.01 + i * 0.47, w: 4.85, h: 0.34, fontSize: 13, color: K.ink2 });
  }
  s.addNotes("최저가, AI 예측가, 최상가를 한 화면에 보여 줍니다. 규칙을 찾지 못하면 틀린 금액 대신 계산을 멈춥니다. 금액이 없는 공고에 0원을 추천하지 않습니다.");

  // ===== 12. 경쟁 비교 OX =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "기존 방식과 비교: 짧게, OX로");
  const sym = { "○": K.green, "△": K.amber, "✕": K.red, "—": "94A3B8" };
  const cell = (v, bid) => {
    const m = v[0], rest = v.slice(1).trim();
    const runs = [{ text: m, options: { color: sym[m] || K.ink, bold: true, fontSize: 16 } }];
    if (rest) runs.push({ text: "  " + rest, options: { color: bid ? K.ink : K.ink2, bold: !!bid, fontSize: 12 } });
    return { text: runs, options: { fill: { color: bid ? K.blueTint : K.white } } };
  };
  const hdr = ["항목", "나라장터", "계산기", "사람 자문", "BidBox"].map((h, i) => ({ text: h,
    options: { bold: true, fontSize: 14, color: i === 4 ? K.white : K.ink, fill: { color: i === 4 ? K.blue : K.slateTint }, align: i ? "center" : "left" } }));
  const rowsD = [["규칙표 고르기", "—", "✕ 넣은 숫자로 점수만", "—", "○ 자동 판정 (규칙 206종)"],
    ["기준 없을 때", "—", "✕ 넣은 값 그대로 계산", "—", "○ 계산 멈춤, 직접 입력"],
    ["세 가지 금액", "△ 검색·제출 중심", "△ 점수만 계산", "○ 금액 제안", "○ 최저·예측·최상 한 화면"],
    ["근거 표시", "—", "✕ 결과 점수 중심", "—", "○ 규칙 번호·원문 인용"],
    ["옛 판·지금 판", "—", "△ 확인 필요", "—", "○ 공고 날짜 기준 적용"],
    ["금액 비공개 공고", "—", "—", "—", "○ 0원 추천 없음"],
    ["요금", "—", "△ 구독 (근거 제외)", "✕ 성공보수 1.5%", "○ 이용권 1,000원 (가설)"]];
  const rows = [hdr].concat(rowsD.map((r) => r.map((v, i) => i === 0
    ? { text: v, options: { bold: true, fontSize: 14, color: K.ink, fill: { color: K.bg } } }
    : Object.assign(cell(v, i === 4), {}))));
  rows.slice(1).forEach((r) => r.slice(1).forEach((c) => (c.options.align = "center")));
  s.addTable(rows, { x: M, y: 1.7, w: CW, colW: [2.35, 2.15, 2.45, 2.15, 3.03], rowH: 0.56, fontFace: F, valign: "middle",
    border: { type: "solid", pt: 0.75, color: K.line }, margin: [0, 0.08, 0, 0.12], objectName: nm("compare") });
  note(s, "○ 지원  △ 부분 지원  ✕ 미지원  — 미확인(근거 미확보, 숫자를 만들지 않음) · 성공보수 1.5%는 비드큐 1:1 분석 안내(bidq.co.kr)", 6.35);
  s.addNotes("강의 팁대로 줄글 대신 OX로 비교했습니다. 규칙표 고르기, 기준이 없을 때 멈추기, 근거 표시는 BidBox만 지원합니다. 확인하지 못한 칸은 미확인으로 두었습니다.");

  // ===== 13. 개발 현황 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "이미 만든 것, 숫자로 확인");
  const tiles = [["206종", "규칙 ID", "기관·시행일별 계산 규칙"], ["21,394건", "자동 적용", "31,305건 중 (2026-10-19)"],
    ["505,271건", "공고 색인", "최근 1년 (2026-09-09)"], ["1.1885%p", "낙찰률 평균 오차", "3,000건 (2026-09-15)"],
    ["90.25%", "예측 범위 포함률", "3,589건 (2026-08-30)"], ["144/144", "상담 숫자 정답", "96개 요청 평가"]];
  tiles.forEach(([v, l, d], i) => {
    const x = M + (i % 3) * 2.83, y = 1.75 + Math.floor(i / 3) * 2.42;
    box(s, { x, y, w: 2.65, h: 2.22, fill: K.bg, line: K.line });
    T(s, v, { x: x + 0.25, y: y + 0.3, w: 2.3, h: 0.7, fontSize: 28, bold: true, color: K.blue });
    T(s, l, { x: x + 0.25, y: y + 1.1, w: 2.3, h: 0.4, fontSize: 15, bold: true, color: K.ink });
    T(s, d, { x: x + 0.25, y: y + 1.5, w: 2.3, h: 0.4, fontSize: 11, color: K.mute });
  });
  box(s, { x: 9.2, y: 1.75, w: 3.53, h: 4.64, fill: K.white, line: K.line, shadow: true });
  T(s, "개발 진행 상태 (10개 항목)", { x: 9.45, y: 1.97, w: 3.1, h: 0.4, fontSize: 14, bold: true, color: K.ink });
  const stat = [[7, "구현 완료", K.green, K.greenTint], [1, "부분 구현", K.amber, K.amberTint], [2, "미구축 · 협약 중 구축", K.red, K.redTint]];
  let sq = 0;
  for (const [n, , c] of stat) for (let k = 0; k < n; k++, sq++)
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.45 + sq * 0.305, y: 2.5, w: 0.26, h: 0.42, rectRadius: 0.05, fill: { color: c }, line: { type: "none" }, objectName: nm("sq") });
  stat.forEach(([n, l, c, t], i) => {
    s.addText(String(n), { shape: pres.shapes.OVAL, x: 9.45, y: 3.2 + i * 0.6, w: 0.42, h: 0.42, fill: { color: t }, line: { type: "none" },
      fontFace: F, fontSize: 14, bold: true, color: c, align: "center", valign: "middle", margin: 0, objectName: nm("count") });
    T(s, l, { x: 10.0, y: 3.24 + i * 0.6, w: 2.6, h: 0.36, fontSize: 13, color: K.ink2 });
  });
  T(s, "협약 중 만들 것\n· 운영 주소, 이용권 결제, 약관, 보안 점검\n· 시험 사용 기록, 지자체 원문", { x: 9.45, y: 5.1, w: 3.15, h: 1.15, fontSize: 12, color: K.ink2 });
  s.addNotes("규칙 206종, 31,305건 중 21,394건 자동 적용, 낙찰률 평균 오차 1.1885%p, 상담 숫자 정답 144/144입니다. 협약 기간에는 운영 공개와 결제, 시험 사용을 만듭니다.");

  // ===== 14. 협약 8개월 일정 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "협약 8개월: 운영 공개에서 유료 전환까지");
  const half = (CW - 0.2) / 2;
  s.addText([{ text: "1단계 · 1~4개월   ", options: { bold: true, color: K.blue } }, { text: "2,000만 원", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 1.72, w: half, h: 0.6, fill: { color: K.blueTint }, line: { type: "none" },
      fontFace: F, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: nm("phase") });
  s.addText([{ text: "2단계 · 5~8개월   ", options: { bold: true, color: K.ink2 } }, { text: "2,000만 원 (중간평가 통과 시)", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + half + 0.2, y: 1.72, w: half, h: 0.6, fill: { color: K.slateTint }, line: { type: "none" },
      fontFace: F, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: nm("phase") });
  s.addShape(pres.shapes.LINE, { x: M + 0.4, y: 3.25, w: CW - 0.8, h: 0, line: { color: K.line, width: 3 }, objectName: nm("tl") });
  const ms = [["운영 공개", "서버·도메인·보안 접속, Windows 실행 확인"], ["회사 점수 재현", "등급 넣은 계정과 비교, 약관 초안"],
    ["이용권 결제", "1토큰 1,000원(가설), 보안 점검"], ["시험 사용 1차", "10사 사용 기록, 중간평가 자료"],
    ["지자체 규칙 보강", "막힌 원문 등록, 영업 담당 채용"], ["시험 사용 확대", "20사(가설), 비수도권 10사 이상"],
    ["단가 확정", "상표 1건, 특허 출원 1건"], ["유료 전환", "유료 회사 수(기준 8사), 결과 보고"]];
  const step = CW / 8;
  ms.forEach(([t, d], i) => {
    const cxm = M + step * i + step / 2, hot = i === 3 || i === 7;
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: cxm - 0.3, y: 2.95, w: 0.6, h: 0.6, fill: { color: hot ? K.blue : K.white },
      line: { color: K.blue, width: 2 }, fontFace: F, fontSize: 16, bold: true, color: hot ? K.white : K.blue, align: "center", valign: "middle", margin: 0, objectName: nm("node") });
    T(s, `${i + 1}개월차`, { x: cxm - 0.72, y: 2.48, w: 1.44, h: 0.3, fontSize: 11, color: K.mute, align: "center" });
    T(s, t, { x: cxm - 0.72, y: 3.75, w: 1.44, h: 0.6, fontSize: 14, bold: true, color: K.ink, align: "center" });
    T(s, d, { x: cxm - 0.72, y: 4.35, w: 1.44, h: 1.0, fontSize: 11, color: K.ink2, align: "center" });
  });
  const goals = [["시험 사용", "10사", "4개월차"], ["시험 사용", "20사", "6개월차 (가설)"], ["유료 전환", "8사", "8개월차 (가설)"]];
  goals.forEach(([a, b, c], i) => {
    const x = M + i * 4.1;
    box(s, { x, y: 5.55, w: 3.9, h: 0.9, fill: K.bg, line: K.line });
    T(s, a, { x: x + 0.3, y: 5.55, w: 1.4, h: 0.9, fontSize: 14, color: K.ink2, valign: "middle" });
    T(s, b, { x: x + 1.55, y: 5.55, w: 1.0, h: 0.9, fontSize: 26, bold: true, color: K.blue, valign: "middle" });
    T(s, c, { x: x + 2.5, y: 5.55, w: 1.35, h: 0.9, fontSize: 12, color: K.mute, valign: "middle" });
  });
  s.addNotes("1단계 4개월은 운영 공개, 결제, 시험 사용 1차입니다. 2단계는 중간평가를 통과해야 집행하며, 시험 사용을 20사로 넓히고 유료 전환을 확인합니다.");

  // ===== S2 간지 =====
  pres.addSection({ title: "S 성장전략" });
  s = divider("S 성장전략", "03", "성장전략", "어떤 고객이 어떤 방식으로 얼마를 낼 것인가", "35%");
  s.addNotes("세 번째, 성장전략입니다.");

  // ===== 16. 시장 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "목표 시장: 55.1조원에서 네 업종 7,494억원부터");
  const ccx = M + 2.4, by = 6.35;
  const circ = [[4.6, K.blueTint, "TAM"], [3.15, K.blue2, "SAM"], [1.75, K.blue, "SOM"]];
  for (const [d, f, n] of circ) s.addShape(pres.shapes.OVAL, { x: ccx - d / 2, y: by - d, w: d, h: d, fill: { color: f }, line: { color: K.white, width: 2 }, objectName: nm(n) });
  T(s, "TAM", { x: ccx - 1, y: by - 4.35, w: 2, h: 0.35, fontSize: 14, bold: true, color: K.blue, align: "center" });
  T(s, "55.1조원", { x: ccx - 1.2, y: by - 4.0, w: 2.4, h: 0.5, fontSize: 22, bold: true, color: K.ink, align: "center" });
  T(s, "SAM", { x: ccx - 1, y: by - 2.95, w: 2, h: 0.3, fontSize: 13, bold: true, color: K.blueDeep, align: "center" });
  T(s, "10.2조원", { x: ccx - 1.2, y: by - 2.65, w: 2.4, h: 0.45, fontSize: 20, bold: true, color: K.ink, align: "center" });
  T(s, "SOM", { x: ccx - 0.8, y: by - 1.35, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: K.white, align: "center" });
  T(s, "7,494억원+", { x: ccx - 0.85, y: by - 1.05, w: 1.7, h: 0.4, fontSize: 15, bold: true, color: K.white, align: "center" });
  const mk2 = [["TAM 전체 시장", "국내 공공조달 용역 계약", "2024 · 조달청 (2025-06-27)"],
    ["SAM 닿을 수 있는 시장", "나라장터 용역 적격심사 낙찰 20,824건", "2025 · 나라장터 낙찰결과 집계"],
    ["SOM 처음 들어갈 시장", "네 업종 면허 적격심사 낙찰 1,869건", "2025-09-15~2026-09-14 · 하한"]];
  mk2.forEach(([a, b, c], i) => {
    const y = 1.85 + i * 1.5;
    T(s, a, { x: 5.55, y, w: 3.2, h: 0.35, fontSize: 14, bold: true, color: i === 2 ? K.blue : K.ink });
    T(s, b, { x: 5.55, y: y + 0.38, w: 3.2, h: 0.6, fontSize: 13, color: K.ink2 });
    T(s, c, { x: 5.55, y: y + 0.95, w: 3.2, h: 0.3, fontSize: 10, color: K.mute });
  });
  // 막대는 도형으로 그린다(구글 슬라이드 변환 때 차트의 항목 이름·순서·숫자 형식이 깨짐)
  T(s, "SOM 업종별 낙찰 (억원)", { x: 8.95, y: 1.85, w: 3.78, h: 0.35, fontSize: 14, bold: true, color: K.ink });
  const som = [["위생관리", 5316, "1,319건"], ["경비업", 3471, "651건"], ["파견", 642, "220건"]];
  som.forEach(([n, v, c], i) => {
    const y = 2.5 + i * 1.25, w = 2.6 * v / 5316;
    T(s, n, { x: 8.95, y, w: 2.0, h: 0.3, fontSize: 13, bold: true, color: K.ink2 });
    box(s, { x: 8.95, y: y + 0.38, w, h: 0.42, fill: K.blue, radius: 0.06, name: "som-bar" });
    T(s, v.toLocaleString("en-US"), { x: 8.95 + w + 0.1, y: y + 0.38, w: 0.8, h: 0.42, fontSize: 14, bold: true, color: K.ink, valign: "middle" });
    T(s, c, { x: 10.6, y, w: 2.1, h: 0.3, fontSize: 11, color: K.mute, align: "right" });
  });
  note(s, "모두 고객 회사가 따는 계약 금액이며 BidBox 매출 아님 · SOM은 면허 정보가 있는 공고(적격심사 금액의 56%)만 센 하한, 업종 사이 중복 있음 · 유료직업소개는 면허 공고 20건(39억원), 적격심사 낙찰 0건", 6.48);
  s.addNotes("TAM은 2024년 공공조달 용역 계약 55.1조원, SAM은 나라장터 용역 적격심사 낙찰 10.2조원, SOM은 네 업종 면허 적격심사 낙찰 7,494억원 이상입니다. 모두 고객이 따는 계약 금액이고 BidBox 매출이 아닙니다.");

  // ===== 17. 비즈니스 모델 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "비즈니스 모델: 검색은 무료, 판단은 이용권으로");
  chip(s, "B2B", 11.93, 0.4, K.blue, K.blueTint, { w: 0.8, size: 12 });
  const hd = [["누가 내나", M], ["무엇을 받나", 4.75], ["어떻게, 언제", 9.35]];
  for (const [t, x] of hd) T(s, t, { x: x + 0.2, y: 1.6, w: 3, h: 0.3, fontSize: 12, bold: true, color: K.mute });
  const bm = [["무료 입구", "모든 방문자", "네 업종 검색, 대시보드, 기준값 직접 입력 계산", "무료 · 지금(구현됨)", "FaGift", 0],
    ["① 주수익", "네 업종 중소기업\n(대표·입찰 비용 결재자)", "규칙 자동 찾기, 세 가지 금액, 근거 상담", "이용권 1토큰 1,000원\n협약 3개월차 결제 시작", "FaTicket", 2],
    ["② 부수익", "입찰 대행·교육 기관", "고객 회사 대신 쓰는 묶음 이용권", "묶음 단가는 사용 확인 뒤\n3년차", "FaLayerGroup", 1],
    ["③ 부수익", "담당자가 여럿이거나\n자기 시스템에 붙이는 회사", "계정 추가, 프로그램 연결(API)", "협약 중 가격 미공개\n4년차(계획)", "FaPlug", 1]];
  for (let i = 0; i < 4; i++) {
    const [tag, who2, what, how, ic, lv] = bm[i], y = 1.95 + i * 1.03, main = lv === 2;
    box(s, { x: M, y, w: 3.85, h: 0.88, fill: K.bg, line: K.line, radius: 0.12 });
    chip(s, tag, M + 0.15, y + 0.28, main ? K.white : (lv ? K.blue : K.mute), main ? K.blue : (lv ? K.blueTint : K.slateTint), { w: 0.95, size: 11 });
    T(s, who2, { x: M + 1.22, y, w: 2.55, h: 0.88, fontSize: 13, bold: true, color: K.ink, valign: "middle" });
    arrow(s, 4.5, y + 0.29, 0.3, K.blue2);
    box(s, { x: 4.95, y, w: 4.25, h: 0.88, fill: main ? K.blueTint : K.white, line: main ? K.blue : K.line, lw: main ? 1.5 : 1, radius: 0.12 });
    await iconCircle(s, ic, 5.1, y + 0.16, 0.56, main ? K.white : K.blue, main ? K.blue : K.blueTint);
    T(s, what, { x: 5.8, y, w: 3.3, h: 0.88, fontSize: 13, color: K.ink, valign: "middle", bold: main });
    arrow(s, 9.27, y + 0.29, 0.3, K.blue2);
    box(s, { x: 9.65, y, w: 3.08, h: 0.88, fill: main ? K.blue : K.white, line: main ? K.blue : K.line, radius: 0.12 });
    T(s, how, { x: 9.8, y, w: 2.85, h: 0.88, fontSize: 12, bold: main, color: main ? K.white : K.ink2, valign: "middle" });
  }
  s.addText([{ text: "유료 8사 × 연 100토큰 × 1,000원 = 연 80만 원 (가설)", options: { bold: true, color: K.ink } },
    { text: "   손익이 아니라 지불 의사를 확인하는 매출", options: { color: K.ink2 } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 6.12, w: 9.6, h: 0.62, fill: { color: K.bg }, line: { color: K.line, width: 1 },
      fontFace: F, fontSize: 15, align: "center", valign: "middle", margin: 0, objectName: nm("formula") });
  chip(s, "낙찰 성공보수 없음", 10.4, 6.12, K.red, K.redTint, { w: 2.33, h: 0.62, size: 14 });
  s.addNotes("검색과 대시보드는 무료이고, 규칙 찾기와 세 가지 금액, 상담은 이용권으로 씁니다. 1토큰 1,000원, 유료 8사가 연 100토큰을 사면 연 80만 원입니다. 이 값은 지불 의사를 보는 가설입니다. 낙찰 성공보수는 받지 않습니다.");

  // ===== 18. 고객 검증·진입 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "고객 검증: 시험 사용으로 지불 의사 확인");
  const fun = [["무료 입구", "지금 · 네 업종 검색·대시보드", 6.6, K.blueTint, K.ink], ["시험 사용 10사", "4개월차 · 사용 기록, 이용권 구매 의사", 5.5, K.blue2, K.ink],
    ["시험 사용 20사", "6개월차 · 비수도권 10사 이상 (가설)", 4.4, K.blue, K.white], ["유료 8사", "8개월차 (가설)", 3.3, K.blueDeep, K.white]];
  fun.forEach(([a, b, w, f, c], i) => {
    const y = 1.8 + i * 1.12, x = M + (6.6 - w) / 2;
    s.addText([{ text: a, options: { bold: true, fontSize: 18, breakLine: true } }, { text: b, options: { fontSize: 12 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.14, x, y, w, h: 0.95, fill: { color: f }, line: { type: "none" },
        fontFace: F, color: c, align: "center", valign: "middle", margin: 0, objectName: nm("funnel") });
  });
  box(s, { x: 7.65, y: 1.8, w: 5.08, h: 2.3, fill: K.greenTint, radius: 0.14 });
  await iconCircle(s, "FaDatabase", 7.9, 2.05, 0.6, K.white, K.green);
  T(s, "지금 가진 근거: 공공데이터", { x: 8.65, y: 2.1, w: 3.9, h: 0.45, fontSize: 16, bold: true, color: K.green });
  T(s, "· 나라장터 공개 데이터 공고 505,271건 색인\n· 용역 적격심사 31,305건 규칙 판정\n· 시험 사용에서 입찰 횟수·검토 시간·비용 측정", { x: 7.95, y: 2.75, w: 4.65, h: 1.25, fontSize: 13, color: K.ink2 });
  box(s, { x: 7.65, y: 4.3, w: 5.08, h: 2.08, fill: K.amberTint, radius: 0.14 });
  await iconCircle(s, "FaFlag", 7.9, 4.55, 0.6, K.white, K.amber);
  T(s, "아직 없는 것, 협약 중 확보", { x: 8.65, y: 4.6, w: 3.9, h: 0.45, fontSize: 16, bold: true, color: K.amber });
  T(s, "· 확정 거래처, 고객 인터뷰, 지불 의사\n· 4개월까지 입찰 대행·교육 기관 2곳과\n  시험 사용 자리 확보 목표", { x: 7.95, y: 5.2, w: 4.65, h: 1.1, fontSize: 13, color: K.ink2 });
  note(s, "무료 기능은 유료로 넘어가는 입구, 낙찰 대행은 목표 아님 · 근거가 없는 시점·횟수는 가설·확정 전으로 표기");
  s.addNotes("아직 매출과 고객이 없는 예비창업 단계라, 시험 사용으로 고객을 검증합니다. 공공데이터로 만든 근거는 이미 있고, 거래처와 지불 의사는 협약 중에 확보합니다.");

  // ===== 19. 5개년 로드맵 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "5개년 로드맵: 용역 → 물품 → 공사 → 자체조달");
  const rd = [["1년차", "네 업종 상용화", "0.75조원", K.blueTint, K.ink],
    ["2년차", "용역 적격심사 전 업종\n물품 착수", "10.2조원", "CCE3FA", K.ink],
    ["3년차", "물품 공개 · 공사 착수\n묶음 이용권", "23.4조원", K.blue2, K.ink],
    ["4년차 (계획)", "공사 공개\n여러 사용자·API", "54.4조원", K.blue, K.white],
    ["5년차 (계획)", "나라장터 밖\n자체조달 기관 검토", "+80.0조원", K.blueDeep, K.white]];
  const sw = (CW - 4 * 0.12) / 5;
  rd.forEach(([y1, t, v, f, c], i) => {
    const h = 1.15 + i * 0.6, x = M + i * (sw + 0.12), y = 6.3 - h;
    box(s, { x, y, w: sw, h, fill: f, radius: 0.1, name: "step" });
    T(s, v, { x: x + 0.2, y: y + h - 0.7, w: sw - 0.4, h: 0.5, fontSize: 20, bold: true, color: c });
    T(s, y1, { x: x + 0.2, y: y + 0.18, w: sw - 0.4, h: 0.35, fontSize: 14, bold: true, color: c });
    T(s, t, { x, y: y - 0.85, w: sw, h: 0.75, fontSize: 13, color: K.ink2, align: "center", valign: "bottom" });
  });
  note(s, "누적 공략 범위 = 고객이 따는 계약 금액(BidBox 매출 아님) · 물품은 2025 입찰 낙찰 전체 13.2조원, 공사 적격심사 31.0조원, 자체조달 80.0조원(2024) · 4~5년차는 계획(안)", 6.48);
  s.addNotes("1년차 네 업종으로 시작해 2년차 용역 전 업종, 3년차 물품, 4년차 공사, 5년차 자체조달 기관으로 넓힙니다. 공략 범위는 고객이 따는 계약 금액입니다.");

  // ===== 20. 자금 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "자금: 정부지원 4,000만 원, 매출은 지불 의사 확인용");
  // 누적 막대도 도형으로(구글 변환 때 차트 숫자 형식이 빠짐). 단위 만원, 합계 각 2,000
  T(s, "정부지원사업비 집행계획 (만원)", { x: M, y: 1.8, w: 7.4, h: 0.35, fontSize: 14, bold: true, color: K.ink });
  const BC = { "외주용역비": K.blue, "지급수수료": K.blueDeep, "광고선전비": K.blue3, "인건비(영업 담당)": "1E3A5F", "무형자산(상표·특허)": K.mute };
  const stages = [["1단계", [["외주용역비", 800, "800"], ["지급수수료", 720, "720"], ["광고선전비", 480, "480"]]],
    ["2단계", [["인건비(영업 담당)", 1000, "1,000"], ["지급수수료", 579.836, "580"], ["무형자산(상표·특허)", 300, "300"], ["광고선전비", 120.164, "120"]]]];
  const bw0 = 6.4, bx1 = M + 0.95;
  stages.forEach(([n, segs], i) => {
    const y = 2.45 + i * 1.45;
    T(s, n, { x: M, y, w: 0.9, h: 0.95, fontSize: 15, bold: true, color: K.ink2, valign: "middle" });
    let xx = bx1;
    for (const [k, v, lab] of segs) {
      const w = bw0 * v / 2000;
      s.addText(lab, { shape: pres.shapes.RECTANGLE, x: xx, y, w, h: 0.95, fill: { color: BC[k] }, line: { color: K.white, width: 1.5 },
        fontFace: F, fontSize: w < 0.6 ? 11 : 13, bold: true, color: K.white, align: "center", valign: "middle", margin: 0, objectName: nm("budget-seg") });
      xx += w;
    }
  });
  Object.entries(BC).forEach(([k, c], i) => {
    const x = M + (i % 3) * 2.5, y = 5.5 + Math.floor(i / 3) * 0.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.07, w: 0.2, h: 0.2, fill: { color: c }, line: { type: "none" }, objectName: nm("legend-key") });
    T(s, k, { x: x + 0.3, y, w: 2.15, h: 0.34, fontSize: 12, color: K.ink2 });
  });
  const fund = [["정부지원 1단계", "2,000만 원", "1~4개월 · 운영 공개·결제·시험 사용", K.blue], ["정부지원 2단계", "2,000만 원", "5~8개월 · 중간평가 통과 후 집행", K.blue],
    ["매출", "연 80만 원", "유료 8사 × 연 100토큰 (가설)", K.amber], ["엔젤·시드", "미확인", "협약 중 투자 유치 일정 없음", K.mute]];
  fund.forEach(([a, v, d, c], i) => {
    const y = 1.75 + i * 1.12;
    box(s, { x: 8.4, y, w: 4.33, h: 0.98, fill: K.bg, line: K.line, radius: 0.12 });
    T(s, a, { x: 8.62, y: y + 0.12, w: 2.0, h: 0.35, fontSize: 13, bold: true, color: K.ink2 });
    T(s, v, { x: 10.45, y: y + 0.08, w: 2.15, h: 0.45, fontSize: 18, bold: true, color: c, align: "right" });
    T(s, d, { x: 8.62, y: y + 0.55, w: 4.0, h: 0.32, fontSize: 11, color: K.mute });
  });
  note(s, "2단계 지급수수료 5,798,360원·광고선전비 1,201,640원 · 대표자 인건비 없음 · 현물 출자 계획 없음 · 월 운영비 3,468,720원(3차 회의 견적) · 2단계 실제 지원액은 중간평가 결과로 결정", 6.48);
  s.addNotes("정부지원 1단계와 2단계 각 2,000만 원입니다. 1단계는 외주용역, 지급수수료, 광고선전비이고, 2단계는 영업 담당 인건비, 지급수수료, 상표·특허입니다. 대표자 인건비와 현물 출자는 넣지 않았습니다.");

  // ===== 21. 사회적 가치 =====
  s = lightSlide("S 성장전략", "S · 성장전략 · ESG", "사회적 가치: 공정한 정보, 일하는 사람의 조건");
  const esg = [["S", "FaScaleBalanced", "조달 정보 차이 줄이기", "규칙표·가격 기준을 공고 화면에 공개, 검색·대시보드·직접 계산 무료", "2024 중소기업 수주 142.1조원(63.1%)", "1년차 시험 사용 20사 중 비수도권 10사 이상"],
    ["S", "FaHelmetSafety", "용역 노동자 근로조건", "근로조건 이행계획 점수(별표 2 시설분야 10점)를 계산에 반영", "고용노동부 도급 노동자 보호지침(2026-09-08)", "점수 반영 구현 · 노임 경고 3년차"],
    ["S", "FaUsers", "약자기업·고용 우수기업", "여성·장애인기업, 청년·여성·장애인 고용 등 가점 계산과 안내", "적격심사 신인도 항목, 규칙에 구현", "2년차 가점 안내 화면"],
    ["G", "FaUserShield", "정보 보호·책임 있는 AI", "필요한 항목만 수집, 상담은 서비스 안 모델, 글자 읽기(OCR) 기본 꺼짐", "약관 첫머리에 낙찰 비보증 명시", "1년차부터"]];
  const ew = (CW - 3 * 0.18) / 4;
  for (let i = 0; i < 4; i++) {
    const [L, ic, t, d, g, when] = esg[i], x = M + i * (ew + 0.18);
    box(s, { x, y: 1.75, w: ew, h: 4.75, fill: K.white, line: K.line, shadow: true });
    s.addText(L, { shape: pres.shapes.OVAL, x: x + 0.25, y: 2.0, w: 0.55, h: 0.55, fill: { color: L === "G" ? K.navy3 : K.blue }, line: { type: "none" },
      fontFace: F, fontSize: 18, bold: true, color: K.white, align: "center", valign: "middle", margin: 0, objectName: nm("esg-letter") });
    s.addImage({ data: await icon(ic, K.blue2), x: x + ew - 0.85, y: 2.0, w: 0.55, h: 0.55, objectName: nm("esg-icon") });
    T(s, t, { x: x + 0.25, y: 2.75, w: ew - 0.45, h: 0.7, fontSize: 16, bold: true, color: K.ink });
    T(s, d, { x: x + 0.25, y: 3.5, w: ew - 0.45, h: 1.3, fontSize: 13, color: K.ink2 });
    T(s, g, { x: x + 0.25, y: 4.85, w: ew - 0.45, h: 0.6, fontSize: 10, color: K.mute });
    s.addText(when, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x + 0.2, y: 5.6, w: ew - 0.4, h: 0.65,
      fill: { color: K.blueTint }, line: { type: "none" }, fontFace: F, fontSize: 11, bold: true, color: K.blueDeep, align: "center", valign: "middle", margin: 0.05, objectName: nm("when") });
  }
  note(s, "지표의 숫자 목표는 시험 사용 측정 뒤 결정 · 약자기업 우대 단가와 노임 경고 기준은 확인 전까지 수치 미설정");
  s.addNotes("네 업종은 인건비 비중이 큰 단순노무 용역이라, 근로조건 항목이 금액 판단과 바로 이어집니다. 정보 차이 줄이기, 근로조건, 약자기업 가점, 정보 보호의 네 가지를 단계별로 도입합니다.");

  // ===== T 간지 =====
  pres.addSection({ title: "T 팀 구성" });
  s = divider("T 팀 구성", "04", "팀 구성", "우리 팀이 가진 역량과 채울 자리", "15%");
  s.addNotes("마지막, 팀 구성입니다.");

  // ===== 23. 팀 =====
  s = lightSlide("T 팀 구성", "T · 팀 구성", "대표자가 서비스 전체를 직접 구현");
  box(s, { x: M, y: 1.75, w: 6.3, h: 4.75, fill: K.bg, line: K.line });
  await iconCircle(s, "FaUserTie", M + 0.35, 2.05, 1.0, K.white, K.blue);
  T(s, "대표자", { x: M + 1.6, y: 2.1, w: 3, h: 0.45, fontSize: 22, bold: true, color: K.ink });
  T(s, "기획 · 개발 · 데이터 검증", { x: M + 1.6, y: 2.6, w: 4.3, h: 0.4, fontSize: 14, color: K.mute });
  const rep = [["FaBriefcase", "경력", "전시컨벤션 사업 매니저 6년 9개월\n국제회의·기업행사 기획·진행, 사업입찰 진행 포함"],
    ["FaRobot", "기술", "규칙 206종 정리, 낙찰률 예측 모델 학습·운영,\n네 업종 화면과 근거 상담을 한 시제품으로 연결"],
    ["FaListCheck", "원칙", "기준이 없으면 계산을 멈춤, 원인→시도→해결→검증 기록"]];
  for (let i = 0; i < 3; i++) {
    const y = 3.3 + i * 0.95;
    await iconCircle(s, rep[i][0], M + 0.35, y, 0.55, K.blue, K.blueTint);
    T(s, rep[i][1], { x: M + 1.05, y: y + 0.1, w: 0.8, h: 0.35, fontSize: 14, bold: true, color: K.blue });
    T(s, rep[i][2], { x: M + 1.8, y: y - 0.02, w: 4.35, h: 0.8, fontSize: 13, color: K.ink2 });
  }
  box(s, { x: 7.2, y: 1.75, w: 5.53, h: 2.05, fill: K.white, line: K.blue, lw: 1.5, dash: "dash" });
  await iconCircle(s, "FaUserPlus", 7.45, 2.0, 0.7, K.blue, K.blueTint);
  T(s, "영업 담당 1명", { x: 8.35, y: 2.02, w: 2.6, h: 0.4, fontSize: 18, bold: true, color: K.ink });
  chip(s, "채용 예정 · 협약 5개월차", 10.5, 2.06, K.amber, K.amberTint, { w: 2.05 });
  T(s, "시험 사용 회사 모집·영업, 입찰 실무 검수\n공공조달 입찰·영업 경력 2~3년 (채용 조건)", { x: 8.35, y: 2.6, w: 4.25, h: 0.9, fontSize: 13, color: K.ink2 });
  box(s, { x: 7.2, y: 4.0, w: 5.53, h: 2.5, fill: K.white, line: K.line });
  T(s, "협력 기관 3곳 (모두 확정 전)", { x: 7.45, y: 4.18, w: 5, h: 0.4, fontSize: 15, bold: true, color: K.ink });
  const prt = [["FaStamp", "변리", "상표 1건·특허 출원 1건", "7개월차"], ["FaLock", "보안 점검", "취약점 점검 1식", "3개월차"], ["FaHandshake", "입찰 대행·교육", "고객 모집 채널", "4개월까지 2곳"]];
  for (let i = 0; i < 3; i++) {
    const y = 4.72 + i * 0.58;
    await iconCircle(s, prt[i][0], 7.45, y, 0.44, K.blue, K.blueTint);
    T(s, prt[i][1], { x: 8.05, y: y + 0.05, w: 1.6, h: 0.34, fontSize: 13, bold: true, color: K.ink });
    T(s, prt[i][2], { x: 9.6, y: y + 0.05, w: 1.9, h: 0.34, fontSize: 12, color: K.ink2 });
    T(s, prt[i][3], { x: 11.4, y: y + 0.05, w: 1.2, h: 0.34, fontSize: 11, color: K.mute, align: "right" });
  }
  note(s, "매출·유료 고객·낙찰 실적은 아직 없음 · 빈자리인 고객 영업은 영업 담당 채용과 시험 사용 검수로 보완");
  s.addNotes("대표자가 규칙 206종, 예측 모델, 네 업종 화면을 직접 만들었습니다. 빈자리인 고객 영업은 협약 5개월차에 공공조달 경력의 영업 담당을 채용해 채웁니다.");

  // ===== 24. 마무리 =====
  pres.addSection({ title: "마무리" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "마무리" });
  T(s, "BidBox", { x: M, y: 1.1, w: 6, h: 0.5, fontSize: 20, bold: true, color: K.blue });
  T(s, "모르면 멈추는\n입찰 도우미", { x: M, y: 1.7, w: 8, h: 1.9, fontSize: 48, bold: true, color: K.white, lineSpacingMultiple: 1.05 });
  const fin = [["206종", "공고에 맞는 계산 규칙"], ["31,305건", "최근 12개월 규칙 판정"], ["0원", "금액 없는 공고 추천"]];
  fin.forEach(([v, l], i) => {
    const x = M + i * 4.1;
    box(s, { x, y: 4.05, w: 3.85, h: 1.35, fill: K.navy2, line: K.navy3, radius: 0.16 });
    T(s, v, { x: x + 0.3, y: 4.2, w: 3.3, h: 0.65, fontSize: 30, bold: true, color: K.skyOnDark });
    T(s, l, { x: x + 0.3, y: 4.88, w: 3.3, h: 0.35, fontSize: 14, color: K.onDarkMute });
  });
  T(s, "감사합니다 · 질문 받겠습니다", { x: M, y: 5.9, w: 8, h: 0.5, fontSize: 22, bold: true, color: K.white });
  T(s, "Narani · 낙찰 보증 없음, 세 가지 금액은 참고 정보", { x: M, y: 6.5, w: 8, h: 0.35, fontSize: 12, color: K.onDarkMute });
  s.addNotes("BidBox는 공고에 맞는 규칙을 찾고, 모르면 멈추는 입찰 도우미입니다. 감사합니다. 질문 받겠습니다.");

  await pres.writeFile({ fileName: OUT });
  if (APPLY) { const { applyTheme } = require(APPLY); await applyTheme(OUT, THEME); }
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
