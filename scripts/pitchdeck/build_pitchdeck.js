// BidBox 피치덱(.pptx) 생성기
// 수치·사실: AGENTS.md 「현재 사업계획서」와 「피치덱」. 새 숫자 없음.
// 여백·글자 크기·카드 간격·명암: .agents/skills/bizplan-writing/references/pitch-design.md (v11부터).
// 구성: 6회차 강의(P·S·S·T, 장 제목 필수, 줄글 대신 숫자·OX·도식). 발표용이라 목차·간지 없음(강의 6쪽 「피칭 X」).
// 디자인: BidBox 웹 화면 색(primary #0073E6, sidebar #020B1A, slate) + Coinbase식 단일 강조색·어두운 띠 위 제품 화면
//         + IBM Carbon식 상태 3색(초록·주황·빨강). 레이아웃 숫자는 아직 v10 측정값(pitch-design.md 2절).
// 실행: NODE_PATH=<pptxgenjs·react-icons·react·react-dom·sharp 설치 폴더>/node_modules \
//       node scripts/pitchdeck/build_pitchdeck.js <출력.pptx> <pptx 스킬 apply_theme.js 경로>
// 검토용 구글 슬라이드에 올리는 절차: .agents/skills/bizplan-writing/references/gslides.md
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
  ink: "0F172A", ink2: "334155", mute: "64748B", line: "E1E7EF", bg: "FFFFFF", white: "FFFFFF",
  page: "F2F5F9", mark: "E2E8F0", // 본문 장 바탕(연한 회청색)과 장 기호 워터마크
  blue: "0073E6", blueDeep: "005BB3", blueTint: "E6F1FF", blue2: "9CC9F7", blue3: "4D9DEC",
  navy: "020B1A", navy2: "0F1B2E", navy3: "1E2A3D", onDarkMute: "A8B3C4", skyOnDark: "6CB4FF",
  green: "059669", greenTint: "ECFDF5", amber: "B45309", amberTint: "FFFBEB", red: "DC2626", redTint: "FEF2F2",
  slateTint: "E8EDF3",
  // 어두운 장 위 상태 색(대비를 위해 한 단계 밝게)
  greenOnDark: "34D399", amberOnDark: "FBBF24", redOnDark: "F87171",
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
// BidBox 로고: 제품 저장소 logo.png에서 make_logo.py로 바탕을 투명하게 뺀 것(가로:세로 870:308)
const LOGO = (dark) => path.join(__dirname, "assets", dark ? "bidbox_logo_dark.png" : "bidbox_logo_light.png");
const LOGO_R = 870 / 308;
const footer = (dark) => [
  { image: { x: M, y: 7.03, w: 0.26 * LOGO_R, h: 0.26, path: LOGO(dark) } },
  { text: { text: "· Narani", options: { x: M + 0.26 * LOGO_R + 0.08, y: 7.02, w: 2, h: 0.3, fontFace: F, fontSize: 10,
    color: dark ? K.onDarkMute : K.mute, margin: 0 } } },
];
pres.defineSlideMaster({ title: "LIGHT", background: { color: K.page },
  objects: [eyebrowPh(false), titlePh(false), ...footer(false)],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: K.mute, align: "right" } });
pres.defineSlideMaster({ title: "DARK", background: { color: K.navy },
  objects: [eyebrowPh(true), titlePh(true), ...footer(true)],
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
function arrow2(s, x, y, w, color = K.blue2) {
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x, y, w, h: 0.3, fill: { color }, line: { type: "none" }, objectName: nm("arrow2") });
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
// 장 기호(P·S·T)를 오른쪽 아래에 아주 옅게 깐다(6회차 강의 슬라이드의 큰 「P」). 카드보다 먼저 넣어 뒤에 깔리게 한다.
function watermark(s, eyebrow) {
  const L = eyebrow.trim()[0];
  // T(팀) 장은 카드 사이로 「T」 가로획만 조각처럼 비쳐 결함처럼 보여서 깔지 않는다
  if (!"PS".includes(L)) return;
  T(s, L, { x: 9.0, y: 2.2, w: 4.2, h: 5.2, fontSize: 340, bold: true, color: K.mark, align: "right", valign: "bottom", objectName: nm("watermark") });
}
// 제목의 [[핵심어]]만 강조색으로(6회차 강의의 2색 제목)
function richTitle(title, accent) {
  if (!title.includes("[[")) return title;
  return title.split(/\[\[|\]\]/).map((t, i) => ({ text: t, options: i % 2 ? { color: accent } : {} })).filter((r) => r.text);
}
function lightSlide(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: section });
  watermark(s, eyebrow);
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(richTitle(title, K.blue), { placeholder: "title" });
  return s;
}
function darkSlide(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(richTitle(title, K.skyOnDark), { placeholder: "title" });
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
  T(s, "공공기관 용역 입찰\n도우미 웹 서비스", { x: M, y: 1.5, w: 6.2, h: 1.9, fontSize: 40, bold: true,
    color: K.white, lineSpacingMultiple: 1.08 });
  T(s, "공고마다 다른 점수 규칙, 찾아서 금액 범위까지", { x: M, y: 3.5, w: 6.2, h: 0.45, fontSize: 18, color: K.skyOnDark });
  s.addImage({ path: LOGO(true), x: M - 0.05, y: 4.05, w: 0.78 * LOGO_R, h: 0.78, objectName: nm("logo") });
  T(s, "(가칭)", { x: M + 0.78 * LOGO_R + 0.08, y: 4.4, w: 1.2, h: 0.35, fontSize: 14, color: K.onDarkMute });
  s.addImage({ data: await icon("FaScaleBalanced", K.greenOnDark), x: M, y: 4.97, w: 0.3, h: 0.3, objectName: nm("esg-icon") });
  T(s, "공고마다 맞는 규칙을 보여 줘서, 조달 정보 차이를 줄이는 도구", { x: M + 0.42, y: 4.95, w: 5.3, h: 0.36, fontSize: 13, color: K.greenOnDark });
  let cx = M;
  for (const [ic, t] of [["FaBroom", "청소"], ["FaShieldHalved", "경비"], ["FaPeopleGroup", "파견"], ["FaHandshake", "직업소개"]]) {
    await iconCircle(s, ic, cx, 5.6, 0.5, K.skyOnDark, K.navy3);
    T(s, t, { x: cx + 0.58, y: 5.7, w: 1.0, h: 0.3, fontSize: 14, color: K.white });
    cx += t.length > 2 ? 1.75 : 1.35;
  }
  T(s, "Narani", { x: M, y: 6.55, w: 1.5, h: 0.4, fontSize: 16, bold: true, color: K.white });
  chip(s, "회사 대상 이용권형(SaaS)", M + 1.65, 6.58, K.skyOnDark, K.navy3, { w: 2.7, h: 0.34, size: 12 });
  shot(s, "dashboard.jpg", 7.2, 1.05, 5.5, 1530 / 661, "대시보드");
  shot(s, "detail_top.jpg", 6.55, 3.75, 5.65, 1568 / 661, "공고 상세");
  s.addNotes("BidBox는 공공기관 청소·경비·파견·직업소개 용역 입찰에서, 공고에 맞는 점수 계산 규칙을 찾아 써 낼 금액의 범위를 알려 주는 웹 서비스입니다. 회사 대상 이용권형 프로그램입니다. 공고에 맞는 규칙과 가격 기준을 화면에 보여 줘서 중소 용역업체의 조달 정보 차이를 줄이는 것이 이 서비스의 사회적 의미입니다. 오른쪽은 실제 동작하는 시제품 화면입니다.");

  pres.addSection({ title: "P 문제 인식" });

  // ===== P1. 고객과 고객이 겪는 일 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "고객: 중소 용역업체 입찰 담당자, [[공고마다 규칙표를 직접 찾음]]");
  box(s, { x: M, y: 1.75, w: 4.3, h: 4.75, fill: K.bg, line: K.line });
  await iconCircle(s, "FaUserTie", M + 0.3, 1.98, 0.85, K.blue, K.blueTint);
  T(s, "입찰 담당자", { x: M + 1.3, y: 2.02, w: 2.9, h: 0.42, fontSize: 20, bold: true, color: K.ink });
  T(s, "공고를 열고, 얼마를 써 낼지 정하는 사람", { x: M + 1.3, y: 2.46, w: 2.9, h: 0.6, fontSize: 13, color: K.mute });
  T(s, "네 업종 허가(면허)를 가진 중소 용역업체", { x: M + 0.3, y: 3.22, w: 3.8, h: 0.32, fontSize: 13, bold: true, color: K.ink2 });
  const ind = [["FaBroom", "청소(위생관리)"], ["FaShieldHalved", "경비"], ["FaPeopleGroup", "파견"], ["FaHandshake", "직업소개"]];
  for (let i = 0; i < 4; i++) {
    const x = M + 0.3 + (i % 2) * 2.0, y = 3.65 + Math.floor(i / 2) * 0.72;
    await iconCircle(s, ind[i][0], x, y, 0.52, K.white, K.blue);
    T(s, ind[i][1], { x: x + 0.62, y: y + 0.1, w: 1.4, h: 0.34, fontSize: 13, bold: true, color: K.ink });
  }
  T(s, "돈을 내는 사람: 회사 대표 또는 입찰 비용 결재자", { x: M + 0.3, y: 5.2, w: 3.8, h: 0.5, fontSize: 12, color: K.ink2 });
  chip(s, "직무로 특정 · 인터뷰·지불 의사 미확인", M + 0.3, 5.88, K.amber, K.amberTint, { w: 3.7 });
  T(s, "공고 한 건을 볼 때마다", { x: 5.15, y: 1.75, w: 7.5, h: 0.35, fontSize: 15, bold: true, color: K.ink });
  const steps = [["FaFileLines", "본문엔 규칙이 없음", "점수 규칙표(별표)·기준값은 첨부 파일에만"],
    ["FaCodeBranch", "기관·시행일마다 다름", "같은 기준의 옛 판과 고친 판이 함께 있음"],
    ["FaPaperclip", "맞는 판을 직접 고름", "본문과 첨부를 다시 맞춰 보며 판단"],
    ["FaTriangleExclamation", "잘못 고르면 탈락 위험", "규칙이 틀리면 써 낸 금액도 틀림"]];
  const cw = (12.73 - 5.15 - 0.2) / 2;
  for (let i = 0; i < 4; i++) {
    const x = 5.15 + (i % 2) * (cw + 0.2), y = 2.2 + Math.floor(i / 2) * 1.42, red = i === 3;
    box(s, { x, y, w: cw, h: 1.28, fill: red ? K.redTint : K.white, line: red ? "FCA5A5" : K.line });
    T(s, `0${i + 1}`, { x: x + 0.25, y: y + 0.16, w: 0.6, h: 0.3, fontSize: 13, bold: true, color: red ? K.red : K.blue });
    await iconCircle(s, steps[i][0], x + cw - 0.82, y + 0.16, 0.6, red ? K.white : K.blue, red ? K.red : K.blueTint);
    T(s, steps[i][1], { x: x + 0.25, y: y + 0.45, w: cw - 1.2, h: 0.4, fontSize: 16, bold: true, color: red ? K.red : K.ink });
    T(s, steps[i][2], { x: x + 0.25, y: y + 0.86, w: cw - 0.4, h: 0.34, fontSize: 12, color: K.ink2 });
  }
  box(s, { x: 5.15, y: 5.1, w: 7.58, h: 1.4, fill: K.white, line: K.line });
  T(s, "실제 예", { x: 5.4, y: 5.25, w: 1.8, h: 0.3, fontSize: 12, bold: true, color: K.blue });
  T(s, "인천광역시\n계약정보", { x: 5.4, y: 5.55, w: 1.8, h: 0.75, fontSize: 15, bold: true, color: K.ink });
  for (const [x, t] of [[7.3, "2017-04-11본"], [10.1, "2017-12-13 개정본"]]) {
    box(s, { x, y: 5.3, w: 2.4, h: 1.0, fill: K.bg, line: K.line, name: "doc" });
    s.addImage({ data: await icon("FaFileLines", K.mute), x: x + 0.2, y: 5.5, w: 0.45, h: 0.6, objectName: nm("doc-icon") });
    T(s, t, { x: x + 0.8, y: 5.55, w: 1.55, h: 0.5, fontSize: 13, bold: true, color: K.ink });
  }
  s.addImage({ data: await icon("FaCircleQuestion", K.amber), x: 9.83, y: 5.66, w: 0.28, h: 0.28, objectName: nm("q") });
  note(s, "같은 기준의 옛 판과 고친 판이 나란히 게시되어, 어느 판을 쓸지 담당자가 직접 판단 · 이름이 비슷해도 허가(면허)가 다른 업종은 제외");
  s.addNotes("고객은 청소·경비·파견·직업소개 네 업종 중소 용역업체의 입찰 담당자입니다. 공고 한 건을 볼 때마다 첨부 파일에서 규칙표를 찾고, 기관과 시행일에 맞는 판을 직접 골라야 합니다. 인천광역시처럼 옛 판과 고친 판이 함께 올라온 경우도 있습니다.");

  // ===== P2. 문제의 근거: 뉴스·통계 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "[[공사 두 번·용역 한 번]], 1년 새 오른 낙찰하한율");
  box(s, { x: M, y: 1.72, w: CW, h: 1.78, fill: K.white, line: K.line });
  T(s, "낙찰하한율(더 낮게 쓰면 탈락)이 2%p(퍼센트포인트)씩 오른 순서", { x: M + 0.3, y: 1.85, w: 7.6, h: 0.32, fontSize: 13, bold: true, color: K.ink2 });
  s.addShape(pres.shapes.LINE, { x: M + 0.7, y: 2.62, w: CW - 1.4, h: 0, line: { color: K.blue2, width: 3 }, objectName: nm("tl") });
  const evs = [["2025-07-01", "지방계약 공사", "행정안전부"], ["2026-01-30", "국가계약 시설공사", "조달청"], ["2026-05-26 공고부터", "일반용역 · 9년 만", "조달청 공고 제2026-260호"]];
  evs.forEach(([d, t, src], i) => {
    const cx2 = M + CW * (i * 2 + 1) / 6, hot = i === 2;
    s.addShape(pres.shapes.OVAL, { x: cx2 - 0.16, y: 2.46, w: 0.32, h: 0.32, fill: { color: hot ? K.blue : K.amber }, line: { color: hot ? K.blue : K.amber, width: 2.5 }, objectName: nm("ev") });
    T(s, d, { x: cx2 - 1.9, y: 2.82, w: 3.8, h: 0.3, fontSize: 14, bold: true, color: hot ? K.blue : K.amber, align: "center" });
    T(s, `${t} (${src})`, { x: cx2 - 1.9, y: 3.12, w: 3.8, h: 0.3, fontSize: 11, color: K.ink2, align: "center" });
  });
  chip(s, "청소·경비 등 시설분야 87.995% → 89.995%", 8.55, 1.86, K.white, K.blue, { w: 3.95, h: 0.36, size: 12 });
  box(s, { x: M, y: 3.7, w: 6.3, h: 2.8, fill: K.white, line: K.line });
  T(s, "같은 고시 안에서도 분야마다 다른 하한율", { x: M + 0.3, y: 3.85, w: 5.8, h: 0.34, fontSize: 14, bold: true, color: K.ink });
  const lows = [["시설(청소·경비 등)", 89.995], ["소프트웨어(중기 경쟁)", 87.995], ["학술연구(고시금액 이상)", 82.495], ["보험", 47.994]];
  lows.forEach(([n, v], i) => {
    const y = 4.35 + i * 0.47, w = 2.9 * v / 100;
    T(s, n, { x: M + 0.3, y, w: 2.3, h: 0.32, fontSize: 12, color: K.ink2, valign: "middle" });
    box(s, { x: M + 2.6, y: y + 0.03, w, h: 0.26, fill: i === 0 ? K.blue : K.blue2, radius: 0.05, name: "low-bar" });
    T(s, `${v}%`, { x: M + 2.7 + w, y, w: 0.9, h: 0.32, fontSize: 12, bold: true, color: K.ink, valign: "middle" });
  });
  T(s, "조달청 공고 제2026-260호(2026-05-22)", { x: M + 0.3, y: 6.17, w: 5.8, h: 0.26, fontSize: 10, color: K.mute });
  const facts = [["142.1조원", "2024년 공공조달 중 중소기업 수주(63.1%), 고객 대부분이 중소기업", "조달청 보도자료 2025-06-27", K.blue],
    ["1.5%", "사람 자문(1:1 분석)에 맡기면 따낸 금액의 성공보수", "비드큐 안내 (bidq.co.kr)", K.red]];
  facts.forEach(([v, l, src, c], i) => {
    const y = 3.7 + i * 1.45;
    box(s, { x: M + 6.55, y, w: 5.58, h: 1.32, fill: K.white, line: K.line });
    T(s, v, { x: M + 6.8, y: y + 0.2, w: 2.2, h: 0.85, fontSize: 28, bold: true, color: c, valign: "middle" });
    T(s, l, { x: M + 9.05, y: y + 0.16, w: 2.95, h: 0.68, fontSize: 13, color: K.ink2 });
    T(s, src, { x: M + 9.05, y: y + 0.9, w: 2.95, h: 0.28, fontSize: 10, color: K.mute });
  });
  note(s, "출처: 조달경제신문(2026-04-17·2026-05-23), 조달청 공고 제2026-260호, 조달청 보도자료(2025-06-27) · 이 장의 숫자는 외부 자료이며 BidBox 측정값 아님");
  s.addNotes("입찰 규칙은 계속 바뀝니다. 2025년 7월 지방계약 공사, 2026년 1월 국가계약 시설공사, 2026년 5월 26일 공고부터 일반용역의 낙찰하한율이 2%포인트씩 올랐습니다. 청소·경비 같은 시설분야는 87.995%에서 89.995%가 됐고, 9년 만의 조정입니다. 같은 고시 안에서도 분야마다 하한율이 다릅니다. 공공조달의 63.1%는 중소기업이 따내고, 사람 자문에 맡기면 성공보수를 내야 합니다.");

  pres.addSection({ title: "S 실현 가능성" });

  // ===== 9. 불편 → 해결 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "고객의 불편을 [[한 화면에서]] 해결");
  T(s, "고객의 불편", { x: M + 0.25, y: 1.62, w: 4, h: 0.35, fontSize: 14, bold: true, color: K.mute });
  T(s, "BidBox의 해결", { x: 7.25, y: 1.62, w: 4, h: 0.35, fontSize: 14, bold: true, color: K.blue });
  const ps = [["기관·시행일마다 규칙이 달라 어느 판을 쓸지 모름", "FaCodeBranch", "공고 날짜에 맞는 판을 골라 자동 적용"],
    ["써 낼 금액의 범위를 직접 계산", "FaCalculator", "최저가·최상가를 거꾸로 계산, 회사 신용 점수·평판 점수 반영"],
    ["지난 입찰에서 얼마에 따냈는지 따로 찾기", "FaChartLine", "AI 예측가와 예상 낙찰률을 금액 범위 옆에 표시"],
    ["매일 맞는 공고를 직접 고르기", "FaMagnifyingGlass", "네 업종이 목록 기본값, 공고 505,271건 색인"]];
  for (let i = 0; i < 4; i++) {
    const y = 2.05 + i * 1.13;
    box(s, { x: M, y, w: 5.75, h: 0.95, fill: K.amberTint, line: "FDE68A", radius: 0.12 });
    T(s, ps[i][0], { x: M + 0.25, y, w: 5.3, h: 0.95, fontSize: 15, color: K.ink2, valign: "middle" });
    arrow(s, 6.47, y + 0.33, 0.45, K.blue);
    box(s, { x: 7.0, y, w: 5.73, h: 0.95, fill: K.blueTint, radius: 0.12 });
    await iconCircle(s, ps[i][1], 7.18, y + 0.17, 0.6, K.white, K.blue);
    T(s, ps[i][2], { x: 7.95, y, w: 4.65, h: 0.95, fontSize: 15, bold: true, color: K.ink, valign: "middle" });
  }
  note(s, "낙찰 보증 없음. 세 가지 금액과 예측 범위는 참고 정보");
  s.addNotes("고객의 네 가지 불편을 BidBox가 한 화면에서 해결합니다. 맞는 판 적용, 금액 범위 계산, AI 예측가 표시, 네 업종 기본 목록입니다.");

  // ===== 10. 제품 화면 (어두운 띠) =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "[[이미 동작하는]] 시제품 1식");
  shot(s, "detail_top.jpg", M + 0.1, 1.72, 7.8, 1568 / 661, "공고 상세: 맞는 규칙과 세 가지 금액");
  shot(s, "dashboard.jpg", 8.75, 1.72, 3.88, 1530 / 661, "대시보드");
  shot(s, "bids_list.jpg", 8.75, 3.63, 3.88, 1530 / 661, "네 업종 공고 탐색");
  const flow = ["공고 모으기", "네 업종 검색", "맞는 규칙 찾기", "회사 점수 반영", "세 가지 금액", "근거 상담(RAG)", "기록으로 보강"];
  flow.forEach((t, i) => {
    const x = M + i * 1.773;
    s.addText(t, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.2, x, y: 5.6, w: 1.5, h: 0.62,
      fill: { color: i === 2 ? K.blue : K.navy3 }, line: { type: "none" }, fontFace: F, fontSize: 13, bold: true,
      color: K.white, align: "center", valign: "middle", margin: 0, objectName: nm("flow") });
    if (i < 6) arrow(s, x + 1.53, 5.76, 0.2, K.skyOnDark);
  });
  note(s, "실제 프로토타입 화면 (2026-10-22 캡처) · 규칙을 확인하지 못하면 직접 입력", 6.45, true);
  s.addNotes("모두 실제 동작하는 시제품 화면입니다. 공고를 모으고, 네 업종을 검색하고, 맞는 규칙을 찾아 회사 점수를 반영한 세 가지 금액과 근거를 보여 줍니다.");

  // ===== 가져온 공고가 화면까지 (수집 · 다시 배움) =====
  // 근거: 제품 저장소 docker-compose.yml, src/tasks/scheduled_tasks.py, retrain_task.py (2026-11-06).
  // 매일 02:00 수집은 기본 실행에서 켜짐. 주간 재학습은 개발 실행·용역만, 승격은 수동. 운영 재학습은 꺼짐.
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "가져온 공고가 화면까지 [[한 줄로 이어짐]]");
  chip(s, "매일 새벽 2시 · 켜 둠", M, 1.62, K.blue, K.blueTint, { w: 2.55, size: 13 });
  const daily = [
    ["FaCloudArrowDown", "나라장터에서 가져옴", "새 공고와 낙찰 결과"],
    ["FaMagnifyingGlass", "검색 목록에 더함", "방금 가져온 것만"],
    ["FaBookOpen", "상담 자료에 더함", "최근 하루치만"],
    ["FaDisplay", "화면에서 바로 씀", "규칙, 금액, 근거 상담"],
  ];
  const cardW = 2.62, cardGap = 0.55;
  for (let i = 0; i < daily.length; i++) {
    const x = M + i * (cardW + cardGap);
    box(s, { x, y: 2.08, w: cardW, h: 1.82, fill: K.white, line: K.line });
    await iconCircle(s, daily[i][0], x + 0.16, 2.22, 0.46, K.white, K.blue);
    T(s, daily[i][1], { x: x + 0.16, y: 2.78, w: cardW - 0.32, h: 0.58, fontSize: 15, bold: true, color: K.ink });
    T(s, daily[i][2], { x: x + 0.16, y: 3.36, w: cardW - 0.32, h: 0.38, fontSize: 13, color: K.ink2 });
    if (i < daily.length - 1) arrow(s, x + cardW + 0.08, 2.8, 0.38, K.blue2);
  }
  chip(s, "다시 배움 · 개발에서만 켜 둠", M, 4.12, K.amber, K.amberTint, { w: 3.35, size: 13 });
  const learn = [
    ["FaCalendarDays", "매주 월요일 새벽 3시", "용역 모델만"],
    ["FaArrowsRotate", "새 낙찰로 다시 배움", "배울 때와 보여줄 때 계산이 같음"],
    ["FaUserCheck", "더 맞으면 교체를 권함", "바꾸는 것은 사람이 함"],
  ];
  const learnW = 3.7, learnGap = 0.5;
  for (let i = 0; i < learn.length; i++) {
    const x = M + i * (learnW + learnGap);
    box(s, { x, y: 4.56, w: learnW, h: 1.52, fill: K.white, line: "FDE68A" });
    await iconCircle(s, learn[i][0], x + 0.16, 4.72, 0.44, K.white, K.amber);
    T(s, learn[i][1], { x: x + 0.72, y: 4.7, w: learnW - 0.88, h: 0.48, fontSize: 15, bold: true, color: K.ink, valign: "middle" });
    T(s, learn[i][2], { x: x + 0.18, y: 5.32, w: learnW - 0.36, h: 0.52, fontSize: 14, color: K.ink2 });
    if (i < learn.length - 1) arrow(s, x + learnW + 0.06, 5.12, 0.38, "FCD34D");
  }
  note(s, "꺼져 있던 날은 다시 켤 때 한 번 채움. 기록이 예전과 달라지면 알림만 하고, 모델은 바꾸지 않음. 운영에서는 다시 배우기를 꺼 둠.");
  s.addNotes("나라장터 공고와 낙찰 결과를 매일 새벽 2시에 자동으로 가져옵니다. 검색 목록과 상담 자료에는 방금 가져온 것만 더하고, 505,271건 전체를 처음부터 다시 만들지 않습니다. 그날 수집을 놓치면 다시 켤 때 한 번 채웁니다. 예측 모델은 개발 실행에서만 매주 월요일 새벽 3시에 용역을 다시 배웁니다. 배울 때와 화면에 보여줄 때 같은 계산을 씁니다. 더 맞으면 교체를 권하고, 지금 쓰는 모델로 바꾸는 것은 사람이 합니다. 운영 실행에서는 다시 배우기를 꺼 두었습니다. 기록이 예전과 달라지면 알림만 하고 모델은 바꾸지 않습니다.");

  // ===== AI 쉽게 보기 (규칙 · ML · DL) =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "BidBox의 AI, [[세 가지 도구로]] 쉽게 보기");
  const tools = [["FaBookOpen", "규칙표 · AI 아님", K.mute, K.slateTint, "규칙 사전",
      "사람이 공고문과 기준 원문을 읽고 정리한 규칙표 206종. 공고 날짜와 기관을 보고 맞는 쪽을 펼침",
      "맞는 규칙이 있을 때만 금액. 없으면 직접 입력"],
    ["FaTree", "머신러닝 · LightGBM", K.blue, K.blueTint, "예측 나무 숲",
      "지난 낙찰 기록을 공부한 나무 여러 그루가 스무고개처럼 질문을 이어 가며, 앞 나무의 실수를 다음 나무가 고쳐 답을 모음",
      "「아마 이쯤, 이 사이」: 실제 값이 예측 범위 안에 든 비율 90.25%"],
    ["FaRobot", "딥러닝 · RAG", K.blueDeep, K.blueTint, "똑똑한 도서관 사서",
      "질문의 뜻을 알아듣고, 공고·낙찰 기록에서 근거를 먼저 찾아 온 뒤 그 근거로만 답함",
      "근거가 없으면 답하지 않음: 거절해야 할 질문 24/24 거절"]];
  const tw = (CW - 2 * 0.3) / 3;
  for (let i = 0; i < 3; i++) {
    const [ic, tag, c, tint, t, d, out] = tools[i], x = M + i * (tw + 0.3);
    box(s, { x, y: 1.75, w: tw, h: 4.05, fill: K.white, line: K.line, shadow: true });
    await iconCircle(s, ic, x + 0.3, 2.0, 0.95, K.white, c);
    chip(s, tag, x + tw - 0.3 - (tag.length * 0.135 + 0.35), 2.12, c, tint);
    T(s, t, { x: x + 0.3, y: 3.1, w: tw - 0.6, h: 0.45, fontSize: 20, bold: true, color: K.ink });
    T(s, d, { x: x + 0.3, y: 3.6, w: tw - 0.6, h: 1.2, fontSize: 13, color: K.ink2 });
    s.addText(out, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x + 0.2, y: 4.9, w: tw - 0.4, h: 0.72,
      fill: { color: tint }, line: { type: "none" }, fontFace: F, fontSize: 12, bold: true, color: i === 0 ? K.ink2 : K.blueDeep,
      align: "center", valign: "middle", margin: 0.06, objectName: nm("tool-out") });
  }
  s.addText([{ text: "규칙", options: { bold: true, color: K.ink } }, { text: "은 사람이 확인한 표로   ·   ", options: { color: K.ink2 } },
    { text: "금액", options: { bold: true, color: K.blue } }, { text: "은 데이터로 배운 모델로   ·   ", options: { color: K.ink2 } },
    { text: "설명", options: { bold: true, color: K.blueDeep } }, { text: "은 근거를 찾는 언어 모델로", options: { color: K.ink2 } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 6.0, w: CW, h: 0.6, fill: { color: K.blueTint }, line: { type: "none" },
      fontFace: F, fontSize: 15, align: "center", valign: "middle", margin: 0, objectName: nm("tools-sum") });
  s.addNotes("BidBox는 세 가지 도구를 씁니다. 규칙 사전은 AI가 아니라 사람이 원문을 읽고 정리한 규칙표 206종입니다. 맞는 규칙이 있을 때만 금액을 내고, 없으면 직접 입력으로 넘깁니다. 예측 나무 숲은 머신러닝 LightGBM으로, 지난 낙찰 기록을 공부한 나무들이 스무고개처럼 질문을 이어 가며 금액과 범위를 냅니다. 실제 값이 범위 안에 든 비율은 90.25%입니다. 똑똑한 도서관 사서는 딥러닝 언어 모델과 검색을 함께 쓰는 RAG로, 근거를 먼저 찾고 그 근거로만 답합니다.");

  // ===== RAG 인텔리전스 (실제 화면 포함) =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "근거 상담(RAG): [[없으면 답하지 않음]]");
  // 왼쪽: AI 어드바이저 실제 화면(2026-11-05 로컬 실행 캡처, 사용자 이름이 보이는 사이드바는 잘라 냄)
  const advW = 6.0, advH = advW / (1015 / 468);
  box(s, { x: M - 0.1, y: 1.62, w: advW + 0.2, h: advH + 0.2, fill: K.navy3, radius: 0.16, shadow: true, name: "adv-frame" });
  s.addImage({ path: CAP("ai_advisor_rag.jpg"), x: M, y: 1.72, w: advW, h: advH, objectName: nm("adv-shot") });
  T(s, "실제 화면: 찾은 문서 [1][3][4][5]에 답이 없자 지어내지 않고 거절", { x: M, y: 1.72 + advH + 0.18, w: advW + 0.4, h: 0.32, fontSize: 13, bold: true, color: K.skyOnDark });
  // 오른쪽: 5단계 흐름(세로)
  const rag = [["FaComments", "질문", "공고·낙찰·규칙을 묻는 질문"],
    ["FaListCheck", "뜻 파악", "공고·숫자·규칙 가운데 무엇을 묻는지 나눔"],
    ["FaMagnifyingGlass", "세 길로 찾기", "비슷한 뜻 · 제목 글자 · 숫자는 저장 자료에서"],
    ["FaDatabase", "근거 모으기", "최근 1년 공고와 낙찰 결과를 모아 둔 자료"],
    ["FaRobot", "근거로 답하기", "서비스 안 언어 모델이 출처를 달고, 없으면 거절"]];
  const fx = M + advW + 0.45, fw = CW - advW - 0.45;
  for (let i = 0; i < 5; i++) {
    const y = 1.62 + i * 0.7, hot = i === 2 || i === 4;
    box(s, { x: fx, y, w: fw, h: 0.6, fill: hot ? K.blue : K.navy2, line: hot ? K.blue : K.navy3, radius: 0.12 });
    await iconCircle(s, rag[i][0], fx + 0.1, y + 0.08, 0.44, hot ? K.blue : K.skyOnDark, hot ? K.white : K.navy3);
    T(s, `${i + 1}. ${rag[i][1]}`, { x: fx + 0.65, y, w: 1.55, h: 0.6, fontSize: 13, bold: true, color: K.white, valign: "middle" });
    T(s, rag[i][2], { x: fx + 2.2, y, w: fw - 2.3, h: 0.6, fontSize: 11, color: hot ? K.white : K.onDarkMute, valign: "middle" });
  }
  // 아래: 품질 숫자
  const rq = [["144/144", "숫자 정답"], ["1.0", "근거를 다시 찾음"], ["72/72", "출처 인용"], ["24/24", "거절해야 할 질문"], ["0건", "근거 없는 답"], ["2.66초", "가장 느린 쪽 첫 글자"]];
  const qw = (CW - 5 * 0.15) / 6;
  rq.forEach(([v, l], i) => {
    const x = M + i * (qw + 0.15);
    box(s, { x, y: 5.35, w: qw, h: 1.0, fill: K.navy2, line: K.navy3, radius: 0.12 });
    T(s, v, { x: x + 0.15, y: 5.42, w: qw - 0.3, h: 0.5, fontSize: 21, bold: true, color: K.skyOnDark });
    T(s, l, { x: x + 0.15, y: 5.92, w: qw - 0.3, h: 0.35, fontSize: 10.5, color: K.onDarkMute });
  });
  note(s, "숫자: 96개 요청 평가(제품 저장소 RAG 품질 정본), 첫 토큰 P95 2,664.19ms · 화면: 로컬 실행 캡처(2026-11-05) · 서비스 안 모델이 기본, 공고 원문을 외부 모델로 보내지 않음", 6.5, true);
  s.addNotes("RAG는 질문에 바로 답하지 않고 근거를 먼저 찾는 방식입니다. 왼쪽은 실제 화면입니다. 낙찰하한율의 뜻을 물었는데 찾은 문서 1, 3, 4, 5번에 그 내용이 없자, 지어내지 않고 답할 수 없다고 했습니다. 질문의 뜻을 나눈 뒤 뜻으로 찾기, 제목 글자로 찾기, 숫자는 데이터베이스에서 바로 조회하는 세 길로 근거를 모으고, 서비스 안의 언어 모델이 출처를 달아 답합니다. 96개 요청 평가에서 숫자 정답 144/144, 출처 인용 72/72, 거절해야 할 질문 24/24를 지켰습니다.");

  // ===== 11. 세 가지 금액 + 확인된 규칙만 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "세 가지 금액을 한 화면에, [[확인된 규칙만 계산]]");
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
  T(s, "확인하지 못한 공고는 금액을 비우고 직접 입력", { x: 7.5, y: 4.55, w: 5.2, h: 0.35, fontSize: 13, bold: true, color: K.red });
  const stops = ["금액 없는 공고에 0원 추천 없음", "기준 금액 없는 공고(비예가)는 가격점수 계산 안 함", "협상 계약은 평가 비율·낙찰률 참고 분포"];
  for (let i = 0; i < 3; i++) {
    s.addImage({ data: await icon("FaCircleStop", K.red), x: 7.5, y: 5.05 + i * 0.47, w: 0.26, h: 0.26, objectName: nm("stop") });
    T(s, stops[i], { x: 7.9, y: 5.01 + i * 0.47, w: 4.85, h: 0.34, fontSize: 13, color: K.ink2 });
  }
  s.addNotes("최저가, AI 예측가, 최상가를 한 화면에 보여 줍니다. 규칙을 확인하지 못한 공고는 금액을 비우고 직접 입력하게 합니다. 금액이 없는 공고에 0원을 추천하지 않습니다.");

  // ===== 12. 경쟁 비교 OX =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "기존 방식과 비교: [[규칙표 고르기와 확인된 계산]]");
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
    ["기준 없을 때", "—", "✕ 넣은 값 그대로 계산", "—", "○ 확인된 규칙만, 없으면 직접 입력"],
    ["세 가지 금액", "△ 검색·제출 중심", "△ 점수만 계산", "○ 금액 제안", "○ 최저·예측·최상 한 화면"],
    ["근거 상담(RAG)", "—", "✕ 결과 점수 중심", "—", "○ 근거 찾아 답·출처 인용"],
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
  s.addNotes("강의 팁대로 줄글 대신 OX로 비교했습니다. 규칙표 고르기, 확인된 규칙만 계산하기, 근거 표시는 BidBox만 지원합니다. 확인하지 못한 칸은 미확인으로 두었습니다.");

  // ===== 13. 개발 현황 =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "이미 만든 것, [[숫자로 확인]]");
  const tiles = [["206종", "규칙 ID", "기관·시행일별 계산 규칙"], ["21,394건", "자동 적용", "31,305건 중 (2026-10-19)"],
    ["505,271건", "공고 색인", "최근 1년 (2026-09-09)"], ["1.1885%p", "낙찰률 평균 오차", "3,000건 (2026-09-15)"],
    ["90.25%", "예측 범위 포함률", "3,589건 (2026-08-30)"], ["144/144", "상담 숫자 정답", "96개 요청 평가"]];
  tiles.forEach(([v, l, d], i) => {
    const x = M + (i % 3) * 2.83, y = 1.75 + Math.floor(i / 3) * 2.42;
    box(s, { x, y, w: 2.65, h: 2.22, fill: K.navy2, line: K.navy3 });
    T(s, v, { x: x + 0.25, y: y + 0.3, w: 2.3, h: 0.7, fontSize: 28, bold: true, color: K.skyOnDark });
    T(s, l, { x: x + 0.25, y: y + 1.1, w: 2.3, h: 0.4, fontSize: 15, bold: true, color: K.white });
    T(s, d, { x: x + 0.25, y: y + 1.5, w: 2.3, h: 0.4, fontSize: 11, color: K.onDarkMute });
  });
  box(s, { x: 9.2, y: 1.75, w: 3.53, h: 4.64, fill: K.navy2, line: K.navy3 });
  T(s, "개발 진행 상태 (10개 항목)", { x: 9.45, y: 1.97, w: 3.1, h: 0.4, fontSize: 14, bold: true, color: K.white });
  const stat = [[7, "구현 완료", K.greenOnDark, K.navy3], [1, "부분 구현", K.amberOnDark, K.navy3], [2, "미구축 · 협약 중 구축", K.redOnDark, K.navy3]];
  let sq = 0;
  for (const [n, , c] of stat) for (let k = 0; k < n; k++, sq++)
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.45 + sq * 0.305, y: 2.5, w: 0.26, h: 0.42, rectRadius: 0.05, fill: { color: c }, line: { type: "none" }, objectName: nm("sq") });
  stat.forEach(([n, l, c, t], i) => {
    s.addText(String(n), { shape: pres.shapes.OVAL, x: 9.45, y: 3.2 + i * 0.6, w: 0.42, h: 0.42, fill: { color: t }, line: { type: "none" },
      fontFace: F, fontSize: 14, bold: true, color: c, align: "center", valign: "middle", margin: 0, objectName: nm("count") });
    T(s, l, { x: 10.0, y: 3.24 + i * 0.6, w: 2.6, h: 0.36, fontSize: 13, color: K.onDarkMute });
  });
  T(s, "협약 중 만들 것\n· 운영 주소, 이용권 결제, 약관, 보안 점검\n· 시험 사용 기록, 지자체 원문", { x: 9.45, y: 4.95, w: 3.15, h: 0.8, fontSize: 12, color: K.onDarkMute });
  s.addText([{ text: "특허 출원 1건 (7개월차)", options: { bold: true, color: K.skyOnDark, breakLine: true } },
    { text: "규칙표 판별과 금액 범위를 거꾸로 계산하는 절차", options: { color: K.white } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: 9.4, y: 5.75, w: 3.13, h: 0.5, fill: { color: K.navy3 }, line: { type: "none" },
      fontFace: F, fontSize: 11, valign: "middle", margin: [0, 0.08, 0, 0.1], objectName: nm("patent") });
  s.addNotes("규칙 206종, 31,305건 중 21,394건 자동 적용, 낙찰률 평균 오차 1.1885%p, 상담 숫자 정답 144/144입니다. 협약 기간에는 운영 공개와 결제, 시험 사용을 만듭니다.");

  // ===== 14. 협약 8개월 일정 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "협약 8개월: 운영 공개에서 [[유료 전환까지]]");
  const half = (CW - 0.2) / 2;
  s.addText([{ text: "1단계 · 1~4개월   ", options: { bold: true, color: K.blue } }, { text: "2,000만 원", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 1.72, w: half, h: 0.6, fill: { color: K.blueTint }, line: { type: "none" },
      fontFace: F, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: nm("phase") });
  s.addText([{ text: "2단계 · 5~8개월   ", options: { bold: true, color: K.ink2 } }, { text: "2,000만 원 (중간평가 통과 시)", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + half + 0.2, y: 1.72, w: half, h: 0.6, fill: { color: K.white }, line: { color: K.line, width: 1 },
      fontFace: F, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: nm("phase") });
  s.addShape(pres.shapes.LINE, { x: M + 0.4, y: 3.25, w: CW - 0.8, h: 0, line: { color: K.line, width: 3 }, objectName: nm("tl") });
  const ms = [["운영 공개", "서버·도메인·보안 접속, Windows 실행 확인"], ["회사 점수 재현", "등급 넣은 계정과 비교, 약관 초안"],
    ["이용권 결제", "1,000원(가설), 보안 점검"], ["시험 사용 1차", "10사 사용 기록, 중간평가 자료"],
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

  pres.addSection({ title: "S 성장전략" });

  // ===== 16. 시장 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "목표 시장: 55.1조원에서 [[네 업종 7,494억원부터]]");
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
  s = lightSlide("S 성장전략", "S · 성장전략", "비즈니스 모델: 검색은 무료, [[판단은 이용권으로]]");
  chip(s, "회사 대상", 11.55, 0.4, K.blue, K.blueTint, { w: 1.18, size: 12 });
  const hd = [["누가 내나", M], ["무엇을 받나", 4.75], ["어떻게, 언제", 9.35]];
  for (const [t, x] of hd) T(s, t, { x: x + 0.2, y: 1.6, w: 3, h: 0.3, fontSize: 12, bold: true, color: K.mute });
  const bm = [["무료 입구", "모든 방문자", "네 업종 검색, 대시보드, 기준값 직접 입력 계산", "무료 · 지금(구현됨)", "FaGift", 0],
    ["① 주수익", "네 업종 중소기업\n(대표·입찰 비용 결재자)", "규칙 자동 찾기, 세 가지 금액, 근거 상담", "이용권(토큰) 1개 1,000원\n협약 3개월차 결제 시작", "FaTicket", 2],
    ["② 부수익", "입찰 대행·교육 기관", "고객 회사 대신 쓰는 묶음 이용권", "묶음 단가는 사용 확인 뒤\n3년차", "FaLayerGroup", 1],
    ["③ 부수익", "담당자가 여럿이거나\n자기 시스템에 붙이는 회사", "계정 추가, 프로그램 연결(API)", "협약 중 가격 미공개\n4년차(계획)", "FaPlug", 1]];
  for (let i = 0; i < 4; i++) {
    const [tag, who2, what, how, ic, lv] = bm[i], y = 1.95 + i * 0.95, main = lv === 2;
    box(s, { x: M, y, w: 3.85, h: 0.8, fill: K.bg, line: K.line, radius: 0.12 });
    chip(s, tag, M + 0.15, y + 0.24, main ? K.white : (lv ? K.blue : K.mute), main ? K.blue : (lv ? K.blueTint : K.slateTint), { w: 0.95, size: 11 });
    T(s, who2, { x: M + 1.22, y, w: 2.55, h: 0.8, fontSize: 13, bold: true, color: K.ink, valign: "middle" });
    if (lv) arrow2(s, 4.47, y + 0.25, 0.42, main ? K.blue : K.blue2); else arrow(s, 4.5, y + 0.25, 0.3, K.blue2);
    box(s, { x: 4.95, y, w: 4.25, h: 0.8, fill: main ? K.blueTint : K.white, line: main ? K.blue : K.line, lw: main ? 1.5 : 1, radius: 0.12 });
    await iconCircle(s, ic, 5.1, y + 0.12, 0.56, main ? K.white : K.blue, main ? K.blue : K.blueTint);
    T(s, what, { x: 5.8, y, w: 3.3, h: 0.8, fontSize: 13, color: K.ink, valign: "middle", bold: main });
    arrow(s, 9.27, y + 0.25, 0.3, K.blue2);
    box(s, { x: 9.65, y, w: 3.08, h: 0.8, fill: main ? K.blue : K.white, line: main ? K.blue : K.line, radius: 0.12 });
    T(s, how, { x: 9.8, y, w: 2.85, h: 0.8, fontSize: 12, bold: main, color: main ? K.white : K.ink2, valign: "middle" });
  }
  s.addText([{ text: "유료 8사 × 연 100이용권(토큰) × 1,000원 = 연 80만 원 (가설)", options: { bold: true, color: K.ink } },
    { text: "   손익이 아니라 지불 의사를 확인하는 매출", options: { color: K.ink2 } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 5.85, w: 9.6, h: 0.6, fill: { color: K.bg }, line: { color: K.line, width: 1 },
      fontFace: F, fontSize: 14, align: "center", valign: "middle", margin: 0, objectName: nm("formula") });
  chip(s, "낙찰 성공보수 없음", 10.4, 5.85, K.red, K.redTint, { w: 2.33, h: 0.6, size: 14 });
  note(s, "⇄ 현금 흐름: 이용권 결제는 고객 → BidBox, 세 가지 금액·근거·상담은 BidBox → 고객 · 무료 입구는 결제 없이 제공");
  s.addNotes("검색과 대시보드는 무료이고, 규칙 찾기와 세 가지 금액, 상담은 이용권으로 씁니다. 1토큰 1,000원, 유료 8사가 연 100토큰을 사면 연 80만 원입니다. 이 값은 지불 의사를 보는 가설입니다. 낙찰 성공보수는 받지 않습니다.");

  // ===== 영업·홍보 전략 (강의 38~40쪽) =====
  s = lightSlide("S 성장전략", "S · 성장전략", "영업·홍보: [[무료로 써 보고, 시험 사용으로]] 모집");
  box(s, { x: M, y: 1.75, w: 5.95, h: 4.2, fill: K.white, line: K.line, shadow: true });
  chip(s, "01 온라인", M + 0.3, 1.98, K.white, K.blue, { w: 1.3, h: 0.36, size: 13 });
  T(s, "무료 입구와 소개 페이지로 첫 접점", { x: M + 1.75, y: 1.98, w: 4.0, h: 0.36, fontSize: 15, bold: true, color: K.ink, valign: "middle" });
  box(s, { x: M + 0.25, y: 2.5, w: 5.45, h: 5.45 / (1530 / 661) + 0.1, fill: K.bg, line: K.line, radius: 0.06, name: "web-frame" });
  s.addImage({ path: CAP("bids_list.jpg"), x: M + 0.3, y: 2.55, w: 5.35, h: 5.35 / (1530 / 661), objectName: nm("web-shot") });
  const onl = [["FaGift", "무료 입구: 네 업종 검색·대시보드·직접 계산", "지금(구현됨)"], ["FaMagnifyingGlass", "소개 페이지 180만 원 · 시험 사용 모집 300만 원", "4개월차"]];
  for (let i = 0; i < 2; i++) {
    const y = 4.97 + i * 0.47;
    await iconCircle(s, onl[i][0], M + 0.3, y, 0.38, K.blue, K.blueTint);
    T(s, onl[i][1], { x: M + 0.8, y: y + 0.02, w: 3.9, h: 0.34, fontSize: 13, color: K.ink2 });
    T(s, onl[i][2], { x: M + 4.55, y: y + 0.02, w: 1.15, h: 0.34, fontSize: 11, color: K.mute, align: "right" });
  }
  const ox = M + 6.18;
  box(s, { x: ox, y: 1.75, w: 5.95, h: 4.2, fill: K.white, line: K.line, shadow: true });
  chip(s, "02 오프라인", ox + 0.3, 1.98, K.white, K.blueDeep, { w: 1.45, h: 0.36, size: 13 });
  T(s, "입찰 실무자가 모이는 자리에서 직접 설명", { x: ox + 1.9, y: 1.98, w: 3.9, h: 0.36, fontSize: 15, bold: true, color: K.ink, valign: "middle" });
  const off = [["FaGraduationCap", "입찰 교육·협회 설명회", "2단계(5~8개월차) · 광고선전비 1,201,640원", "계획(안)"],
    ["FaHandshake", "입찰 대행·지역 교육 기관에 제안", "계산식 확인 도구로 · 4개월까지 2곳과 시험 사용 자리", "확정 전"],
    ["FaUserTie", "영업 담당이 시험 사용 회사 모집", "협약 5개월차 채용 · 입찰 실무 검수", "채용 예정"]];
  for (let i = 0; i < 3; i++) {
    const y = 2.6 + i * 1.08;
    box(s, { x: ox + 0.25, y, w: 5.45, h: 0.92, fill: K.bg, line: K.line, radius: 0.12 });
    await iconCircle(s, off[i][0], ox + 0.4, y + 0.16, 0.6, K.white, K.blueDeep);
    T(s, off[i][1], { x: ox + 1.15, y: y + 0.1, w: 3.2, h: 0.36, fontSize: 15, bold: true, color: K.ink });
    T(s, off[i][2], { x: ox + 1.15, y: y + 0.48, w: 4.45, h: 0.34, fontSize: 12, color: K.ink2 });
    chip(s, off[i][3], ox + 5.6 - (off[i][3].length * 0.135 + 0.35), y + 0.12, K.amber, K.amberTint);
  }
  s.addText([{ text: "무료로 써 보기  →  시험 사용 10사  →  20사  →  유료 8사 (가설)", options: { bold: true, color: K.ink } },
    { text: "    낙찰 대행은 목표 아님", options: { color: K.ink2 } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 6.12, w: CW, h: 0.58, fill: { color: K.blueTint }, line: { type: "none" },
      fontFace: F, fontSize: 15, align: "center", valign: "middle", margin: 0, objectName: nm("sales-flow") });
  s.addNotes("온라인은 지금 동작하는 무료 검색과 대시보드를 입구로 쓰고, 소개 페이지와 온라인 모집으로 시험 사용 회사를 모읍니다. 오프라인은 입찰 교육·협회 설명회, 입찰 대행·교육 기관 제안, 영업 담당의 직접 모집입니다. 확정된 거래처는 아직 없습니다.");

  // ===== 18. 고객 검증·진입 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "고객 검증: [[시험 사용으로 지불 의사 확인]]");
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
  s = lightSlide("S 성장전략", "S · 성장전략", "5개년 로드맵: [[용역 → 물품 → 공사 → 자체조달]]");
  const rd = [["1년차", "네 업종 상용화", "0.75조원", K.blueTint, K.ink],
    ["2년차", "용역 적격심사 전 업종\n물품 착수", "10.2조원", "CCE3FA", K.ink],
    ["3년차", "물품 공개 · 공사 착수\n묶음 이용권", "23.4조원", K.blue2, K.ink],
    ["4년차 (계획)", "공사 공개\n여러 사용자·API", "54.4조원", K.blue, K.white],
    ["5년차 (계획)", "나라장터 밖\n자체조달 기관 검토", "+80.0조원", K.blueDeep, K.white]];
  // 강의 50~52쪽: 사업 계획과 기술 계획을 함께, 왼쪽→오른쪽 상승
  const tech = ["지자체 규칙 보강\n사용 기록으로 규칙 보강", "기술용역 낙찰방법 44종 정리\n물품 규칙·모델 착수", "공사 규칙·모델 착수", "여러 사용자 계정\n프로그램 연결(API)", "자체조달 공고 연동\n기관별 검토"];
  const X0 = M + 1.0, sw = (CW - 1.0 - 4 * 0.12) / 5;
  chip(s, "사업", M, 3.95, K.white, K.blue, { w: 0.8, h: 0.4, size: 13 });
  chip(s, "기술", M, 5.37, K.white, K.navy3, { w: 0.8, h: 0.4, size: 13 });
  rd.forEach(([y1, t, v, f, c], i) => {
    const h = 0.95 + i * 0.45, x = X0 + i * (sw + 0.12), y = 5.0 - h;
    box(s, { x, y, w: sw, h, fill: f, radius: 0.1, name: "step" });
    T(s, y1, { x: x + 0.15, y: y + 0.1, w: sw - 0.3, h: 0.32, fontSize: 13, bold: true, color: c });
    T(s, v, { x: x + 0.15, y: y + h - 0.5, w: sw - 0.3, h: 0.42, fontSize: 18, bold: true, color: c });
    T(s, t, { x, y: y - 0.65, w: sw, h: 0.58, fontSize: 12, color: K.ink2, align: "center", valign: "bottom" });
    box(s, { x, y: 5.15, w: sw, h: 0.85, fill: K.white, line: K.line, radius: 0.1, name: "tech" });
    T(s, tech[i], { x: x + 0.08, y: 5.15, w: sw - 0.16, h: 0.85, fontSize: 11, color: K.ink, align: "center", valign: "middle" });
  });
  note(s, "누적 공략 범위 = 고객이 따는 계약 금액(BidBox 매출 아님) · 물품은 2025 입찰 낙찰 전체 13.2조원, 공사 적격심사 31.0조원, 자체조달 80.0조원(2024) · 4~5년차는 계획(안)", 6.48);
  s.addNotes("1년차 네 업종으로 시작해 2년차 용역 전 업종, 3년차 물품, 4년차 공사, 5년차 자체조달 기관으로 넓힙니다. 기술은 지자체 규칙 보강에서 시작해 낙찰방법 정리, 물품·공사 규칙과 모델로 이어집니다. 공략 범위는 고객이 따는 계약 금액입니다.");

  // ===== 20. 자금 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "자금: 정부지원 [[4,000만 원]], 매출은 지불 의사 확인용");
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
    ["매출", "연 80만 원", "유료 8사 × 연 100이용권(토큰) (가설)", K.amber], ["엔젤·시드", "미확인", "협약 중 투자 유치 일정 없음", K.mute]];
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
  s = lightSlide("S 성장전략", "S · 성장전략 · ESG", "사회적 가치: [[모두가 같은 정보로 입찰]]");
  const esg = [["S", "FaScaleBalanced", "조달 정보 차이 줄이기", "규칙표·가격 기준을 공고 화면에 공개, 검색·대시보드·직접 계산 무료", "2024 중소기업 수주 142.1조원(63.1%)", "1년차 시험 사용 20사 중 비수도권 10사 이상"],
    ["S", "FaUsers", "약자기업·고용 우수기업", "여성·장애인기업, 청년·여성·장애인 고용 등 가점은 지금 계산. 안내 화면은 2년차", "적격심사 회사 평판 항목, 규칙에 구현", "2년차 가점 안내 화면"],
    ["G", "FaUserShield", "정보 보호·책임 있는 AI", "필요한 항목만 수집, 상담은 서비스 안 모델, 글자 읽기(OCR) 기본 꺼짐", "약관 첫머리에 낙찰 비보증 명시", "1년차부터"]];
  const ew = (CW - 2 * 0.3) / 3;
  for (let i = 0; i < 3; i++) {
    const [L, ic, t, d, g, when] = esg[i], x = M + i * (ew + 0.3);
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
  note(s, "지표의 숫자 목표는 시험 사용 측정 뒤 결정 · 약자기업 우대 단가는 확인 전까지 수치 미설정");
  s.addNotes("사회적 가치는 공공조달 정보 차이를 줄이는 데서 나옵니다. 공고에 맞는 규칙표와 가격 기준을 보여 줍니다. 약자기업·고용 우수기업 가점은 지금 계산에 반영하고, 받을 수 있는 가점 안내 화면은 2년차입니다. 정보 보호와 낙찰 비보증을 지킵니다.");

  pres.addSection({ title: "T 팀 구성" });

  // ===== 23. 팀 =====
  s = lightSlide("T 팀 구성", "T · 팀 구성", "대표자가 [[서비스 전체를 직접 구현]]");
  box(s, { x: M, y: 1.75, w: 6.3, h: 4.75, fill: K.bg, line: K.line });
  await iconCircle(s, "FaUserTie", M + 0.35, 2.05, 1.0, K.white, K.blue);
  T(s, "대표자", { x: M + 1.6, y: 2.1, w: 3, h: 0.45, fontSize: 22, bold: true, color: K.ink });
  T(s, "기획 · 개발 · 데이터 검증", { x: M + 1.6, y: 2.6, w: 4.3, h: 0.4, fontSize: 14, color: K.mute });
  const rep = [["FaBriefcase", "경력", "전시컨벤션 사업 매니저 6년 9개월\n국제회의·기업행사 기획·진행, 사업입찰 진행 포함"],
    ["FaRobot", "기술", "규칙 206종 정리, 낙찰률 예측 모델 학습·운영,\n네 업종 화면과 근거 상담을 한 시제품으로 연결"],
    ["FaListCheck", "원칙", "확인된 규칙만 계산, 없으면 직접 입력. 원인→시도→해결→검증 기록"],
    ["FaDatabase", "인프라", "동작하는 시제품 1식, 나라장터 공고 505,271건 색인\n서버는 사지 않고 임대(협약 중)"]];
  for (let i = 0; i < 4; i++) {
    const y = 3.2 + i * 0.8;
    await iconCircle(s, rep[i][0], M + 0.35, y, 0.5, K.blue, K.blueTint);
    T(s, rep[i][1], { x: M + 1.0, y: y + 0.08, w: 0.85, h: 0.35, fontSize: 14, bold: true, color: K.blue });
    T(s, rep[i][2], { x: M + 1.85, y: y - 0.03, w: 4.3, h: 0.75, fontSize: 13, color: K.ink2 });
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
  s.addImage({ path: LOGO(true), x: M - 0.03, y: 0.95, w: 0.5 * LOGO_R, h: 0.5, objectName: nm("logo") });
  T(s, "확인된 규칙만\n계산하는 도우미", { x: M, y: 1.7, w: 10.5, h: 1.9, fontSize: 44, bold: true, color: K.white, lineSpacingMultiple: 1.05 });
  const fin = [["206종", "공고에 맞는 계산 규칙"], ["31,305건", "최근 12개월 규칙 판정"], ["0원", "금액 없는 공고 추천"]];
  fin.forEach(([v, l], i) => {
    const x = M + i * 4.1;
    box(s, { x, y: 4.05, w: 3.85, h: 1.35, fill: K.navy2, line: K.navy3, radius: 0.16 });
    T(s, v, { x: x + 0.3, y: 4.2, w: 3.3, h: 0.65, fontSize: 30, bold: true, color: K.skyOnDark });
    T(s, l, { x: x + 0.3, y: 4.88, w: 3.3, h: 0.35, fontSize: 14, color: K.onDarkMute });
  });
  T(s, "감사합니다 · 질문 받겠습니다", { x: M, y: 5.9, w: 8, h: 0.5, fontSize: 22, bold: true, color: K.white });
  T(s, "Narani · 낙찰 보증 없음, 세 가지 금액은 참고 정보", { x: M, y: 6.5, w: 8, h: 0.35, fontSize: 12, color: K.onDarkMute });
  s.addNotes("BidBox는 확인된 규칙만 계산하는 입찰 도우미입니다. 규칙을 확인하지 못한 공고는 금액을 비우고 직접 입력하게 합니다. 감사합니다. 질문 받겠습니다.");

  await pres.writeFile({ fileName: OUT });
  if (APPLY) { const { applyTheme } = require(APPLY); await applyTheme(OUT, THEME); }
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
