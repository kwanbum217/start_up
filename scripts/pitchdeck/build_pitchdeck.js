// BidBox 피치덱(.pptx) 생성기
// 수치·사실: AGENTS.md 「현재 사업계획서」와 「피치덱」. 새 숫자 없음.
// 여백·글자 크기·카드 간격·명암: .agents/skills/bizplan-writing/references/pitch-design.md (v11부터).
// v19 (2026-11-10): 문장·숫자·노트는 v18과 같고 디자인만 정리한다. 워터마크·카드 뒤 판(plate)을 걷어
// 잘린 조각과 어긋난 겹침을 없애고, 카드를 한 격자에 맞춘다. 시장(넓이 비례 겹친 원, 같은 출발선 막대),
// 자금(오른쪽 숫자 칸), 검증(협약 중 확보 칸), 개발 현황(숫자 원 높이)을 다시 그리고, 문제 근거·마무리는 상자 없이 선으로 나눈다.
// v18 (2026-11-10): 장 화면은 v17과 같고, 발표자 노트만 7회차 강의와 10분 발표에 맞춘다.
// 장마다 말할 문장(합계 약 1,950자, 분당 250~300자로 7~8분, 장 넘김 포함 9분 안팎)을 앞에, 뺀 설명은 「■ 질의응답 참고」 뒤에 둔다. 첫 인사·마무리에 회사 이름과
// 「[이름]」 자리(저장소가 공개라 실명 없음), 장 사이 연결 말, 10장 작성 메모 삭제. 새 숫자 없음. v17은 덮어쓰지 않는다.
// v17 (2026-11-10): 장 위 문장을 줄이고, 강조 한 줄·큰 숫자·제품 화면·도표가 장을 채운다.
// 뺀 설명은 발표자 노트에 둔다. 숫자·출처·가설·낙찰 보증 없음은 장에 남긴다. v16은 덮어쓰지 않는다.
// v16 (2026-11-09): v15의 파랑 횟수는 유지하고, 6회차 우수사례처럼 제목·카드 제목은 굵게 둔다.
// 본문·눈썹·칩은 보통. 제출 글꼴은 맑은 고딕. 검토용 구글 슬라이드만 Noto Sans KR.
// v15 (2026-11-09): v14 배치·문장·숫자·모서리·그림자를 유지하고, 파랑은 제목 핵심어와
// 큰 숫자에만 두고 그 밖 굵기를 뺀다. 비교 시안 시안_강조절제_20261109와 같다.
// v14: v13 배치를 유지하고, 구글 화면에서 안 읽히거나 겹친 자리만 고친다.
// 3년차 머리는 흰 글자가 읽히는 파랑, 마무리 숫자 카드는 서로 떨어짐,
// 고객·팀 장의 배경 판은 앞 카드와 겹치지 않음. 문장·숫자는 v13과 같다.
// 구성: 6회차 강의(P·S·S·T, 장 제목 필수, 줄글 대신 숫자·OX·도식). 발표용이라 목차·간지 없음(강의 6쪽 「피칭 X」).
// 디자인: BidBox 웹 화면 색(primary #0073E6, sidebar #020B1A, slate) + Coinbase식 단일 강조색·어두운 띠 위 제품 화면
//         + IBM Carbon식 상태 3색(초록·주황·빨강). 강의 예시의 인물·현장 사진은 넣지 않는다.
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
// v16부터 제목·카드 제목은 굵게, 본문은 보통이다. 파랑 횟수는 v15와 같다.
// PITCH_QUIET=0 이면 v14 모습(파랑·굵기 모두)으로 다시 만든다.
// PITCH_FLAT=1: 채택하지 않은 비교 시안. 문장·숫자는 그대로 두고 모서리·그림자·바탕만 바꾼다.
const FLAT = process.env.PITCH_FLAT === "1";
const QUIET = process.env.PITCH_QUIET !== "0";
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
if (FLAT) {
  K.page = "FFFFFF";
  K.mark = "F4F4F4";
}

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
    options: { name: "title", type: "title", x: M, y: 0.80, w: CW, h: 1.00, fontFace: F, fontSize: 32, bold: true,
      color: dark ? C.background1 : C.text1, valign: "top", align: "left", margin: 0 },
    text: "",
  },
});
const eyebrowPh = (dark) => ({
  placeholder: {
    options: { name: "eyebrow", type: "body", x: M, y: 0.50, w: CW, h: 0.26, fontFace: F, fontSize: 12, bold: !QUIET,
      color: dark ? (QUIET ? K.onDarkMute : K.skyOnDark) : (FLAT || QUIET ? K.mute : K.blue), valign: "top", margin: 0 },
    text: "",
  },
});
// BidBox 로고: 제품 저장소 logo.png에서 make_logo.py로 바탕을 투명하게 뺀 것(가로:세로 870:308)
const LOGO = (dark) => path.join(__dirname, "assets", dark ? "bidbox_logo_dark.png" : "bidbox_logo_light.png");
const LOGO_R = 870 / 308;
const footer = (dark) => [
  { image: { x: M, y: 7.03, w: 0.26 * LOGO_R, h: 0.26, path: LOGO(dark) } },
  { text: { text: "· Narani", options: { x: M + 0.26 * LOGO_R + 0.08, y: 7.02, w: 2, h: 0.3, fontFace: F, fontSize: 11,
    color: dark ? K.onDarkMute : K.mute, margin: 0 } } },
];
pres.defineSlideMaster({ title: "LIGHT", background: { color: K.page },
  objects: [eyebrowPh(false), titlePh(false), ...footer(false)],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 11, color: K.mute, align: "right" } });
pres.defineSlideMaster({ title: "DARK", background: { color: K.navy },
  objects: [eyebrowPh(true), titlePh(true), ...footer(true)],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 11, color: K.onDarkMute, align: "right" } });
pres.defineSlideMaster({ title: "COVER", background: { color: K.navy }, objects: [] });

// ---------- 도우미 ----------
let objN = 0;
const nm = (p) => `${p}-${++objN}`;
function T(s, text, o) {
  s.addText(text, Object.assign({ fontFace: F, margin: 0, isTextBox: true, valign: "top", objectName: nm("text") }, o));
}
function box(s, o) {
  if (QUIET) {
    if ((o.line === K.blue || o.line === K.blueDeep) && o.fill !== K.blue && o.fill !== K.blueDeep) {
      o = Object.assign({}, o, { line: K.line });
    }
    if (o.fill === K.blueTint && o.w >= 4 && o.h >= 2) o = Object.assign({}, o, { fill: K.white });
  }
  const shape = o.radius === 0 ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE;
  const opt = { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill || K.white },
    line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
    objectName: nm(o.name || "card") };
  if (shape === pres.shapes.ROUNDED_RECTANGLE) opt.rectRadius = o.radius == null ? 0.14 : o.radius;
  if (o.shadow) opt.shadow = { type: "outer", color: "000000", opacity: 0.12, blur: 8, offset: 2, angle: 90 };
  s.addShape(shape, opt);
}
// v12. 카드보다 먼저 깔아, 카드와 화면이 판 가장자리에 겹치게 한다.
function plate(s, o = {}) {
  box(s, { x: o.x == null ? M : o.x, y: o.y == null ? 1.88 : o.y, w: o.w == null ? CW : o.w, h: o.h == null ? 4.55 : o.h,
    fill: o.fill || K.white, line: o.line, radius: o.radius == null ? 0.18 : o.radius, shadow: o.shadow !== false, name: "plate" });
}
function chip(s, text, x, y, color, tint, o = {}) {
  if (QUIET) {
    const solid = tint === K.blue || tint === K.blueDeep;
    const wash = tint === K.blueTint || tint === K.skyOnDark;
    const textBrand = color === K.blue || color === K.blueDeep || color === K.skyOnDark || color === K.blue2 || color === K.blue3;
    if (solid) { color = K.ink; tint = K.slateTint; }
    else if (wash || textBrand) {
      if (color === K.skyOnDark) color = K.onDarkMute;
      else if (textBrand || color === K.white) color = K.ink2;
      if (wash) tint = K.slateTint;
    }
  }
  const w = o.w || Math.max(0.9, text.length * 0.135 + 0.35);
  s.addText(text, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.16, x, y, w, h: o.h || 0.32,
    fill: { color: tint }, line: { type: "none" }, fontFace: F, fontSize: o.size || 12, bold: true, color,
    align: "center", valign: "middle", margin: 0, objectName: nm("chip") });
  return w;
}
const iconCache = {};
async function icon(name, color, px = 256) {
  if (QUIET && (color === K.blue || color === K.blueDeep || color === K.blue2 || color === K.blue3 || color === K.skyOnDark)) {
    color = color === K.skyOnDark ? K.onDarkMute : K.ink2;
  }
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
  if (QUIET) {
    const brand = (c) => c === K.blue || c === K.blueDeep || c === K.blue2 || c === K.blue3 || c === K.blueTint || c === K.skyOnDark;
    if (brand(fg) || brand(bg)) {
      const onDark = bg === K.navy || bg === K.navy2 || bg === K.navy3 || fg === K.skyOnDark;
      if (onDark && (bg === K.navy || bg === K.navy2 || bg === K.navy3)) fg = K.onDarkMute;
      else if (onDark) { fg = K.onDarkMute; bg = K.navy3; }
      else { fg = K.ink2; bg = K.slateTint; }
    }
  }
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: nm("icon-bg") });
  const p = d * 0.27;
  s.addImage({ data: await icon(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: nm("icon") });
}
// 시안(2026-11-10, PITCH_BG=1): 어두운 장 다섯 장(표지·근거 상담·숫자·자금·마무리)에만 코드로 그린 추상 무늬 바탕.
// 무늬는 한 언어(브랜드 파랑 계열의 옅은 점·선)로 두고 단락 뜻만 모양으로 바꾼다.
// flow: 흩어진 점이 한 줄로 모임(공고 → 규칙 → 금액). net: 점과 선의 연결망(수집·검색·근거). rings: 바깥으로 퍼지는 동심원(확장).
const BG = process.env.PITCH_BG === "1";
function seeded(seed) {
  let v = seed >>> 0;
  return () => { v = (Math.imul(v, 1664525) + 1013904223) >>> 0; return v / 4294967296; };
}
async function bgImage(kind, seed = 7) {
  const PW = 2560, PH = 1440, r = seeded(seed);
  const sky = "#6CB4FF", blue = "#0073E6";
  let g = "";
  if (kind === "flow") {
    // 오른쪽 아래에 흩어진 점, 그 점들에서 출발한 곡선이 오른쪽 한 점(공고 → 금액)으로 모인다.
    const ex = PW - 190, ey = 1060;
    for (let i = 0; i < 26; i++) {
      const sx = 1180 + r() * 900, sy = 260 + r() * 1100;
      const c1x = sx + 260 + r() * 160, c1y = sy, c2x = ex - 360, c2y = ey + (r() - 0.5) * 120;
      g += `<path d="M${sx.toFixed(0)},${sy.toFixed(0)} C${c1x.toFixed(0)},${c1y.toFixed(0)} ${c2x.toFixed(0)},${c2y.toFixed(0)} ${ex},${ey}" fill="none" stroke="${sky}" stroke-opacity="${(0.07 + r() * 0.08).toFixed(3)}" stroke-width="2"/>`;
      g += `<circle cx="${sx.toFixed(0)}" cy="${sy.toFixed(0)}" r="${(3 + r() * 4).toFixed(1)}" fill="${sky}" fill-opacity="${(0.22 + r() * 0.2).toFixed(2)}"/>`;
    }
    for (let i = 0; i < 140; i++) {
      const x = 900 + Math.pow(r(), 0.7) * 1650, y = r() * PH;
      g += `<circle cx="${x.toFixed(0)}" cy="${y.toFixed(0)}" r="${(1.5 + r() * 2).toFixed(1)}" fill="${sky}" fill-opacity="${(0.06 + r() * 0.12).toFixed(2)}"/>`;
    }
    // 모이는 점은 장 안쪽에 둔다(가장자리에서 잘린 반원이 되지 않게).
    g += `<circle cx="${ex}" cy="${ey}" r="46" fill="${blue}" fill-opacity="0.14"/><circle cx="${ex}" cy="${ey}" r="10" fill="${sky}" fill-opacity="0.6"/>`;
  } else if (kind === "net") {
    // 가장자리 쪽에 더 많은 점과 가까운 점끼리 잇는 선. 가운데(내용 자리)는 옅게 둔다.
    const pts = [];
    for (let i = 0; i < 90; i++) {
      let x = r() * PW, y = r() * PH;
      const cx = Math.abs(x - PW / 2) / (PW / 2), cy = Math.abs(y - PH / 2) / (PH / 2);
      if (Math.max(cx, cy) < 0.55 && r() < 0.7) { x = r() < 0.5 ? r() * PW * 0.12 : PW - r() * PW * 0.12; }
      pts.push([x, y]);
    }
    pts.forEach(([x, y], i) => {
      const near = pts.map(([a, b], j) => [Math.hypot(a - x, b - y), j]).filter(([d, j]) => j !== i).sort((a, b) => a[0] - b[0]).slice(0, 2);
      for (const [d, j] of near) if (d < 360) g += `<line x1="${x.toFixed(0)}" y1="${y.toFixed(0)}" x2="${pts[j][0].toFixed(0)}" y2="${pts[j][1].toFixed(0)}" stroke="${sky}" stroke-opacity="0.08" stroke-width="1.6"/>`;
    });
    pts.forEach(([x, y], i) => {
      const hot = i % 11 === 0;
      if (hot) g += `<circle cx="${x.toFixed(0)}" cy="${y.toFixed(0)}" r="16" fill="${blue}" fill-opacity="0.12"/>`;
      g += `<circle cx="${x.toFixed(0)}" cy="${y.toFixed(0)}" r="${hot ? 5 : 3}" fill="${hot ? blue : sky}" fill-opacity="${hot ? 0.55 : 0.22}"/>`;
    });
  } else if (kind === "rings") {
    // 오른쪽 아래에서 바깥으로 퍼지는 동심원과 원 위의 점(용역 → 물품 → 공사).
    const cx = PW - 120, cy = PH + 80;
    for (let k = 1; k <= 14; k++) {
      const rad = 150 * k, op = Math.max(0.035, 0.16 - k * 0.009);
      g += `<circle cx="${cx}" cy="${cy}" r="${rad}" fill="none" stroke="${sky}" stroke-opacity="${op.toFixed(3)}" stroke-width="${k % 3 === 0 ? 2.4 : 1.4}"/>`;
      if (k % 3 === 0) {
        const a = Math.PI + 0.25 + r() * 1.1;
        g += `<circle cx="${(cx + rad * Math.cos(a)).toFixed(0)}" cy="${(cy + rad * Math.sin(a)).toFixed(0)}" r="6" fill="${blue}" fill-opacity="0.5"/>`;
      }
    }
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${PW}" height="${PH}"><rect width="100%" height="100%" fill="#${K.navy}"/>${g}</svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function setBg(s, kind, seed) {
  if (BG) s.background = { data: await bgImage(kind, seed) };
}
function toneArrow(color) {
  if (!QUIET) return color;
  if (color === K.skyOnDark) return K.onDarkMute;
  if (color === K.blue || color === K.blue2 || color === K.blue3 || color === K.blueDeep) return K.mute;
  return color;
}
function arrow2(s, x, y, w, color = K.blue2) {
  color = toneArrow(color);
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x, y, w, h: 0.3, fill: { color }, line: { type: "none" }, objectName: nm("arrow2") });
}
function arrow(s, x, y, w, color = K.blue2) {
  color = toneArrow(color);
  s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h: 0.3, fill: { color }, line: { type: "none" }, objectName: nm("arrow") });
}
// 어두운 틀 위의 제품 화면. 이름표는 기본으로 아래, labelTop이면 위(겹칠 때 가리지 않음).
function shot(s, file, x, y, w, ratio, label, o = {}) {
  const h = w / ratio, p = o.pad == null ? 0.08 : o.pad;
  box(s, { x: x - p, y: y - p, w: w + 2 * p, h: h + 2 * p, fill: o.frame || K.navy3,
    radius: o.radius == null ? 0.12 : o.radius, shadow: o.shadow !== false, line: o.line, lw: o.lw, name: "shot-frame" });
  s.addImage({ path: CAP(file), x, y, w, h, objectName: nm("shot") });
  if (label) {
    const lw = Math.min(o.labelW || (label.length * 0.135 + 0.42), w - 0.24);
    chip(s, label, x + 0.12, o.labelTop ? y + 0.12 : y + h - 0.42, K.white, K.navy, { size: 12, w: lw, h: 0.3 });
  }
  return h;
}
// 장 기호(P·S·T)를 오른쪽 아래에 아주 옅게 깐다(6회차 강의 슬라이드의 큰 「P」). 카드보다 먼저 넣어 뒤에 깔리게 한다.
// v19: 카드 뒤에서 반원·화살촉 조각처럼 잘려 보여 그리지 않는다. PITCH_MARK=1이면 v18까지 모습.
function watermark(s, eyebrow, o = {}) {
  if (process.env.PITCH_MARK !== "1") return;
  const L = eyebrow.trim()[0];
  // T(팀) 장은 카드 사이로 「T」 가로획만 조각처럼 비쳐 결함처럼 보여서 깔지 않는다
  if (!"PS".includes(L)) return;
  T(s, L, { x: o.x == null ? 9.0 : o.x, y: o.y == null ? 2.2 : o.y, w: o.w == null ? 4.2 : o.w, h: o.h == null ? 5.2 : o.h,
    fontSize: o.size == null ? 340 : o.size, bold: true, color: K.mark, align: "right", valign: "bottom", objectName: nm("watermark") });
}
// 제목의 [[핵심어]]만 강조색으로(6회차 강의의 2색 제목)
function richTitle(title, accent) {
  if (!title.includes("[[")) return title;
  return title.split(/\[\[|\]\]/).map((t, i) => {
    if (i % 2) return { text: t, options: { color: accent, bold: true } };
    return { text: t, options: QUIET ? { bold: true } : {} };
  }).filter((r) => r.text);
}
function lightSlide(section, eyebrow, title, mark) {
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: section });
  watermark(s, eyebrow, mark);
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
function note(s, text, y = 6.43, dark = false) {
  T(s, text, { x: M, y, w: CW, h: 0.32, fontSize: 11, color: dark ? K.onDarkMute : K.mute });
}

if (FLAT) {
  const darkFill = (hex) => {
    if (!hex || typeof hex !== "string" || hex.length < 6) return false;
    const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    return (r + g + b) / 3 < 90;
  };
  const flatten = (opt, wasRound) => {
    if (!opt) return opt;
    const o = Object.assign({}, opt);
    delete o.rectRadius;
    const hadShadow = !!o.shadow;
    delete o.shadow;
    if (wasRound || hadShadow) {
      const noLine = !o.line || o.line.type === "none";
      if (noLine) {
        const fill = o.fill && o.fill.color;
        o.line = { color: darkFill(fill) ? "334155" : "E0E0E0", width: 1 };
      }
    }
    return o;
  };
  const origAdd = pres.addSlide.bind(pres);
  pres.addSlide = (opts) => {
    const s = origAdd(opts);
    const origShape = s.addShape.bind(s);
    s.addShape = (shape, opt) => {
      const wasRound = shape === pres.shapes.ROUNDED_RECTANGLE;
      const sh = wasRound ? pres.shapes.RECTANGLE : shape;
      return origShape(sh, flatten(opt, wasRound));
    };
    const origText = s.addText.bind(s);
    s.addText = (text, opt) => {
      if (opt && opt.shape === pres.shapes.ROUNDED_RECTANGLE) {
        opt = flatten(Object.assign({}, opt, { shape: pres.shapes.RECTANGLE }), true);
      }
      return origText(text, opt);
    };
    return s;
  };
}

if (QUIET) {
  const brandText = new Set([K.blue, K.blueDeep, K.blue2, K.blue3, K.skyOnDark]);
  const plainOf = (text) => {
    if (typeof text === "string") return text;
    if (Array.isArray(text)) return text.map((t) => (typeof t === "string" ? t : (t && t.text) || "")).join("");
    return "";
  };
  const heroNumber = (text, size) => {
    const s = plainOf(text).replace(/\s+/g, "");
    if (!size || size < 26) return false;
    if (/^0\d$/.test(s)) return false;
    return /\d/.test(s) && s.length <= 18;
  };
  // 우수사례는 제목·카드 제목이 굵고 본문은 보통이다. 16pt 미만과 워터마크는 보통으로 둔다.
  const keepWeight = (text, size, color) => {
    if (color === K.mark) return false;
    if (size >= 16) return true;
    return heroNumber(text, size);
  };
  const quietColor = (color, size, text) => {
    if (!brandText.has(color)) return color;
    if (heroNumber(text, size)) return color;
    if (color === K.skyOnDark) return K.onDarkMute;
    if (color === K.blue2 || color === K.blue3) return K.mute;
    return K.ink2;
  };
  const quietText = (text, opt) => {
    if (!opt) return [text, opt];
    const o = Object.assign({}, opt);
    const size = o.fontSize || 0;
    const whole = plainOf(text);
    const hero = heroNumber(whole, size);
    if (o.placeholder === "title") {
      if (Array.isArray(text)) {
        text = text.map((r) => {
          if (!r || typeof r === "string") return { text: r || "", options: { bold: true } };
          return { text: r.text, options: Object.assign({}, r.options, { bold: true }) };
        });
      }
      o.bold = true;
      return [text, o];
    }
    if (o.bold && !keepWeight(whole, size, o.color)) o.bold = false;
    if (o.color) o.color = quietColor(o.color, size, whole);
    if (o.fill && (o.fill.color === K.blue || o.fill.color === K.blueDeep) && /^[A-Z]$/.test(whole.trim())) {
      o.fill = Object.assign({}, o.fill, { color: K.slateTint });
      o.color = K.ink;
    }
    if (o.fill && o.fill.color === K.blueTint) o.fill = Object.assign({}, o.fill, { color: K.slateTint });
    if (o.line && (o.line.color === K.blue || o.line.color === K.blueDeep) && !(o.fill && (o.fill.color === K.blue || o.fill.color === K.blueDeep))) {
      o.line = Object.assign({}, o.line, { color: K.line });
    }
    if (Array.isArray(text)) {
      text = text.map((r) => {
        if (!r || typeof r === "string" || !r.options) return r;
        const ro = Object.assign({}, r.options);
        const rs = ro.fontSize || size;
        if (ro.bold && !keepWeight(r.text || "", rs, ro.color) && !hero) ro.bold = false;
        if (ro.color) ro.color = quietColor(ro.color, rs, r.text || "");
        return { text: r.text, options: ro };
      });
    }
    return [text, o];
  };
  const quietShape = (opt) => {
    if (!opt || !opt.line) return opt;
    const c = opt.line.color;
    if (c !== K.blue && c !== K.blue2 && c !== K.blue3 && c !== K.blueDeep) return opt;
    const fillC = opt.fill && opt.fill.color;
    if (fillC === c) return opt;
    return Object.assign({}, opt, { line: Object.assign({}, opt.line, { color: K.mute }) });
  };
  const quietRows = (rows) => rows.map((row) => row.map((cell) => {
    if (!cell || typeof cell !== "object") return cell;
    const c = Object.assign({}, cell);
    const dull = (ro, sample) => {
      if (!ro) return ro;
      const o = Object.assign({}, ro);
      if (o.bold && !heroNumber(sample || "", o.fontSize || 0)) o.bold = false;
      return o;
    };
    if (Array.isArray(c.text)) {
      c.text = c.text.map((r) => {
        if (!r || typeof r === "string" || !r.options) return r;
        return { text: r.text, options: dull(r.options, r.text) };
      });
    }
    if (c.options) c.options = dull(c.options, typeof c.text === "string" ? c.text : plainOf(c.text));
    return c;
  }));
  const origAdd = pres.addSlide.bind(pres);
  pres.addSlide = (opts) => {
    const s = origAdd(opts);
    const origText = s.addText.bind(s);
    s.addText = (text, opt) => {
      const [t, o] = quietText(text, opt);
      return origText(t, o);
    };
    const origShape = s.addShape.bind(s);
    s.addShape = (shape, opt) => origShape(shape, quietShape(opt));
    const origTable = s.addTable.bind(s);
    s.addTable = (rows, opt) => origTable(quietRows(rows), opt);
    return s;
  };
}

(async () => {
  // ===== 1. 표지 =====
  pres.addSection({ title: "표지" });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  await setBg(s, "flow", 11);
  T(s, "2026 예비창업패키지 · 피치덱", { x: M, y: 1.0, w: 6, h: 0.4, fontSize: 14, bold: true, color: K.skyOnDark });
  T(s, "공공기관 용역 입찰\n도우미 웹 서비스", { x: M, y: 1.5, w: 6.2, h: 1.9, fontSize: 40, bold: true,
    color: K.white, lineSpacingMultiple: 1.08 });
  T(s, "공고마다 다른 점수 규칙, 찾아서 금액 범위까지", { x: M, y: 3.5, w: 6.2, h: 0.45, fontSize: 18, color: K.skyOnDark });
  s.addImage({ path: LOGO(true), x: M - 0.05, y: 4.05, w: 0.78 * LOGO_R, h: 0.78, objectName: nm("logo") });
  T(s, "(가칭)", { x: M + 0.78 * LOGO_R + 0.08, y: 4.4, w: 1.2, h: 0.35, fontSize: 14, color: K.onDarkMute });
  s.addImage({ data: await icon("FaScaleBalanced", K.greenOnDark), x: M, y: 4.97, w: 0.3, h: 0.3, objectName: nm("esg-icon") });
  T(s, "조달 정보 차이를 줄이는 도구", { x: M + 0.42, y: 4.95, w: 5.3, h: 0.36, fontSize: 16, bold: true, color: K.greenOnDark });
  let cx = M;
  for (const [ic, t] of [["FaBroom", "청소"], ["FaShieldHalved", "경비"], ["FaPeopleGroup", "파견"], ["FaHandshake", "직업소개"]]) {
    await iconCircle(s, ic, cx, 5.6, 0.5, K.skyOnDark, K.navy3);
    T(s, t, { x: cx + 0.58, y: 5.7, w: 1.0, h: 0.3, fontSize: 14, color: K.white });
    cx += t.length > 2 ? 1.75 : 1.35;
  }
  T(s, "Narani", { x: M, y: 6.55, w: 1.5, h: 0.4, fontSize: 16, bold: true, color: K.white });
  chip(s, "회사 대상 이용권형(SaaS)", M + 1.65, 6.58, K.skyOnDark, K.navy3, { w: 2.7, h: 0.34, size: 12 });
  // 공고 상세가 주인공. 대시보드는 그 위에 겹치고, 이름표는 서로 다른 자리에 둔다.
  shot(s, "detail_top.jpg", 6.85, 0.38, 5.95, 1568 / 661, "공고 상세", { labelTop: true });
  shot(s, "dashboard.jpg", 7.55, 2.42, 5.15, 1530 / 661, "대시보드");
  s.addNotes("공공기관 용역 입찰에서 공고마다 다른 점수 규칙을 찾아, 써 낼 금액의 범위까지 알려 주는 BidBox. 이 서비스를 만든 Narani 대표 [이름]입니다. 오른쪽은 지금 실제로 동작하는 시제품 화면입니다. ■ 질의응답 참고(발표 때 읽지 않음): 청소·경비·파견·직업소개 네 업종 회사가 이용권을 사서 쓰는 웹 서비스. 사회적 의미는 공고에 맞는 규칙과 가격 기준을 보여 줘 중소 용역업체의 조달 정보 차이를 줄이는 것.");

  pres.addSection({ title: "P 문제 인식" });

  // ===== P1. 고객과 고객이 겪는 일 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "고객은 입찰 담당자, [[규칙표를 직접 찾음]]");
  // v19: 오른쪽 판은 왼쪽 카드와 같은 위·아래 선(1.95~6.15)에 맞춘다.
  plate(s, { x: 5.2, y: 1.95, w: W - M - 5.2, h: 4.2, fill: K.slateTint, line: K.line, shadow: false });
  box(s, { x: M, y: 1.95, w: 4.3, h: 4.2, fill: K.bg, line: K.line, shadow: true });
  await iconCircle(s, "FaUserTie", M + 0.28, 2.15, 0.72, K.blue, K.blueTint);
  T(s, "입찰 담당자", { x: M + 1.16, y: 2.28, w: 2.9, h: 0.46, fontSize: 22, bold: true, color: K.ink });
  const ind = [["FaBroom", "청소"], ["FaShieldHalved", "경비"], ["FaPeopleGroup", "파견"], ["FaHandshake", "직업소개"]];
  for (let i = 0; i < 4; i++) {
    const x = M + 0.35 + (i % 2) * 2.0, y = 3.15 + Math.floor(i / 2) * 1.05;
    await iconCircle(s, ind[i][0], x, y, 0.62, K.white, K.blue);
    T(s, ind[i][1], { x: x + 0.74, y: y + 0.14, w: 1.15, h: 0.36, fontSize: 16, bold: true, color: K.ink });
  }
  chip(s, "인터뷰 · 지불 의사 미확인", M + 0.28, 5.55, K.amber, K.amberTint, { w: 3.55, h: 0.36 });
  const steps = [["FaFileLines", "본문에 규칙 없음"],
    ["FaCodeBranch", "기관·날짜마다 다름"],
    ["FaPaperclip", "맞는 판을 직접 고름"],
    ["FaTriangleExclamation", "틀리면 탈락"]];
  for (let i = 0; i < 4; i++) {
    const y = 2.1 + i * 0.64, red = i === 3;
    T(s, `0${i + 1}`, { x: 5.45, y, w: 0.85, h: 0.56, fontSize: 28, bold: true, color: red ? K.red : K.blue, valign: "middle" });
    await iconCircle(s, steps[i][0], 6.3, y + 0.06, 0.44, red ? K.white : K.blue, red ? K.red : K.blueTint);
    T(s, steps[i][1], { x: 6.9, y, w: 5.6, h: 0.56, fontSize: 20, bold: true, color: red ? K.red : K.ink, valign: "middle" });
  }
  // 실제 예 칸은 판 안쪽 0.2인치 여백에 맞춘다(판 오른쪽 끝과 붙지 않게).
  const exX = 5.4, exW = W - M - 0.2 - 5.4;
  box(s, { x: exX, y: 4.8, w: exW, h: 1.15, fill: K.white, line: K.line });
  T(s, "실제 예\n인천광역시", { x: exX + 0.2, y: 4.8, w: 1.5, h: 1.15, fontSize: 14, bold: true, color: K.ink, valign: "middle" });
  const docW = 2.3, docX2 = exX + exW - 0.18 - docW, docX1 = docX2 - 0.42 - docW;
  for (const [x, t] of [[docX1, "2017-04-11본"], [docX2, "2017-12-13 개정본"]]) {
    box(s, { x, y: 4.94, w: docW, h: 0.86, fill: K.bg, line: K.line, name: "doc" });
    s.addImage({ data: await icon("FaFileLines", K.mute), x: x + 0.12, y: 5.13, w: 0.36, h: 0.48, objectName: nm("doc-icon") });
    T(s, t, { x: x + 0.54, y: 5.11, w: docW - 0.62, h: 0.52, fontSize: 14, bold: true, color: K.ink, valign: "middle" });
  }
  s.addImage({ data: await icon("FaCircleQuestion", K.amber), x: docX1 + docW + 0.12, y: 5.24, w: 0.26, h: 0.26, objectName: nm("q") });
  note(s, "옛 판과 고친 판이 함께 있음 · 면허 다른 업종은 제외");
  s.addNotes("입찰 공고를 열면, 본문에는 점수 규칙이 없습니다. 규칙표는 첨부 파일에만 있고, 기관과 날짜마다 다르며, 옛 판과 고친 판이 함께 있습니다. 네 업종 중소 용역업체의 입찰 담당자가 이것을 직접 고르고, 잘못 고르면 탈락할 수 있습니다. ■ 질의응답 참고(발표 때 읽지 않음): 실제 예는 인천광역시 계약정보의 2017-04-11본과 2017-12-13 개정본. 돈을 내는 사람은 회사 대표 또는 입찰 비용 결재자. 이름이 비슷해도 허가가 다른 업종은 제외. 고객 인터뷰와 지불 의사는 아직 확인 전.");

  // ===== P2. 문제의 근거: 뉴스·통계 =====
  s = lightSlide("P 문제 인식", "P · 문제 인식", "1년 새 하한율, [[세 번 · 2%p씩]]");
  // v19: 상자 없이 큰 숫자·선·막대만. 위아래는 가로선, 좌우는 세로선으로 나눈다.
  const rule = (x, y, w, h) => s.addShape(pres.shapes.LINE, { x, y, w, h, line: { color: K.line, width: 1.25 }, objectName: nm("rule") });
  const leftW = 4.55, rightX = M + leftW + 0.4;
  T(s, "87.995%", { x: M, y: 2.0, w: leftW, h: 0.42, fontSize: 20, bold: true, color: K.mute });
  T(s, "89.995%", { x: M, y: 2.4, w: leftW, h: 0.8, fontSize: 48, bold: true, color: K.blue });
  T(s, "청소·경비 · 9년 만", { x: M, y: 3.22, w: leftW, h: 0.34, fontSize: 16, bold: true, color: K.ink });
  rule(M + leftW + 0.2, 2.0, 0, 1.6);
  s.addShape(pres.shapes.LINE, { x: rightX + 0.3, y: 2.55, w: W - M - rightX - 0.6, h: 0, line: { color: K.blue2, width: 3 }, objectName: nm("tl") });
  const evs = [["2025-07-01", "지방계약 공사"], ["2026-01-30", "국가 시설공사"], ["2026-05-26", "일반용역"]];
  const tlW = W - M - rightX;
  evs.forEach(([d, t], i) => {
    const cx2 = rightX + tlW * (i * 2 + 1) / 6, hot = i === 2;
    s.addShape(pres.shapes.OVAL, { x: cx2 - 0.13, y: 2.42, w: 0.26, h: 0.26, fill: { color: hot ? K.blue : K.amber }, line: { type: "none" }, objectName: nm("ev") });
    T(s, d, { x: cx2 - 1.15, y: 2.78, w: 2.3, h: 0.28, fontSize: 13, bold: true, color: hot ? K.blue : K.amber, align: "center" });
    T(s, t, { x: cx2 - 1.15, y: 3.08, w: 2.3, h: 0.32, fontSize: 14, bold: true, color: K.ink, align: "center" });
  });
  rule(M, 3.82, CW, 0);
  const barTop = 3.98, botLW = 6.2;
  T(s, "분야마다 다른 하한율", { x: M, y: barTop, w: botLW, h: 0.32, fontSize: 16, bold: true, color: K.ink });
  const lows = [["시설(청소·경비)", 89.995], ["소프트웨어", 87.995], ["학술연구", 82.495], ["보험", 47.994]];
  lows.forEach(([n, v], i) => {
    const y = barTop + 0.46 + i * 0.44, w = 2.9 * v / 100;
    T(s, n, { x: M, y, w: 2.2, h: 0.32, fontSize: 13, color: K.ink2, valign: "middle" });
    box(s, { x: M + 2.25, y: y + 0.05, w, h: 0.24, fill: i === 0 ? K.blue : K.blue2, radius: 0.05, name: "low-bar" });
    T(s, `${v}%`, { x: M + 2.33 + w, y, w: 0.95, h: 0.32, fontSize: 13, bold: true, color: K.ink, valign: "middle" });
  });
  const fx0 = M + botLW + 0.4;
  rule(fx0 - 0.2, barTop, 0, 2.2);
  const facts = [["142.1조원", "중소기업 수주 63.1%", "조달청 2025-06-27", K.blue],
    ["1.5%", "사람 자문 성공보수", "비드큐", K.red]];
  facts.forEach(([v, l, src, c], i) => {
    const y = barTop + i * 1.15;
    T(s, v, { x: fx0 + 0.1, y, w: W - M - fx0 - 0.1, h: 0.6, fontSize: 36, bold: true, color: c });
    T(s, `${l}  ·  ${src}`, { x: fx0 + 0.12, y: y + 0.62, w: W - M - fx0 - 0.12, h: 0.3, fontSize: 14, color: K.ink2 });
  });
  rule(fx0 + 0.1, barTop + 1.06, W - M - fx0 - 0.1, 0);
  note(s, "출처: 조달경제신문(2026-04-17·2026-05-23), 조달청 공고 제2026-260호, 조달청 보도자료(2025-06-27) · 이 장의 숫자는 외부 자료이며 BidBox 측정값 아님", 6.43);
  s.addNotes("게다가 규칙은 계속 바뀝니다. 낙찰하한율이 1년 새 세 번, 2%포인트씩 올랐고, 청소·경비는 9년 만에 87.995%에서 89.995%가 됐습니다. ■ 질의응답 참고(발표 때 읽지 않음): 2025-07-01 지방계약 공사, 2026-01-30 국가계약 시설공사, 2026-05-26 공고부터 일반용역. 같은 고시 안에서도 분야마다 하한율이 다름. 중소기업 수주 142.1조원(63.1%, 조달청 2025-06-27). 사람 자문 성공보수 1.5%(비드큐).");

  pres.addSection({ title: "S 실현 가능성" });

  // ===== 9. 불편 → 해결 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "불편은 [[한 화면에서]]");
  T(s, "불편", { x: M, y: 1.95, w: 2.4, h: 0.3, fontSize: 14, color: K.mute });
  T(s, "해결", { x: M + 3.05, y: 1.95, w: 2.2, h: 0.3, fontSize: 14, bold: true, color: K.blue });
  const ps = [["어느 판인지 모름", "FaCodeBranch", "날짜에 맞는 판"],
    ["금액을 직접 계산", "FaCalculator", "최저가 · 최상가"],
    ["지난 낙찰을 따로", "FaChartLine", "AI 예측가"],
    ["공고를 매일 고름", "FaMagnifyingGlass", "네 업종 기본"]];
  for (let i = 0; i < 4; i++) {
    const y = 2.34 + i * 0.98;
    box(s, { x: M, y, w: 2.55, h: 0.84, fill: K.amberTint, line: "FDE68A", radius: 0.12, shadow: true });
    T(s, ps[i][0], { x: M + 0.1, y, w: 2.35, h: 0.84, fontSize: 15, bold: true, color: K.ink, align: "center", valign: "middle" });
    arrow(s, M + 2.64, y + 0.27, 0.34, K.blue);
    box(s, { x: M + 3.08, y, w: 2.5, h: 0.84, fill: K.white, line: K.line, radius: 0.12, shadow: true });
    await iconCircle(s, ps[i][1], M + 3.2, y + 0.17, 0.5, K.white, K.blue);
    T(s, ps[i][2], { x: M + 3.76, y, w: 1.72, h: 0.84, fontSize: 14, bold: true, color: K.ink, valign: "middle" });
  }
  const solveW = 6.05, solveX = W - M - solveW, imgY = 2.22, imgH = solveW / (1568 / 661);
  box(s, { x: solveX - 0.1, y: 2.08, w: solveW + 0.2, h: 4.28, fill: K.navy, radius: 0.16, shadow: true, name: "solve-stage" });
  s.addImage({ path: CAP("detail_top.jpg"), x: solveX, y: imgY, w: solveW, h: imgH, objectName: nm("solve-shot") });
  T(s, "한 화면", { x: solveX + 0.24, y: imgY + imgH + 0.18, w: 2.4, h: 0.36, fontSize: 16, bold: true, color: K.white });
  T(s, "최저 · 예측 · 최상", { x: solveX + 0.24, y: imgY + imgH + 0.56, w: solveW - 0.48, h: 0.46, fontSize: 22, bold: true, color: K.white });
  chip(s, "낙찰 보증 없음", solveX + solveW - 2.2, imgY + imgH + 0.16, K.white, K.red, { w: 2.0, h: 0.36, size: 13 });
  s.addNotes("그래서 저희는 이 불편을 한 화면에 모았습니다. 공고 날짜에 맞는 규칙표를 골라 적용하고, 써 낼 수 있는 최저가와 최상가, AI 예측가를 함께 보여 줍니다. 낙찰을 보증하지는 않고, 세 금액은 참고 정보입니다. ■ 질의응답 참고(발표 때 읽지 않음): 최저가·최상가는 거꾸로 계산하고 회사 신용 점수와 평판 점수를 반영. AI 예측가 옆에 예상 낙찰률. 네 업종이 목록 기본값. 공고 505,271건 색인.");

  // ===== 10. 제품 화면 (어두운 띠) =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "[[이미 동작하는]] 시제품 1식");
  // 공고 상세가 가장 크고, 목록·대시보드는 오른쪽 모서리에 겹친다.
  const heroW = 8.55, heroX = M, heroY = 1.92;
  const heroH = shot(s, "detail_top.jpg", heroX, heroY, heroW, 1568 / 661, "공고 상세", { labelTop: true, labelW: 1.45 });
  const sideW = 3.7, sideX = 8.85;
  shot(s, "bids_list.jpg", sideX, heroY, sideW, 1530 / 661, "네 업종 공고 탐색", { labelTop: true, labelW: 2.2 });
  shot(s, "dashboard.jpg", sideX, heroY + 1.85, sideW, 1530 / 661, "대시보드", { labelW: 1.35 });
  const flow = ["공고 모으기", "네 업종 검색", "맞는 규칙 찾기", "회사 점수 반영", "세 가지 금액", "근거 상담(RAG)", "기록으로 보강"];
  const flowY = 5.68, flowW = 1.46, flowGap = (CW - 7 * flowW) / 6;
  flow.forEach((t, i) => {
    const x = M + i * (flowW + flowGap);
    s.addText(t, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.16, x, y: flowY, w: flowW, h: 0.48,
      fill: { color: i === 2 ? K.blue : K.navy3 }, line: { type: "none" }, fontFace: F, fontSize: 12, bold: true,
      color: K.white, align: "center", valign: "middle", margin: 0, objectName: nm("flow") });
    if (i < 6) arrow(s, x + flowW + (flowGap - 0.16) / 2, flowY + 0.1, 0.16, K.skyOnDark);
  });
  note(s, "실제 프로토타입 화면 (2026-10-22 캡처) · 규칙을 확인하지 못하면 직접 입력", 6.43, true);
  s.addNotes("지금 보시는 화면은 모두 실제로 동작하는 시제품입니다. ■ 질의응답 참고(발표 때 읽지 않음): 공고 모으기, 네 업종 검색, 맞는 규칙 찾기, 회사 점수 반영, 세 가지 금액, 근거 상담(RAG)까지 한 흐름.");

  // ===== 가져온 공고가 화면까지 (수집 · 다시 배움) =====
  // 근거: 제품 저장소 docker-compose.yml, src/tasks/scheduled_tasks.py, retrain_task.py (2026-11-06).
  // 매일 02:00 수집은 기본 실행에서 켜짐. 주간 재학습은 개발 실행·용역만, 승격은 수동. 운영 재학습은 꺼짐.
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "가져온 공고가 [[화면까지]]");
  chip(s, "매일 새벽 2시 · 켜 둠", M, 1.95, K.blue, K.blueTint, { w: 2.7, size: 13 });
  const daily = [
    ["FaCloudArrowDown", "나라장터에서 가져옴", "새 공고와 낙찰 결과"],
    ["FaMagnifyingGlass", "검색 목록에 더함", "방금 가져온 것만"],
    ["FaBookOpen", "상담 자료에 더함", "최근 하루치만"],
    ["FaDisplay", "화면에서 바로 씀", "규칙, 금액, 근거 상담"],
  ];
  const cardGap = 0.3;
  const cardW = (CW - 3 * cardGap) / 4;
  for (let i = 0; i < daily.length; i++) {
    const x = M + i * (cardW + cardGap);
    box(s, { x, y: 2.42, w: cardW, h: 1.7, fill: K.bg, line: K.line, shadow: true });
    await iconCircle(s, daily[i][0], x + 0.16, 2.54, 0.44, K.white, K.blue);
    T(s, daily[i][1], { x: x + 0.16, y: 3.08, w: cardW - 0.32, h: 0.52, fontSize: 16, bold: true, color: K.ink });
    T(s, daily[i][2], { x: x + 0.16, y: 3.62, w: cardW - 0.32, h: 0.36, fontSize: 14, color: K.ink2 });
    if (i < daily.length - 1) arrow(s, x + cardW + 0.02, 3.05, cardGap - 0.04, K.blue2);
  }
  chip(s, "다시 배움 · 개발에서만 켜 둠", M, 4.28, K.amber, K.amberTint, { w: 3.5, size: 13 });
  const learn = [
    ["FaCalendarDays", "매주 월요일 새벽 3시", "용역 모델만"],
    ["FaArrowsRotate", "새 낙찰로 다시 배움", "계산식은 같음"],
    ["FaUserCheck", "더 맞으면 교체를 권함", "바꾸는 것은 사람이 함"],
  ];
  const learnGap = 0.3;
  const learnW = (CW - 2 * learnGap) / 3;
  for (let i = 0; i < learn.length; i++) {
    const x = M + i * (learnW + learnGap);
    box(s, { x, y: 4.74, w: learnW, h: 1.36, fill: K.amberTint, line: "FDE68A", shadow: true });
    await iconCircle(s, learn[i][0], x + 0.16, 4.86, 0.42, K.white, K.amber);
    T(s, learn[i][1], { x: x + 0.68, y: 4.84, w: learnW - 0.84, h: 0.46, fontSize: 16, bold: true, color: K.ink, valign: "middle" });
    T(s, learn[i][2], { x: x + 0.18, y: 5.36, w: learnW - 0.36, h: 0.58, fontSize: 14, color: K.ink2 });
    if (i < learn.length - 1) arrow(s, x + learnW + 0.02, 5.22, learnGap - 0.04, "FCD34D");
  }
  note(s, "놓친 날은 다시 켤 때 한 번 · 운영은 배움 꺼 둠, 알림만");
  s.addNotes("자료는 나라장터에서 매일 새벽 2시에 자동으로 가져오고, 새로 들어온 것만 더합니다. 예측 모델은 개발 환경에서 매주 다시 배우고, 모델 교체는 사람이 결정합니다. ■ 질의응답 참고(발표 때 읽지 않음): 505,271건 전체를 다시 만들지 않음. 수집을 놓친 날은 다시 켤 때 한 번 채움. 다시 배움은 개발 실행에서 매주 월요일 새벽 3시, 용역 모델만. 배울 때와 보여 줄 때 같은 계산. 운영 실행은 다시 배우기를 꺼 두고, 기록이 달라지면 알림만.");

  // ===== AI 쉽게 보기 (규칙 · ML · DL) =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "AI는 [[세 가지]]");
  const tools = [
    ["FaBookOpen", "규칙표 · AI 아님", K.mute, K.slateTint, "규칙 사전", "206종", "없으면 직접 입력"],
    ["FaTree", "LightGBM", K.blue, K.blueTint, "예측 나무", "90.25%", "범위 안에 든 비율"],
    ["FaRobot", "RAG", K.blueDeep, K.blueTint, "근거 상담", "24/24", "없으면 거절"],
  ];
  // v19: 세 카드는 같은 크기·같은 줄. 가운데 강조는 파란 테두리로만.
  const tw = (CW - 2 * 0.3) / 3;
  for (const i of [0, 1, 2]) {
    const [ic, tag, c, tint, t, num, out] = tools[i];
    const mid = i === 1;
    const x = M + i * (tw + 0.3);
    const y = 1.95;
    const w = tw;
    const h = 3.85;
    box(s, { x, y, w, h, fill: K.white, line: mid ? K.blue : K.line, lw: mid ? 1.5 : 1, shadow: true });
    await iconCircle(s, ic, x + 0.28, y + 0.22, 0.7, K.white, c);
    const chipW = Math.min(tag.length * 0.14 + 0.4, w - 1.3);
    chip(s, tag, x + w - 0.22 - chipW, y + 0.32, c, tint, { w: chipW });
    T(s, t, { x: x + 0.28, y: y + 1.08, w: w - 0.56, h: 0.4, fontSize: 18, bold: true, color: K.ink });
    T(s, num, { x: x + 0.24, y: y + 1.55, w: w - 0.48, h: 0.95, fontSize: 40, bold: true, color: i === 0 ? K.ink : K.blue });
    s.addText(out, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x + 0.22, y: y + h - 0.78, w: w - 0.44, h: 0.52,
      fill: { color: tint }, line: { type: "none" }, fontFace: F, fontSize: 14, bold: true, color: i === 0 ? K.ink2 : K.blueDeep,
      align: "center", valign: "middle", margin: 0, objectName: nm("tool-out") });
  }
  s.addText([
    { text: "규칙", options: { bold: true, color: K.ink } }, { text: "  →  ", options: { color: K.mute } },
    { text: "금액", options: { bold: true, color: K.blue } }, { text: "  →  ", options: { color: K.mute } },
    { text: "설명", options: { bold: true, color: K.blueDeep } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 6.1, w: CW, h: 0.52, fill: { color: K.slateTint }, line: { type: "none" },
    fontFace: F, fontSize: 20, align: "center", valign: "middle", margin: 0, objectName: nm("tools-sum") });
  s.addNotes("여기에 쓰는 도구는 세 가지입니다. 규칙 사전은 AI가 아니라 사람이 정리한 규칙표 206종입니다. 예측 나무는 지난 낙찰을 배운 머신러닝 LightGBM이고, 근거 상담은 근거를 먼저 찾고 그 근거로만 답하는 RAG입니다. ■ 질의응답 참고(발표 때 읽지 않음): 맞는 규칙이 없으면 직접 입력. 예측 범위 안에 실제 값이 든 비율 90.25%(3,589건, 2026-08-30). RAG는 딥러닝 언어 모델과 검색을 함께 씀. 예측 모델은 딥러닝이 아님.");

  // ===== RAG 인텔리전스 (실제 화면 포함) =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "근거 상담(RAG): [[없으면 답하지 않음]]");
  await setBg(s, "net", 23);
  // 왼쪽: AI 어드바이저 실제 화면(2026-11-05 로컬 실행 캡처, 사용자 이름이 보이는 사이드바는 잘라 냄)
  const advW = 7.35, advH = advW / (1015 / 468), advY = 1.92;
  box(s, { x: M - 0.08, y: advY - 0.08, w: advW + 0.16, h: advH + 0.16, fill: K.navy3, radius: 0.12, shadow: true, name: "adv-frame" });
  s.addImage({ path: CAP("ai_advisor_rag.jpg"), x: M, y: advY, w: advW, h: advH, objectName: nm("adv-shot") });
  const rag = [["FaComments", "질문"], ["FaListCheck", "뜻 파악"], ["FaMagnifyingGlass", "세 길로 찾기"],
    ["FaDatabase", "근거 모으기"], ["FaRobot", "근거로 답"]];
  const fx = M + advW + 0.3, fw = CW - advW - 0.3;
  const stepH = 0.52, stepGap = (advH - 5 * stepH) / 4;
  for (let i = 0; i < 5; i++) {
    const y = advY + i * (stepH + stepGap), hot = i === 4;
    box(s, { x: fx, y, w: fw, h: stepH, fill: hot ? K.blue : K.navy2, line: hot ? K.blue : K.navy3, radius: 0.1 });
    await iconCircle(s, rag[i][0], fx + 0.08, y + 0.08, 0.36, hot ? K.blue : K.skyOnDark, hot ? K.white : K.navy3);
    T(s, `${i + 1}  ${rag[i][1]}`, { x: fx + 0.52, y, w: fw - 0.62, h: stepH, fontSize: 15, bold: true, color: K.white, valign: "middle" });
  }
  // 품질 숫자. 칸 너비 약 1.9인치라 32pt 「144/144」은 넘친다. 칸이 담는 최대인 28pt.
  const rq = [["144/144", "숫자 정답"], ["1.0", "근거를 다시 찾음"], ["72/72", "출처 인용"], ["24/24", "거절해야 할 질문"], ["0건", "근거 없는 답"], ["2.66초", "가장 느린 쪽 첫 글자"]];
  const qGap = 0.14, qw = (CW - 5 * qGap) / 6;
  rq.forEach(([v, l], i) => {
    const x = M + i * (qw + qGap);
    box(s, { x, y: 5.55, w: qw, h: 0.8, fill: K.navy2, line: K.navy3, radius: 0.1 });
    T(s, v, { x: x + 0.04, y: 5.58, w: qw - 0.08, h: 0.42, fontSize: 20, bold: true, color: K.skyOnDark, align: "center" });
    T(s, l, { x: x + 0.04, y: 5.99, w: qw - 0.08, h: 0.3, fontSize: 11, color: K.onDarkMute, align: "center" });
  });
  note(s, "96요청 평가 · 첫 글자 2,664.19ms · 화면 2026-11-05 · 서비스 안 모델", 6.45, true);
  s.addNotes("근거 상담은 근거가 없으면 답하지 않습니다. 왼쪽 화면처럼, 찾은 문서에 내용이 없자 지어내지 않고 답할 수 없다고 했습니다. 96개 요청 평가에서 숫자 정답 144/144를 지켰습니다. ■ 질의응답 참고(발표 때 읽지 않음): 뜻으로·제목 글자로 찾고 숫자는 데이터베이스에서 바로 조회. 서비스 안 언어 모델이 출처를 달아 답함. 출처 인용 72/72, 거절해야 할 질문 24/24, 근거 없는 답 0건, 가장 느린 쪽 첫 글자 2,664.19ms.");

  // ===== 11. 세 가지 금액 + 확인된 규칙만 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "[[확인된 규칙만]] 계산");
  box(s, { x: M, y: 1.95, w: 5.55, h: 4.35, fill: K.white, line: K.line });
  T(s, "세 가지 금액", { x: M + 0.28, y: 2.12, w: 4.9, h: 0.4, fontSize: 18, bold: true, color: K.ink });
  const bx0 = M + 0.4, bw = 4.7;
  box(s, { x: bx0, y: 3.72, w: bw, h: 0.28, fill: K.line, radius: 0.14, name: "range-track" });
  box(s, { x: bx0 + bw * 0.12, y: 3.72, w: bw * 0.76, h: 0.28, fill: K.blue2, radius: 0.14, name: "range-band" });
  const mk = [[0.12, "최저가", "써도 되는\n가장 낮은 금액", K.blueDeep], [0.5, "AI 예측가", "AI가 예상한\n금액", K.blue], [0.88, "최상가", "써도 되는\n가장 높은 금액", K.blueDeep]];
  for (const [p, n, d, c] of mk) {
    const x = bx0 + bw * p;
    s.addShape(pres.shapes.OVAL, { x: x - 0.18, y: 3.68, w: 0.36, h: 0.36, fill: { color: c }, line: { color: K.white, width: 2 }, objectName: nm("marker") });
    T(s, n, { x: x - 0.78, y: 3.05, w: 1.56, h: 0.46, fontSize: 16, bold: true, color: c, align: "center" });
    T(s, d, { x: x - 0.82, y: 4.16, w: 1.64, h: 0.7, fontSize: 13, color: K.ink2, align: "center" });
  }
  await iconCircle(s, "FaBuilding", M + 0.35, 5.15, 0.48, K.blue, K.blueTint);
  T(s, "회사 점수 있으면 최저·최상", { x: M + 0.98, y: 5.2, w: 4.3, h: 0.4, fontSize: 16, bold: true, color: K.ink, valign: "middle" });
  chip(s, "낙찰 보증 없음 · 참고 정보", M + 0.28, 5.78, K.red, K.redTint, { size: 12, w: 2.6 });
  const rr = 880 / 450, rw = 6.05, rx = W - M - rw;
  box(s, { x: rx - 0.08, y: 1.87, w: rw + 0.16, h: rw / rr + 0.16, fill: K.navy3, radius: 0.12, shadow: true, name: "rule-frame" });
  s.addImage({ path: CAP("detail_rule.jpg"), x: rx, y: 1.95, w: rw, h: rw / rr, objectName: nm("rule-shot") });
  const ruleBottom = 1.95 + rw / rr;
  const redY = ruleBottom + 0.38, redH = 1.95 + 4.35 - redY;
  box(s, { x: rx - 0.08, y: redY, w: rw + 0.16, h: redH, fill: K.white, line: K.line, radius: 0.12, shadow: true });
  T(s, "금액을 비우고 직접 입력", { x: rx + 0.14, y: redY + 0.12, w: rw - 0.28, h: 0.34, fontSize: 16, bold: true, color: K.red });
  T(s, "0원 추천 없음   ·   비예가 계산 안 함   ·   협상은 분포만", { x: rx + 0.14, y: redY + 0.5, w: rw - 0.28, h: 0.34, fontSize: 14, color: K.ink2 });
  s.addNotes("금액도 같은 원칙입니다. 확인된 규칙이 있을 때만 계산하고, 없으면 금액을 비워 직접 입력하게 합니다. ■ 질의응답 참고(발표 때 읽지 않음): 금액이 없는 공고에 0원 추천 없음. 비예가는 가격점수 계산 안 함. 협상에 의한 계약은 평가비율과 낙찰률 참고 분포만.");

  // ===== 12. 경쟁 비교 OX =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "비교는 [[OX로]]");
  plate(s, { x: M - 0.08, y: 1.86, w: CW + 0.16, h: 4.38, fill: K.white, line: K.line });
  const sym = { "○": K.green, "△": K.amber, "✕": K.red, "—": "94A3B8" };
  const cell = (v, bid) => {
    const m = v[0], rest = v.slice(1).trim();
    const runs = [{ text: m, options: { color: sym[m] || K.ink, bold: true, fontSize: 20 } }];
    if (rest) runs.push({ text: "  " + rest, options: { color: bid ? K.ink : K.ink2, bold: !!bid, fontSize: 13 } });
    return { text: runs, options: { fill: { color: bid ? K.blueTint : K.white } } };
  };
  const hdr = ["항목", "나라장터", "계산기", "사람 자문", "BidBox"].map((h, i) => ({ text: h,
    options: { bold: true, fontSize: 14, color: i === 4 ? K.white : K.ink, fill: { color: i === 4 ? K.blue : K.slateTint }, align: i ? "center" : "left" } }));
  const rowsD = [["규칙표 고르기", "—", "✕", "—", "○ 자동 · 206종"],
    ["기준 없을 때", "—", "✕", "—", "○ 직접 입력"],
    ["세 가지 금액", "△", "△", "○", "○ 한 화면"],
    ["근거 상담", "—", "✕", "—", "○ 출처"],
    ["옛 판·지금 판", "—", "△", "—", "○ 날짜 기준"],
    ["금액 비공개", "—", "—", "—", "○ 0원 없음"],
    ["요금", "—", "△ 구독", "✕ 1.5%", "○ 1,000원"]];
  const rows = [hdr].concat(rowsD.map((r) => r.map((v, i) => i === 0
    ? { text: v, options: { bold: true, fontSize: 14, color: K.ink, fill: { color: K.bg } } }
    : Object.assign(cell(v, i === 4), {}))));
  rows.slice(1).forEach((r) => r.slice(1).forEach((c) => (c.options.align = "center")));
  s.addTable(rows, { x: M, y: 1.95, w: CW, colW: [2.35, 2.15, 2.45, 2.15, 3.033], rowH: 0.52, fontFace: F, valign: "middle",
    border: { type: "solid", pt: 0.75, color: K.line }, margin: [0.04, 0.06, 0.04, 0.08], objectName: nm("compare") });
  note(s, "○ 지원  △ 부분 지원  ✕ 미지원  — 미확인(근거 미확보, 숫자를 만들지 않음) · 성공보수 1.5%는 비드큐 1:1 분석 안내(bidq.co.kr)", 6.36);
  s.addNotes("다른 방법과 비교하면, 입찰 계산기처럼 점수를 내는 도구는 이미 있습니다. 다른 점은 공고 날짜에 맞는 규칙표를 저희가 먼저 찾고, 맞는 규칙이 없으면 계산하지 않는다는 것입니다. 사람 자문은 성공보수 1.5%, 저희는 이용권 1,000원 가설입니다. ■ 질의응답 참고(발표 때 읽지 않음): 나라장터의 규칙표 고르기는 확인하지 못함. 계산기는 넣은 숫자로 점수만 내고 기준이 없어도 그대로 계산. 확인하지 못한 칸은 미확인(—).");

  // ===== 13. 개발 현황 =====
  s = darkSlide("S 실현 가능성", "S · 실현 가능성", "만든 것은 [[숫자로]]");
  await setBg(s, "net", 41);
  const tiles = [["206종", "규칙 ID", "기관·시행일별 계산 규칙"], ["21,394건", "자동 적용", "31,305건 중 (2026-10-19)"],
    ["505,271건", "공고 색인", "최근 1년 (2026-09-09)"], ["1.1885%p", "낙찰률 평균 오차", "3,000건 (2026-09-15)"],
    ["90.25%", "예측 범위 포함률", "3,589건 (2026-08-30)"], ["144/144", "상담 숫자 정답", "96개 요청 평가"]];
  const tileGap = 0.24, tileW = (CW - 3.45 - tileGap - 2 * tileGap) / 3;
  tiles.forEach(([v, l, d], i) => {
    const x = M + (i % 3) * (tileW + tileGap), y = 1.95 + Math.floor(i / 3) * 2.28;
    box(s, { x, y, w: tileW, h: 2.08, fill: K.navy2, line: K.navy3 });
    T(s, v, { x: x + 0.14, y: y + 0.18, w: tileW - 0.28, h: 0.64, fontSize: 28, bold: true, color: K.skyOnDark });
    T(s, l, { x: x + 0.16, y: y + 0.92, w: tileW - 0.32, h: 0.4, fontSize: 16, bold: true, color: K.white });
    T(s, d, { x: x + 0.16, y: y + 1.36, w: tileW - 0.32, h: 0.52, fontSize: 13, color: K.onDarkMute });
  });
  const panelX = M + 3 * (tileW + tileGap) - 0.16;
  box(s, { x: panelX, y: 1.95, w: 3.45, h: 4.55, fill: K.navy2, line: K.navy3 });
  T(s, "개발 진행 상태 (10개 항목)", { x: panelX + 0.2, y: 2.1, w: 3.05, h: 0.55, fontSize: 16, bold: true, color: K.white });
  const stat = [[7, "구현 완료", K.greenOnDark, K.navy3], [1, "부분 구현", K.amberOnDark, K.navy3], [2, "미구축 · 협약 중 구축", K.redOnDark, K.navy3]];
  let sq = 0;
  for (const [n, , c] of stat) for (let k = 0; k < n; k++, sq++)
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: panelX + 0.2 + sq * 0.3, y: 2.75, w: 0.24, h: 0.4, rectRadius: 0.05, fill: { color: c }, line: { type: "none" }, objectName: nm("sq") });
  stat.forEach(([n, l, c, t], i) => {
    s.addText(String(n), { shape: pres.shapes.OVAL, x: panelX + 0.2, y: 3.35 + i * 0.52, w: 0.4, h: 0.4, fill: { color: t }, line: { type: "none" },
      fontFace: F, fontSize: 14, bold: true, color: c, align: "center", valign: "middle", margin: 0, objectName: nm("count") });
    T(s, l, { x: panelX + 0.72, y: 3.35 + i * 0.52, w: 2.5, h: 0.4, fontSize: 13, color: K.onDarkMute, valign: "middle" });
  });
  T(s, "협약 중\n운영 주소 · 결제 · 약관\n보안 · 시험 기록 · 지자체", { x: panelX + 0.2, y: 4.95, w: 3.05, h: 0.78, fontSize: 14, color: K.onDarkMute });
  s.addText([{ text: "특허 1건 · 7개월차", options: { bold: true, color: K.white } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: panelX + 0.16, y: 5.78, w: 3.13, h: 0.58, fill: { color: K.navy3 }, line: { type: "none" },
      fontFace: F, fontSize: 12, valign: "middle", margin: [0.04, 0.08, 0.04, 0.1], objectName: nm("patent") });
  s.addNotes("지금까지 규칙 206종을 정리했고, 최근 12개월 용역 적격심사 31,305건 중 21,394건에 자동 적용합니다. ■ 질의응답 참고(발표 때 읽지 않음): 낙찰률 평균 오차 1.1885%p(3,000건, 2026-09-15). 예측 범위 포함률 90.25%. 상담 숫자 정답 144/144. 개발 항목 10개 중 구현 7, 부분 1, 협약 중 구축 2(운영 주소·결제·약관, 보안·시험 기록·지자체). 특허 1건 7개월차.");

  // ===== 14. 협약 8개월 일정 =====
  s = lightSlide("S 실현 가능성", "S · 실현 가능성", "8개월, [[유료 전환까지]]");
  plate(s, { y: 2.62, h: 2.02, fill: K.white, line: K.line, shadow: false });
  const half = (CW - 0.3) / 2;
  s.addText([{ text: "1단계 · 1~4개월   ", options: { bold: true, color: K.blue } }, { text: "2,000만 원", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 1.95, w: half, h: 0.58, fill: { color: K.blueTint }, line: { type: "none" },
      fontFace: F, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: nm("phase") });
  s.addText([{ text: "2단계 · 5~8개월   ", options: { bold: true, color: K.ink2 } }, { text: "2,000만 원 (중간평가 통과 시)", options: { bold: true, color: K.ink } }],
    { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M + half + 0.3, y: 1.95, w: half, h: 0.58, fill: { color: K.white }, line: { color: K.line, width: 1 },
      fontFace: F, fontSize: 15, align: "center", valign: "middle", margin: 0, objectName: nm("phase") });
  s.addShape(pres.shapes.LINE, { x: M + 0.5, y: 3.48, w: CW - 1.0, h: 0, line: { color: K.line, width: 3 }, objectName: nm("tl") });
  const ms = ["운영 공개", "회사 점수", "이용권 결제", "시험 10사", "지자체 규칙", "시험 20사", "단가 확정", "유료 전환"];
  const step = CW / 8;
  ms.forEach((t, i) => {
    const cxm = M + step * i + step / 2, hot = i === 3 || i === 7;
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: cxm - 0.28, y: 3.2, w: 0.56, h: 0.56, fill: { color: hot ? K.blue : K.white },
      line: { color: K.blue, width: 2 }, fontFace: F, fontSize: 16, bold: true, color: hot ? K.white : K.blue, align: "center", valign: "middle", margin: 0, objectName: nm("node") });
    T(s, `${i + 1}개월`, { x: cxm - 0.72, y: 2.72, w: 1.44, h: 0.32, fontSize: 13, color: K.mute, align: "center" });
    T(s, t, { x: cxm - 0.72, y: 3.86, w: 1.44, h: 0.7, fontSize: 16, bold: true, color: K.ink, align: "center" });
  });
  const goals = [["시험 사용", "10사", "4개월차"], ["시험 사용", "20사", "6개월차 (가설)"], ["유료 전환", "8사", "8개월차 (가설)"]];
  const goalGap = 0.24, goalW = (CW - 2 * goalGap) / 3;
  goals.forEach(([a, b, c], i) => {
    const x = M + i * (goalW + goalGap);
    box(s, { x, y: 4.85, w: goalW, h: 1.45, fill: K.bg, line: K.line });
    T(s, b, { x: x + 0.16, y: 4.95, w: goalW - 0.32, h: 0.72, fontSize: 36, bold: true, color: K.blue, align: "center" });
    T(s, `${a}  ·  ${c}`, { x: x + 0.16, y: 5.68, w: goalW - 0.32, h: 0.4, fontSize: 14, color: K.ink2, align: "center" });
  });
  s.addNotes("앞으로 8개월은 운영 공개와 결제, 시험 사용 10사로 시작해, 중간평가를 통과하면 20사로 넓히고 유료 전환을 확인합니다. ■ 질의응답 참고(발표 때 읽지 않음): 1개월 운영 공개(서버·도메인·보안 접속, Windows 실행 확인). 2개월 회사 점수 재현, 등급 넣은 계정과 비교, 약관 초안. 3개월 이용권 1,000원(가설), 보안 점검. 4개월 시험 사용 10사, 사용 기록, 중간평가 자료. 5개월 지자체 규칙 보강, 막힌 원문 등록, 영업 담당 채용. 6개월 시험 사용 20사(가설), 비수도권 10사 이상. 7개월 단가 확정, 상표 1건, 특허 출원 1건. 8개월 유료 전환, 유료 회사 수 기준 8사, 결과 보고.");

  pres.addSection({ title: "S 성장전략" });

  // ===== 16. 시장 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "[[7,494억원]]부터",
    { x: 11.2, y: 3.6, w: 2.2, h: 3.4, size: 180 });
  // v19: 왼쪽은 넓이가 금액에 비례하는 겹친 원(아래 맞춤), 오른쪽은 TAM·SAM·SOM 줄과 같은 출발선의 업종 막대.
  // 지름 비율 = √(금액 비율). SAM/TAM 10.2/55.1, SOM/TAM 0.7494/55.1.
  const dT = 3.6, cx0 = M + 2.05, base = 5.72;
  const circ = [["TAM", dT, "DCEBFB", K.ink2], ["SAM", dT * Math.sqrt(10.2 / 55.1), K.blue2, K.ink], ["SOM", dT * Math.sqrt(0.7494 / 55.1), K.blue, K.white]];
  circ.forEach(([k, d, fill, ink]) => {
    s.addShape(pres.shapes.OVAL, { x: cx0 - d / 2, y: base - d, w: d, h: d, fill: { color: fill }, line: { color: K.white, width: 1.5 }, objectName: nm(k) });
    if (k !== "SOM") T(s, k, { x: cx0 - 0.6, y: base - d + 0.16, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: ink, align: "center" });
  });
  const dS = circ[2][1];
  T(s, "SOM", { x: cx0 + dS / 2 + 0.08, y: base - dS / 2 - 0.15, w: 0.8, h: 0.3, fontSize: 12, bold: true, color: K.blue, valign: "middle" });
  T(s, "원 넓이는 금액 비율", { x: M, y: 5.8, w: 4.1, h: 0.28, fontSize: 11, color: K.mute, align: "center" });
  const mx = M + 4.6, mW2 = W - M - mx;
  const markets = [
    ["TAM", "전체 시장", "55.1조원", "국내 공공조달 용역 계약", "2024 · 조달청 (2025-06-27)"],
    ["SAM", "닿을 수 있는 시장", "10.2조원", "나라장터 용역 적격심사 낙찰 20,824건", "2025 · 나라장터 낙찰결과 집계"],
    ["SOM", "처음 들어갈 시장", "7,494억원+", "네 업종 면허 적격심사 낙찰 1,869건", "2025-09-15~2026-09-14 · 하한"],
  ];
  markets.forEach(([k, role, num, what, src], i) => {
    const y = 1.95 + i * 0.86, som = i === 2;
    T(s, k, { x: mx, y: y + 0.04, w: 1.55, h: 0.3, fontSize: 14, bold: true, color: som ? K.blue : K.ink2 });
    T(s, role, { x: mx, y: y + 0.36, w: 1.55, h: 0.3, fontSize: 12, color: K.mute });
    T(s, num, { x: mx + 1.6, y, w: 2.4, h: 0.66, fontSize: 30, bold: true, color: som ? K.blue : K.ink, valign: "middle" });
    T(s, what, { x: mx + 4.05, y: y + 0.04, w: mW2 - 4.05, h: 0.3, fontSize: 12, color: K.ink2 });
    T(s, src, { x: mx + 4.05, y: y + 0.36, w: mW2 - 4.05, h: 0.3, fontSize: 11, color: K.mute });
    if (i < 2) s.addShape(pres.shapes.LINE, { x: mx, y: y + 0.78, w: mW2, h: 0, line: { color: K.line, width: 1 }, objectName: nm("mkt-rule") });
  });
  T(s, "SOM 업종별 낙찰 (억원)", { x: mx, y: 4.62, w: 4.2, h: 0.3, fontSize: 14, bold: true, color: K.ink });
  const som = [["위생관리", 5316, "1,319건"], ["경비업", 3471, "651건"], ["파견", 642, "220건"]];
  const barX = mx + 1.3, barMax = mW2 - 1.3 - 1.9;
  som.forEach(([n, v, c], i) => {
    const y = 5.0 + i * 0.4, bw2 = barMax * v / 5316;
    T(s, n, { x: mx, y, w: 1.25, h: 0.32, fontSize: 13, bold: true, color: K.ink2, valign: "middle" });
    box(s, { x: barX, y: y + 0.04, w: bw2, h: 0.24, fill: i === 0 ? K.blue : K.blue3, radius: 0.05, name: "som-bar" });
    T(s, [{ text: v.toLocaleString("en-US"), options: { bold: true, color: K.ink } }, { text: "  " + c, options: { color: K.mute, fontSize: 12 } }],
      { x: barX + bw2 + 0.1, y, w: 1.8, h: 0.32, fontSize: 14, valign: "middle" });
  });
  T(s, "고객 계약액(매출 아님) · SOM은 하한(면허 56%, 업종 중복) · 직업소개 20건·39억, 낙찰 0건", { x: M, y: 6.25, w: CW, h: 0.32, fontSize: 13, color: K.mute });
  s.addNotes("시장은 공공조달 용역 55.1조원, 그중 나라장터 용역 적격심사 10.2조원이고, 처음 들어갈 네 업종은 7,494억원 이상입니다. 모두 고객이 따는 계약 금액이고 저희 매출이 아닙니다. ■ 질의응답 참고(발표 때 읽지 않음): TAM 2024년, 조달청 2025-06-27. SAM 2025년 20,824건. SOM 2025-09-15~2026-09-14, 1,869건. SOM은 하한(면허 정보가 있는 공고가 금액의 56%, 업종 중복). 위생관리 5,316억원, 경비업 3,471억원, 파견 642억원. 직업소개 20건·39억원, 적격심사 낙찰 0건.");

  // ===== 17. 비즈니스 모델 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "검색은 무료, [[판단은 이용권]]");
  const bm = [
    ["무료", "0원", "지금", "검색 · 대시보드", "FaGift", K.white, K.ink, false],
    ["① 주수익", "1,000원", "3개월차", "규칙 · 금액 · 상담", "FaTicket", K.blue, K.white, true],
    ["② 묶음", "미정", "3년차", "대행 · 교육", "FaLayerGroup", K.white, K.ink, false],
    ["③ API", "미공개", "4년차", "여러 사용자", "FaPlug", K.white, K.ink, false],
  ];
  const bmGap = 0.42, bmW = (CW - 3 * bmGap) / 4;
  for (let i = 0; i < 4; i++) {
    const [tag, hero, when, what, ic, fill, ink, main] = bm[i];
    const x = M + i * (bmW + bmGap);
    box(s, { x, y: 2.05, w: bmW, h: 3.35, fill, line: main ? K.blue : K.line, lw: main ? 1.5 : 1, shadow: true });
    await iconCircle(s, ic, x + 0.22, 2.22, 0.48, main ? K.blue : K.ink2, main ? K.white : K.slateTint);
    T(s, tag, { x: x + 0.8, y: 2.28, w: bmW - 1.0, h: 0.4, fontSize: 16, bold: true, color: ink, valign: "middle" });
    T(s, hero, { x: x + 0.16, y: 2.9, w: bmW - 0.32, h: 0.85, fontSize: main ? 32 : 28, bold: true, color: main ? K.white : K.blue, align: "center" });
    T(s, when, { x: x + 0.16, y: 3.8, w: bmW - 0.32, h: 0.4, fontSize: 16, bold: true, color: ink, align: "center" });
    T(s, what, { x: x + 0.16, y: 4.35, w: bmW - 0.32, h: 0.7, fontSize: 15, color: main ? K.white : K.ink2, align: "center" });
    if (i < 3) arrow(s, x + bmW + 0.05, 3.4, bmGap - 0.1, K.blue2);
  }
  box(s, { x: M, y: 5.58, w: 8.9, h: 0.85, fill: K.white, line: K.line, shadow: true });
  T(s, "연 80만 원", { x: M + 0.25, y: 5.66, w: 3.1, h: 0.68, fontSize: 28, bold: true, color: K.blue, valign: "middle" });
  T(s, "8사 × 100이용권 × 1,000원 · 가설", { x: M + 3.4, y: 5.7, w: 5.3, h: 0.6, fontSize: 16, color: K.ink2, valign: "middle" });
  chip(s, "성공보수 없음", 9.75, 5.74, K.white, K.red, { w: 2.95, h: 0.52, size: 16 });
  s.addNotes("검색과 대시보드는 무료이고, 규칙·금액·상담은 이용권으로 씁니다. 이용권 1개 1,000원, 유료 8사가 1년에 100개씩 사면 연 80만 원이고, 지불 의사를 보는 가설입니다. 낙찰 성공보수는 받지 않습니다. ■ 질의응답 참고(발표 때 읽지 않음): 기준값 직접 입력 계산도 무료. 결제는 협약 3개월차. 묶음 이용권은 입찰 대행·교육 기관 대상, 3년차. 여러 사용자·API는 4년차 계획, 협약 중 가격 비공개. 연 80만 원은 손익이 아님.");

  // ===== 영업·홍보 전략 (강의 38~40쪽) =====
  s = lightSlide("S 성장전략", "S · 성장전략", "무료로 써 보고, [[시험 사용]]");
  const saleW = (CW - 0.3) / 2, saleH = 3.95;
  box(s, { x: M, y: 1.95, w: saleW, h: saleH, fill: K.navy, radius: 0.14, shadow: true });
  chip(s, "01 온라인", M + 0.22, 2.1, K.white, K.blue, { w: 1.35, h: 0.34, size: 12 });
  T(s, "첫 접점", { x: M + 1.7, y: 2.08, w: saleW - 1.95, h: 0.38, fontSize: 16, bold: true, color: K.white, valign: "middle" });
  // v19: 화면이 남색 카드 오른쪽 끝을 넘지 않게 좌우 0.28인치 안에 둔다.
  const webW = saleW - 0.56, webH = webW / (1530 / 661);
  const webX = M + 0.28;
  const onl = [["FaGift", "무료 입구 · 지금", "검색·대시보드"], ["FaMagnifyingGlass", "소개 180만 · 모집 300만", "4개월차"]];
  for (let i = 0; i < 2; i++) {
    const y = 2.54 + webH + 0.08 + i * 0.36;
    await iconCircle(s, onl[i][0], M + 0.24, y, 0.32, K.blue, K.blueTint);
    T(s, onl[i][1], { x: M + 0.64, y, w: saleW - 1.85, h: 0.32, fontSize: 13, color: K.onDarkMute });
    T(s, onl[i][2], { x: M + saleW - 1.2, y, w: 1.0, h: 0.32, fontSize: 11, color: K.skyOnDark, align: "right" });
  }
  const ox = M + saleW + 0.3;
  box(s, { x: ox, y: 1.95, w: saleW, h: saleH, fill: K.white, line: K.line, shadow: true });
  chip(s, "02 오프라인", ox + 0.22, 2.1, K.white, K.blueDeep, { w: 1.5, h: 0.34, size: 12 });
  T(s, "직접 설명", { x: ox + 1.85, y: 2.08, w: saleW - 2.1, h: 0.38, fontSize: 16, bold: true, color: K.ink, valign: "middle" });
  const off = [["FaGraduationCap", "교육 · 협회 설명회", "2단계 · 계획"],
    ["FaHandshake", "대행 · 교육 기관 2곳", "4개월 · 확정 전"],
    ["FaUserTie", "영업 담당 채용", "5개월차"]];
  for (let i = 0; i < 3; i++) {
    const y = 2.62 + i * 1.02;
    box(s, { x: ox + 0.18, y, w: saleW - 0.36, h: 0.88, fill: K.bg, line: K.line, radius: 0.12 });
    await iconCircle(s, off[i][0], ox + 0.32, y + 0.16, 0.56, K.white, K.blueDeep);
    T(s, off[i][1], { x: ox + 1.02, y: y + 0.08, w: saleW - 1.4, h: 0.36, fontSize: 16, bold: true, color: K.ink });
    T(s, off[i][2], { x: ox + 1.02, y: y + 0.44, w: saleW - 1.4, h: 0.32, fontSize: 14, color: K.ink2 });
  }
  box(s, { x: webX - 0.06, y: 2.48, w: webW + 0.12, h: webH + 0.12, fill: K.navy3, radius: 0.1, shadow: true, name: "web-frame" });
  s.addImage({ path: CAP("bids_list.jpg"), x: webX, y: 2.54, w: webW, h: webH, objectName: nm("web-shot") });
  s.addText([
    { text: "무료", options: { bold: true } }, { text: "  →  ", options: { color: K.mute } },
    { text: "10사", options: { bold: true, color: K.blue } }, { text: "  →  ", options: { color: K.mute } },
    { text: "20사", options: { bold: true, color: K.blue } }, { text: "  →  ", options: { color: K.mute } },
    { text: "8사", options: { bold: true, color: K.blue } }, { text: "   가설 · 낙찰 대행은 목표 아님", options: { color: K.ink2 } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.12, x: M, y: 6.08, w: CW, h: 0.52, fill: { color: K.slateTint }, line: { type: "none" },
    fontFace: F, fontSize: 20, color: K.ink, align: "center", valign: "middle", margin: 0, objectName: nm("sales-flow") });
  s.addNotes("고객은 무료 검색으로 처음 만나고, 교육·협회 설명회와 입찰 대행·교육 기관을 거쳐 시험 사용으로 이어 갑니다. ■ 질의응답 참고(발표 때 읽지 않음): 소개 페이지 180만 원, 시험 사용 모집 300만 원, 4개월차. 설명회는 2단계 계획(광고선전비 1,201,640원). 대행·교육 기관은 4개월까지 2곳, 확정 전. 영업 담당 5개월차 채용. 확정 거래처 없음. 낙찰 대행은 목표 아님.");

  // ===== 18. 고객 검증·진입 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "지불 의사는 [[시험 사용으로]]");
  // v19: 깔때기는 왼쪽 칸(폭 6.25) 가운데에, 오른쪽 카드와 0.3인치 띄운다.
  const funColW = 7.15 - 0.3 - M;
  const fun = [["무료 입구", "지금", funColW, K.blueTint, K.ink], ["10사", "4개월차", funColW - 1.1, K.blue2, K.ink],
    ["20사", "6개월차 · 가설", funColW - 2.2, K.blue, K.white], ["8사", "8개월차 · 가설", funColW - 3.3, K.blueDeep, K.white]];
  fun.forEach(([a, b, w, f, c], i) => {
    const y = 1.95 + i * 1.08, x = M + (funColW - w) / 2;
    s.addText([{ text: a, options: { bold: true, fontSize: 26, breakLine: true } }, { text: b, options: { fontSize: 14 } }],
      { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.14, x, y, w, h: 0.96, fill: { color: f }, line: { type: "none" },
        fontFace: F, color: c, align: "center", valign: "middle", margin: 0, objectName: nm("funnel") });
  });
  box(s, { x: 7.15, y: 1.95, w: 5.55, h: 2.15, fill: K.greenTint, radius: 0.14, shadow: true });
  T(s, "지금 있는 근거", { x: 7.35, y: 2.08, w: 5.15, h: 0.32, fontSize: 14, bold: true, color: K.green });
  T(s, "505,271", { x: 7.35, y: 2.36, w: 5.15, h: 0.52, fontSize: 32, bold: true, color: K.ink });
  T(s, "공고 색인", { x: 7.35, y: 2.94, w: 5.15, h: 0.26, fontSize: 14, color: K.ink2 });
  T(s, "31,305", { x: 7.35, y: 3.22, w: 5.15, h: 0.48, fontSize: 28, bold: true, color: K.ink });
  T(s, "규칙 판정", { x: 7.35, y: 3.76, w: 5.15, h: 0.26, fontSize: 14, color: K.ink2 });
  box(s, { x: 7.15, y: 4.28, w: 5.55, h: 2.0, fill: K.amberTint, radius: 0.14, shadow: true });
  T(s, "협약 중 확보", { x: 7.35, y: 4.42, w: 5.15, h: 0.32, fontSize: 14, bold: true, color: K.amber });
  // v19: 「2곳」이 인터뷰 밑에 놓여 「인터뷰 2곳」으로 읽혀서, 줄마다 이름과 값을 붙인다(값은 장·노트에 있던 사실).
  const sec = [["거래처", "대행·교육 기관 2곳"], ["인터뷰", "확인 전"], ["지불 의사", "유료 8사 · 가설"]];
  sec.forEach(([k, v], i) => {
    const y = 4.82 + i * 0.44;
    T(s, k, { x: 7.35, y, w: 1.6, h: 0.38, fontSize: 16, bold: true, color: K.ink, valign: "middle" });
    T(s, v, { x: 8.95, y, w: 3.55, h: 0.38, fontSize: 16, bold: i === 0, color: i === 0 ? K.amber : K.ink2, valign: "middle", align: "right" });
  });
  note(s, "낙찰 대행은 목표 아님 · 없는 횟수는 가설");
  s.addNotes("아직 매출과 고객이 없습니다. 고객 인터뷰와 지불 의사도 확인 전이라, 시험 사용 10사, 20사, 유료 8사 순서로 직접 확인하겠습니다. ■ 질의응답 참고(발표 때 읽지 않음): 10사는 4개월차, 사용 기록과 이용권 구매 의사. 20사는 6개월차, 비수도권 10사 이상(가설). 유료 8사는 8개월차(가설). 지금 근거는 공고 505,271건 색인과 31,305건 규칙 판정. 입찰 횟수·검토 시간·비용은 시험 사용에서 잼.");

  // ===== 19. 5개년 로드맵 =====
  s = lightSlide("S 성장전략", "S · 성장전략", "용역 → 물품 → [[공사]]");
  // 강의 52쪽: 열이 연차, 행이 사업·기술. 범위 숫자는 연차 바로 아래.
  const rd = [["1년차", "네 업종", "0.75조", K.blueTint, K.ink],
    ["2년차", "용역 확대", "10.2조", K.blue2, K.ink],
    ["3년차", "물품 · 공사", "23.4조", K.blue, K.white],
    ["4년차", "공사 · API", "54.4조", K.blueDeep, K.white],
    ["5년차", "자체조달", "+80.0조", K.navy, K.white]];
  const tech = ["지자체 규칙", "낙찰방법 44종", "공사 규칙", "계정 · API", "공고 연동"];
  const labW = 0.72, gap = 0.08, colW = (CW - labW - 4 * gap) / 5;
  const headY = 1.98, amtY = 2.52, bizY = 3.28, techY = 4.78;
  T(s, "범위", { x: M, y: amtY, w: labW, h: 0.68, fontSize: 12, bold: true, color: K.blue, valign: "middle" });
  T(s, "사업", { x: M, y: bizY, w: labW, h: 1.38, fontSize: 12, bold: true, color: K.ink, valign: "middle" });
  T(s, "기술", { x: M, y: techY, w: labW, h: 1.22, fontSize: 12, bold: true, color: K.ink2, valign: "middle" });
  rd.forEach(([year, biz, amt, fill, ink], i) => {
    const x = M + labW + i * (colW + gap);
    const onDeep = i >= 3;
    box(s, { x, y: headY, w: colW, h: 0.46, fill, radius: 0.08, name: "year" });
    T(s, year, { x: x + 0.06, y: headY, w: colW - 0.12, h: 0.46, fontSize: 13, bold: true, color: ink, align: "center", valign: "middle" });
    T(s, amt, { x, y: amtY, w: colW, h: 0.62, fontSize: 20, bold: true, color: onDeep ? K.blue : K.ink, align: "center", valign: "middle" });
    box(s, { x, y: bizY, w: colW, h: 1.38, fill: K.white, line: K.line, radius: 0.08, name: "biz" });
    T(s, biz, { x: x + 0.08, y: bizY + 0.08, w: colW - 0.16, h: 1.22, fontSize: 16, bold: true, color: K.ink, align: "center", valign: "middle" });
    box(s, { x, y: techY, w: colW, h: 1.22, fill: K.slateTint, line: K.line, radius: 0.08, name: "tech" });
    T(s, tech[i], { x: x + 0.08, y: techY + 0.06, w: colW - 0.16, h: 1.1, fontSize: 15, color: K.ink2, align: "center", valign: "middle" });
  });
  T(s, "고객 계약액 · 4~5년차 계획 · 물품 13.2조 · 공사 31.0조 · 자체조달 80.0조(2024)", { x: M, y: 6.18, w: CW, h: 0.32, fontSize: 13, color: K.mute });
  s.addNotes("향후 5년은 네 업종에서 용역 전체, 물품, 공사, 자체조달 기관 순으로 넓힙니다. ■ 질의응답 참고(발표 때 읽지 않음): 기술은 지자체 규칙 보강, 낙찰방법 44종 정리, 물품·공사 규칙과 모델 순. 공략 범위 0.75조·10.2조·23.4조·54.4조원은 고객 계약액이고 4~5년차는 계획.");

  // ===== 20. 자금 =====
  s = darkSlide("S 성장전략", "S · 성장전략", "정부지원 [[4,000만 원]]");
  await setBg(s, "rings", 5);
  // 누적 막대도 도형으로(구글 변환 때 차트 숫자 형식이 빠짐). 단위 만원, 합계 각 2,000.
  // 밝은 칸 글자는 남색. 1E3A5F는 남색 바탕에서 사라지므로 어두운 장 토큰으로 바꾼다.
  T(s, "정부지원사업비 집행계획 (만원)", { x: M, y: 1.95, w: 7.6, h: 0.36, fontSize: 16, bold: true, color: K.white });
  const BC = { "외주용역비": K.skyOnDark, "지급수수료": K.blue3, "광고선전비": K.blue2, "인건비(영업 담당)": K.greenOnDark, "무형자산(상표·특허)": K.amberOnDark };
  const stages = [["1단계", [["외주용역비", 800, "800"], ["지급수수료", 720, "720"], ["광고선전비", 480, "480"]]],
    ["2단계", [["인건비(영업 담당)", 1000, "1,000"], ["지급수수료", 579.836, "580"], ["무형자산(상표·특허)", 300, "300"], ["광고선전비", 120.164, "120"]]]];
  const bw0 = 6.55, bx1 = M + 1.05;
  stages.forEach(([n, segs], i) => {
    const y = 2.55 + i * 1.15;
    T(s, n, { x: M, y, w: 0.95, h: 0.8, fontSize: 16, bold: true, color: K.white, valign: "middle" });
    let xx = bx1;
    for (const [k, v, lab] of segs) {
      const w = bw0 * v / 2000;
      s.addText(lab, { shape: pres.shapes.RECTANGLE, x: xx, y, w, h: 0.8, fill: { color: BC[k] }, line: { color: K.navy, width: 1.25 },
        fontFace: F, fontSize: w < 0.55 ? 11 : 14, bold: true, color: K.navy, align: "center", valign: "middle", margin: 0, objectName: nm("budget-seg") });
      xx += w;
    }
  });
  Object.entries(BC).forEach(([k, c], i) => {
    const x = M + 1.05 + (i % 3) * 2.2, y = 4.95 + Math.floor(i / 3) * 0.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.08, w: 0.22, h: 0.22, fill: { color: c }, line: { type: "none" }, objectName: nm("legend-key") });
    T(s, k, { x: x + 0.32, y, w: 1.85, h: 0.36, fontSize: 12, color: K.white, valign: "middle" });
  });
  const fund = [["정부지원 1단계", "2,000만 원", "1~4개월", K.skyOnDark], ["정부지원 2단계", "2,000만 원", "중간평가 후", K.skyOnDark],
    ["매출", "연 80만 원", "8사 × 100이용권 · 가설", K.amberOnDark], ["엔젤·시드", "미확인", "일정 없음", K.onDarkMute]];
  // v19: 이름·설명은 한 줄(작은 글자), 금액은 그 아래 한 줄. 칸 높이 안에서 줄이 겹치지 않게 띄운다.
  const fx1 = 8.5, fw1 = W - M - fx1, fh = 0.98, fg = 0.12;
  fund.forEach(([a, v, d, c], i) => {
    const y = 1.95 + i * (fh + fg);
    box(s, { x: fx1, y, w: fw1, h: fh, fill: K.navy3, radius: 0.12 });
    T(s, [{ text: a, options: { bold: true, color: K.white } }, { text: "  ·  " + d, options: { color: K.onDarkMute } }],
      { x: fx1 + 0.18, y: y + 0.1, w: fw1 - 0.36, h: 0.3, fontSize: 12, valign: "middle" });
    T(s, v, { x: fx1 + 0.18, y: y + 0.42, w: fw1 - 0.36, h: 0.5, fontSize: 26, bold: true, color: c, valign: "middle" });
  });
  T(s, "대표자 인건비 없음 · 현물 출자 없음 · 월 운영비 3,468,720원", { x: M, y: 6.28, w: CW, h: 0.3, fontSize: 13, color: K.onDarkMute });
  s.addNotes("필요한 돈은 정부지원 1단계와 2단계 각 2,000만 원이고, 2단계는 중간평가를 통과한 뒤에만 씁니다. 대표자 인건비는 넣지 않았습니다. ■ 질의응답 참고(발표 때 읽지 않음): 1단계 외주용역 800만, 지급수수료 720만, 광고선전비 480만. 2단계 영업 담당 인건비 1,000만, 지급수수료 5,798,360원, 상표·특허 300만, 광고선전비 1,201,640원. 현물 출자 없음. 엔젤·시드 일정 없음. 월 운영비 3,468,720원은 3차 회의 견적.");

  // ===== 21. 사회적 가치 =====
  s = lightSlide("S 성장전략", "S · 성장전략 · ESG", "[[같은 정보로]] 입찰");
  const esg = [["S", "FaScaleBalanced", "정보 차이", "142.1조원", "무료로 같은 정보", "비수도권 10사"],
    ["S", "FaUsers", "약자 · 고용 가점", "지금 계산", "안내 화면은 2년차", "2년차"],
    ["G", "FaUserShield", "정보 보호", "서비스 안 모델", "낙찰 비보증", "1년차부터"]];
  const ew = (CW - 2 * 0.3) / 3;
  for (let i = 0; i < 3; i++) {
    const [L, ic, t, d, g, when] = esg[i], x = M + i * (ew + 0.3);
    box(s, { x, y: 1.95, w: ew, h: 4.2, fill: K.white, line: K.line, shadow: true });
    s.addText(L, { shape: pres.shapes.OVAL, x: x + 0.25, y: 2.15, w: 0.52, h: 0.52, fill: { color: L === "G" ? K.navy3 : K.blue }, line: { type: "none" },
      fontFace: F, fontSize: 18, bold: true, color: K.white, align: "center", valign: "middle", margin: 0, objectName: nm("esg-letter") });
    s.addImage({ data: await icon(ic, K.blue2), x: x + ew - 0.78, y: 2.15, w: 0.5, h: 0.5, objectName: nm("esg-icon") });
    T(s, t, { x: x + 0.24, y: 2.85, w: ew - 0.48, h: 0.55, fontSize: 20, bold: true, color: K.ink });
    T(s, d, { x: x + 0.24, y: 3.5, w: ew - 0.48, h: 0.7, fontSize: i === 0 ? 28 : 26, bold: true, color: K.blue });
    T(s, g, { x: x + 0.24, y: 4.35, w: ew - 0.48, h: 0.7, fontSize: 16, color: K.ink2 });
    s.addText(when, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x + 0.2, y: 5.35, w: ew - 0.4, h: 0.58,
      fill: { color: K.blueTint }, line: { type: "none" }, fontFace: F, fontSize: 12, bold: true, color: K.blueDeep, align: "center", valign: "middle", margin: 0.04, objectName: nm("when") });
  }
  note(s, "지표의 숫자 목표는 시험 사용 측정 뒤 결정 · 약자기업 우대 단가는 확인 전까지 수치 미설정");
  s.addNotes("이 일의 의미는 공공조달 정보 차이를 줄이는 데 있습니다. 규칙과 가격 기준을 화면에 보여 주고, 검색과 직접 계산은 무료로 엽니다. ■ 질의응답 참고(발표 때 읽지 않음): 시험 사용 20사 중 비수도권 10사 이상. 여성·장애인기업, 청년·여성·장애인 고용 가점은 지금 계산에 반영, 안내 화면은 2년차. 필요한 항목만 수집, 상담은 서비스 안 모델, OCR 기본 꺼짐. 약관 첫머리에 낙찰 비보증. 지표 숫자 목표와 약자기업 우대 단가는 미정.");

  pres.addSection({ title: "T 팀 구성" });

  // ===== 23. 팀 =====
  s = lightSlide("T 팀 구성", "T · 팀 구성", "대표자가 [[직접 구현]]");
  box(s, { x: M, y: 1.95, w: 6.2, h: 4.35, fill: K.bg, line: K.line, shadow: true });
  await iconCircle(s, "FaUserTie", M + 0.28, 2.12, 0.78, K.white, K.blue);
  T(s, "대표자", { x: M + 1.22, y: 2.16, w: 3, h: 0.4, fontSize: 22, bold: true, color: K.ink });
  T(s, "기획 · 개발 · 데이터 검증", { x: M + 1.22, y: 2.56, w: 4.5, h: 0.32, fontSize: 14, color: K.mute });
  const rep = [["FaBriefcase", "경력", "6년 9개월"],
    ["FaRobot", "기술", "규칙 206종 · 시제품"],
    ["FaListCheck", "원칙", "확인된 규칙만"],
    ["FaDatabase", "인프라", "공고 505,271건"]];
  for (let i = 0; i < 4; i++) {
    const y = 3.15 + i * 0.72;
    await iconCircle(s, rep[i][0], M + 0.28, y, 0.5, K.blue, K.blueTint);
    T(s, rep[i][1], { x: M + 0.92, y: y + 0.08, w: 0.85, h: 0.36, fontSize: 14, bold: true, color: K.ink2 });
    T(s, rep[i][2], { x: M + 1.85, y: y + 0.04, w: 4.1, h: 0.42, fontSize: 18, bold: true, color: K.ink });
  }
  box(s, { x: 6.95, y: 1.95, w: 5.78, h: 1.95, fill: K.white, line: K.blue, lw: 1.5, dash: "dash", shadow: true });
  await iconCircle(s, "FaUserPlus", 7.28, 2.12, 0.58, K.blue, K.blueTint);
  T(s, "영업 담당 1명", { x: 8.02, y: 2.12, w: 2.3, h: 0.36, fontSize: 18, bold: true, color: K.ink });
  chip(s, "채용 예정 · 협약 5개월차", 10.35, 2.14, K.amber, K.amberTint, { w: 2.2, h: 0.32 });
  T(s, "경력 2~3년", { x: 8.02, y: 2.7, w: 4.4, h: 0.7, fontSize: 28, bold: true, color: K.ink });
  box(s, { x: 6.95, y: 4.14, w: 5.78, h: 2.16, fill: K.white, line: K.line, shadow: true });
  T(s, "협력 기관 3곳 (모두 확정 전)", { x: 7.28, y: 4.26, w: 5.2, h: 0.36, fontSize: 16, bold: true, color: K.ink });
  const prt = [["FaStamp", "변리", "상표 1건·특허 출원 1건", "7개월차"], ["FaLock", "보안 점검", "취약점 점검 1식", "3개월차"], ["FaHandshake", "입찰 대행·교육", "고객 모집 채널", "4개월까지 2곳"]];
  for (let i = 0; i < 3; i++) {
    const y = 4.72 + i * 0.5;
    await iconCircle(s, prt[i][0], 7.32, y, 0.4, K.blue, K.blueTint);
    T(s, prt[i][1], { x: 7.84, y: y + 0.04, w: 1.55, h: 0.32, fontSize: 14, bold: true, color: K.ink });
    T(s, prt[i][2], { x: 9.4, y: y + 0.04, w: 2.0, h: 0.32, fontSize: 13, color: K.ink2 });
    T(s, prt[i][3], { x: 11.35, y: y + 0.04, w: 1.2, h: 0.32, fontSize: 12, color: K.mute, align: "right" });
  }
  note(s, "매출·유료 고객·낙찰 실적 없음 · 영업은 채용으로 보완");
  s.addNotes("이 일을 할 팀입니다. 저는 전시컨벤션 사업 매니저로 6년 9개월 일하며 사업입찰 진행을 맡았고, 지금의 시제품을 직접 만들었습니다. 영업 담당 1명은 협약 5개월차에 채용합니다. ■ 질의응답 참고(발표 때 읽지 않음): 국제회의·기업행사 기획·진행. 규칙 206종 정리, 예측 모델 학습·운영, 공고 505,271건 색인. 서버는 협약 중 임대. 영업 담당 조건은 공공조달 입찰·영업 경력 2~3년, 시험 사용 모집과 입찰 실무 검수. 협력 기관 3곳(변리 7개월차, 보안 점검 3개월차, 대행·교육 4개월까지 2곳) 모두 확정 전. 매출·유료 고객·낙찰 실적 없음.");

  // ===== 24. 마무리 =====
  pres.addSection({ title: "마무리" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "마무리" });
  await setBg(s, "flow", 29);
  s.addImage({ path: LOGO(true), x: M - 0.03, y: 0.95, w: 0.5 * LOGO_R, h: 0.5, objectName: nm("logo") });
  T(s, "확인된 규칙만\n계산하는 도우미", { x: M, y: 1.7, w: 10.5, h: 1.9, fontSize: 44, bold: true, color: K.white, lineSpacingMultiple: 1.05 });
  const fin = [["206종", "공고에 맞는 계산 규칙"], ["31,305건", "최근 12개월 규칙 판정"], ["0원", "금액 없는 공고 추천"]];
  const finW = 3.86, finGap = 0.28;
  // v19: 가운데만 떠 있던 카드를 없애고, 세 숫자를 같은 줄에 윗선으로 나눈다.
  for (let i = 0; i < 3; i++) {
    const [v, l] = fin[i];
    const x = M + i * (finW + finGap);
    s.addShape(pres.shapes.LINE, { x, y: 4.05, w: finW, h: 0, line: { color: K.navy3, width: 2 }, objectName: nm("fin-rule") });
    T(s, v, { x, y: 4.2, w: finW, h: 0.7, fontSize: 36, bold: true, color: K.skyOnDark });
    T(s, l, { x, y: 4.92, w: finW, h: 0.35, fontSize: 14, color: K.onDarkMute });
  }
  T(s, "감사합니다 · 질문 받겠습니다", { x: M, y: 5.9, w: 8, h: 0.5, fontSize: 22, bold: true, color: K.white });
  T(s, "Narani · 낙찰 보증 없음, 세 가지 금액은 참고 정보", { x: M, y: 6.5, w: 8, h: 0.35, fontSize: 12, color: K.onDarkMute });
  s.addNotes("정리하면, 저희는 확인된 규칙만 계산하고 모르면 금액을 비웁니다. 네 업종에서 시작해 물품과 공사로 넓히고, 협약 8개월 동안 유료 8사라는 가설을 확인하겠습니다. 중소 용역업체도 같은 정보로 입찰하게 하는 것이 저희 목표입니다. Narani 대표 [이름]이었습니다. 감사합니다. 질문 받겠습니다.");

  await pres.writeFile({ fileName: OUT });
  if (APPLY) { const { applyTheme } = require(APPLY); await applyTheme(OUT, THEME); }
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
