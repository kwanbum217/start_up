"""1-1 시안성 규칙 적용: 여백, 자간 -3, 장평 제거, 문단 종류별 크기·굵기·간격, 글머리 들여쓰기, 표 음영 정규화
사용: python3 -I apply_style.py <압축 푼 docx 폴더> [--step 0~3]   (lxml 필요: pip install lxml)
문단 종류는 첫 글자와 표 안 여부로 가른다. 그림·표 폭·글꼴 슬롯은 건드리지 않는다."""
import sys, pathlib
try:
    from lxml import etree
except ImportError:
    sys.exit('lxml이 필요하다. pip install lxml 하거나 lxml이 있는 파이썬으로 실행한다.')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 1-1절 값 (1/2pt 단위)
SZ_CHAP, SZ_TITLE, SZ_SUB, SZ_BODY, SZ_CELL = 26, 22, 22, 20, 18
SPACE = -3                    # -0.15pt
MARGIN_TOP = MARGIN_BOTTOM = 1417
MARGIN_HEADER = MARGIN_FOOTER = 850
LINE_BODY, LINE_CELL = 240, 216
SHADES = ('D9D9D9', 'F2F2F2', 'FFFFFF')

PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
       'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
       'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
       'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
       'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
       'sectPr', 'pPrChange']
RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
       'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
       'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
       'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange']

args = [a for a in sys.argv[1:] if not a.startswith('--')]
step = 0
if '--step' in sys.argv:
    step = int(sys.argv[sys.argv.index('--step') + 1])
    if not 0 <= step <= 3:
        sys.exit('--step 은 0~3 이다.')
if step >= 1: SZ_SUB = 21     # ① 소제목 10.5pt
if step >= 2: SZ_TITLE = 20   # ② 표 제목 10pt
if step >= 3: SZ_CHAP = 24    # ③ 장 제목 12pt

root = pathlib.Path(args[0]) / 'word'
BUL = ('○', '◦', '-', '·', '▪', '※')   # 표 칸 안 글머리

def child(parent, tag, order):
    """parent 아래 tag 요소를 찾고, 없으면 order 순서에 맞는 자리에 만든다."""
    el = parent.find(W + tag)
    if el is not None:
        return el
    el = etree.Element(W + tag)
    i = order.index(tag)
    for prev in reversed(order[:i]):
        ref = parent.find(W + prev)
        if ref is not None:
            ref.addnext(el)
            break
    else:
        parent.insert(0, el)
    return el

def ppr(p):
    el = p.find(W + 'pPr')
    if el is None:
        el = etree.Element(W + 'pPr'); p.insert(0, el)
    return el

def text(p):
    return ''.join(t.text or '' for t in p.iter(W + 't')).strip()

def has_image(p):
    return any(e.tag.endswith('}blip') or e.tag.endswith('}imagedata') for e in p.iter())

def style_run(r, sz=None, bold=None):
    rp = r.find(W + 'rPr')
    if rp is None:
        rp = etree.Element(W + 'rPr'); r.insert(0, rp)
    child(rp, 'spacing', RPR).set(W + 'val', str(SPACE))   # 기존 자간 교체
    w = rp.find(W + 'w')
    if w is not None:
        rp.remove(w)                                       # 장평 제거
    if sz is not None:
        for tag in ('sz', 'szCs'):
            child(rp, tag, RPR).set(W + 'val', str(sz))
    if bold:
        b = child(rp, 'b', RPR)
        b.attrib.pop(W + 'val', None)

def style_para(p, sz, bold=None, before=None, after=None, line=None, ind=None):
    pp = ppr(p)
    for r in p.iter(W + 'r'):
        style_run(r, sz, bold)
    if line:
        sp = child(pp, 'spacing', PPR)
        sp.set(W + 'line', str(line)); sp.set(W + 'lineRule', 'auto')
    if before is not None or after is not None:
        sp = child(pp, 'spacing', PPR)
        for side, v in (('before', before), ('after', after)):
            sp.attrib.pop(W + side + 'Lines', None)
            sp.attrib.pop(W + side + 'Autospacing', None)
            if v is not None:
                sp.set(W + side, str(v))
    i = pp.find(W + 'ind')
    if ind == 'remove':
        if i is not None:
            pp.remove(i)
    elif ind:
        i = child(pp, 'ind', PPR)
        for k in list(i.attrib):
            del i.attrib[k]
        i.set(W + 'left', str(ind[0])); i.set(W + 'hanging', str(ind[1]))

# 문단 종류별 값: 크기, 굵게, 앞, 뒤, 들여쓰기
KINDS = {
    '장 제목': dict(sz=SZ_CHAP, bold=True, before=240, after=120, ind='keep'),
    '표 제목': dict(sz=SZ_TITLE, bold=True, before=120, after=60, ind='keep'),
    '소제목': dict(sz=SZ_SUB, bold=True, before=120, after=60, ind=(300, 300)),
    '글머리 2단': dict(sz=SZ_BODY, before=0, after=40, ind=(600, 200)),
    '글머리 3단': dict(sz=SZ_BODY, before=0, after=40, ind=(800, 200)),
    '주석': dict(sz=SZ_CELL, before=0, after=40, ind=(200, 200)),
    '굵은 소제목': dict(sz=SZ_SUB, bold=True, before=120, after=60, ind='remove'),
    '그림 캡션': dict(sz=SZ_CELL, before=0, after=120, ind='keep'),
    '본문': dict(sz=SZ_BODY, before=0, after=40, ind='remove'),
}

def all_bold(p):
    runs = [r for r in p.iter(W + 'r') if ''.join(x.text or '' for x in r.iter(W + 't')).strip()]
    def bold(r):
        b = r.find(f'{W}rPr/{W}b')
        return b is not None and b.get(W + 'val') not in ('0', 'false')
    return bool(runs) and all(bold(r) for r in runs)

def kind_of(p, t):
    if t.startswith('□'): return '장 제목'
    if t.startswith('<'): return '표 제목'
    if t.startswith('※'): return '주석'
    if t.startswith('[그림'): return '그림 캡션'
    if t.startswith(('○', '◦')): return '소제목'
    if len(t) <= 40 and all_bold(p): return '굵은 소제목'   # 글머리 없이 굵게만 쓴 짧은 제목
    if t.startswith('-'): return '글머리 2단'
    if t.startswith('·'): return '글머리 3단'
    return '본문'

tree = etree.parse(str(root / 'document.xml'))
body = tree.getroot().find(W + 'body')
count = {}

# 여백: 위아래 1417, 머리말·꼬리말 850 (좌우는 양식 값 유지)
for mar in body.iter(W + 'pgMar'):
    mar.set(W + 'top', str(MARGIN_TOP)); mar.set(W + 'bottom', str(MARGIN_BOTTOM))
    mar.set(W + 'header', str(MARGIN_HEADER)); mar.set(W + 'footer', str(MARGIN_FOOTER))

# 본문 문단
for p in body.iter(W + 'p'):
    if next(p.iterancestors(W + 'tc'), None) is not None:
        continue
    if has_image(p):
        count['그림 문단'] = count.get('그림 문단', 0) + 1   # 그림 문단은 자간만 맞춘다
        for r in p.iter(W + 'r'):
            style_run(r)
        continue
    k = kind_of(p, text(p))
    count[k] = count.get(k, 0) + 1
    style_para(p, line=LINE_BODY, **KINDS[k])

# 표 음영: 회색은 D9D9D9·F2F2F2·FFFFFF 중 가까운 값으로 (머리 칸 판정보다 먼저)
SH03 = {s: tuple(int(s[i:i + 2], 16) for i in (0, 2, 4)) for s in SHADES}
shaded = 0
for shd in body.iter(W + 'shd'):
    fill = (shd.get(W + 'fill') or '').upper()
    if len(fill) != 6 or any(c not in '0123456789ABCDEF' for c in fill):
        continue
    rgb = tuple(int(fill[i:i + 2], 16) for i in (0, 2, 4))
    if max(rgb) - min(rgb) > 20:      # 회색이 아니면 그대로
        continue
    near = min(SHADES, key=lambda s: max(abs(rgb[i] - SH03[s][i]) for i in range(3)))
    if near != fill:
        shd.set(W + 'fill', near); shaded += 1

def is_head(tc):
    s = tc.find(f'{W}tcPr/{W}shd')
    return s is not None and (s.get(W + 'fill') or '').upper() == 'D9D9D9'

# 표 칸 문단: 머리 칸만 굵게, 글머리만 left 200·hanging 200, 줄 간격 216·앞뒤 0
for tbl in body.iter(W + 'tbl'):
    rows = tbl.findall(W + 'tr')
    for row, tr in enumerate(rows):
        tcs = tr.findall(W + 'tc')
        heads = [row == 0 and is_head(tc) for tc in tcs]
        if row == 0 and not any(heads):   # 머리 칸 음영이 없는 표는 첫 행 전체를 머리로
            heads = [True] * len(tcs)
        for tc, head in zip(tcs, heads):
            for p in tc.findall(W + 'p'):
                t = text(p)
                count['표 칸'] = count.get('표 칸', 0) + 1
                style_para(p, SZ_CELL, bold=True if head else None, before=0, after=0,
                           line=LINE_CELL, ind=(200, 200) if t.startswith(BUL) else 'remove')

tree.write(str(root / 'document.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)

# 머리말·꼬리말은 자간·장평만 맞춘다
for f in sorted(root.glob('*.xml')):
    if not (f.name.startswith('header') or f.name.startswith('footer')):
        continue
    t = etree.parse(str(f))
    for r in t.getroot().iter(W + 'r'):
        style_run(r)
    t.write(str(f), xml_declaration=True, encoding='UTF-8', standalone=True)

order = ['장 제목', '표 제목', '소제목', '굵은 소제목', '글머리 2단', '글머리 3단', '주석', '그림 캡션', '본문', '그림 문단']
print('문단 분류 — ' + ', '.join(f'{k} {count[k]}' for k in order if count.get(k)))
print(f'표 칸 {count.get("표 칸", 0)}, 음영 정규화 {shaded}개, 자간 {SPACE}, 위아래 여백 {MARGIN_TOP}',
      f'글자 크기 — 장 제목 {SZ_CHAP / 2:g}pt, 표 제목 {SZ_TITLE / 2:g}pt, 소제목 {SZ_SUB / 2:g}pt,',
      f'본문 {SZ_BODY / 2:g}pt, 표·주석·캡션 {SZ_CELL / 2:g}pt, --step {step}')
