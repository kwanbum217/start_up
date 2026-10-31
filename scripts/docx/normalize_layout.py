"""표·그림 레이아웃 통일: 표 폭=본문 폭·가운데 정렬, 셀 여백·세로 정렬 통일, 셀 안 HWP식 들여쓰기 제거, 그림 폭=본문 폭
사용: python3 -I normalize_layout.py <압축 푼 docx 폴더>   (lxml 필요: pip install lxml)"""
import sys, pathlib
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
CELL_MAR = {'top': 40, 'left': 80, 'bottom': 40, 'right': 80}
MIN_PAD_CELL = 400  # 이보다 좁은 칸은 좌우 여백 0

path = pathlib.Path(sys.argv[1]) / 'word' / 'document.xml'
tree = etree.parse(str(path)); body = tree.getroot().find(W + 'body')
sect = body.find(W + 'sectPr')
pg, mar = sect.find(W + 'pgSz'), sect.find(W + 'pgMar')
BODY = int(pg.get(W + 'w')) - int(mar.get(W + 'left')) - int(mar.get(W + 'right'))
stats = {'표': 0, '셀': 0, '그림': 0}

def child(parent, tag, before=()):
    """parent 아래 tag 요소를 찾고, 없으면 스키마 순서상 before 태그들보다 앞에 만든다."""
    el = parent.find(W + tag)
    if el is None:
        el = etree.Element(W + tag)
        for b in before:
            ref = parent.find(W + b)
            if ref is not None:
                ref.addprevious(el); break
        else:
            parent.append(el)
    return el

def scale(widths, total):
    raw = [w * total / sum(widths) for w in widths]
    out = [int(r) for r in raw]
    for i in sorted(range(len(raw)), key=lambda i: raw[i] - out[i], reverse=True)[:total - sum(out)]:
        out[i] += 1
    return out

TBLPR_AFTER_JC = ('tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook')
for tbl in body.iter(W + 'tbl'):
    pr = tbl.find(W + 'tblPr')
    tw = child(pr, 'tblW', ('jc',) + TBLPR_AFTER_JC)
    tw.set(W + 'w', str(BODY)); tw.set(W + 'type', 'dxa')
    child(pr, 'jc', TBLPR_AFTER_JC).set(W + 'val', 'center')
    ind = pr.find(W + 'tblInd')
    if ind is not None: pr.remove(ind)
    cm = child(pr, 'tblCellMar', ('tblLook', 'tblCaption', 'tblDescription'))
    for k in list(cm): cm.remove(k)
    for side, v in CELL_MAR.items():
        e = etree.SubElement(cm, W + side); e.set(W + 'w', str(v)); e.set(W + 'type', 'dxa')
    grid = tbl.find(W + 'tblGrid'); cols = grid.findall(W + 'gridCol')
    new = scale([int(c.get(W + 'w')) for c in cols], BODY)
    for c, w in zip(cols, new): c.set(W + 'w', str(w))
    for tr in tbl.findall(W + 'tr'):
        gb = tr.find(f'{W}trPr/{W}gridBefore')
        pos = int(gb.get(W + 'val')) if gb is not None else 0
        for tc in tr.findall(W + 'tc'):
            tcpr = tc.find(W + 'tcPr')
            if tcpr is None:
                tcpr = etree.Element(W + 'tcPr'); tc.insert(0, tcpr)
            gs = tcpr.find(W + 'gridSpan'); span = int(gs.get(W + 'val')) if gs is not None else 1
            width = sum(new[pos:pos + span]); pos += span
            tcw = child(tcpr, 'tcW', ('gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'))
            tcw.set(W + 'w', str(width)); tcw.set(W + 'type', 'dxa')
            m = tcpr.find(W + 'tcMar')
            if m is not None: tcpr.remove(m)
            if width < MIN_PAD_CELL:
                m = child(tcpr, 'tcMar', ('textDirection', 'tcFitText', 'vAlign', 'hideMark'))
                for side in ('left', 'right'):
                    e = etree.SubElement(m, W + side); e.set(W + 'w', '0'); e.set(W + 'type', 'dxa')
            child(tcpr, 'vAlign', ('hideMark',)).set(W + 'val', 'center')
            for p in tc.findall(W + 'p'):
                ppr = p.find(W + 'pPr')
                if ppr is None: continue
                i = ppr.find(W + 'ind')
                if i is not None:
                    left = i.get(W + 'left')
                    if not (left and left == i.get(W + 'hanging')):  # 1-1절 글머리 내어쓰기는 남김
                        ppr.remove(i)
                sp = ppr.find(W + 'spacing')
                if sp is not None and sp.get(W + 'lineRule') in ('atLeast', 'exact'):
                    sp.set(W + 'line', '240'); sp.set(W + 'lineRule', 'auto')
            stats['셀'] += 1
    stats['표'] += 1

def cell_inner(tc):
    """셀 안쪽 폭(dxa) = tcW - 좌우 여백(셀별 tcMar 우선, 없으면 표 기본값)"""
    tcpr = tc.find(W + 'tcPr'); width = int(tcpr.find(W + 'tcW').get(W + 'w'))
    m = tcpr.find(W + 'tcMar'); pad = 0
    for side in ('left', 'right'):
        e = m.find(W + side) if m is not None else None
        pad += int(e.get(W + 'w')) if e is not None else CELL_MAR[side]
    return width - pad

for inline in body.iter(WP + 'inline'):
    ext = inline.find(WP + 'extent')
    cx, cy = int(ext.get('cx')), int(ext.get('cy'))
    tc = next(inline.iterancestors(W + 'tc'), None)
    target = (BODY if tc is None else cell_inner(tc)) * 635  # 표 칸 안 그림은 칸 안쪽 폭에 맞춤
    if cx == target: continue
    ncy = round(cy * target / cx)
    ext.set('cx', str(target)); ext.set('cy', str(ncy))
    for a in inline.iter(A + 'ext'):
        if a.get('cx') == str(cx): a.set('cx', str(target)); a.set('cy', str(ncy))
    p = next(inline.iterancestors(W + 'p'))
    ppr = p.find(W + 'pPr')
    if ppr is None:
        ppr = etree.Element(W + 'pPr'); p.insert(0, ppr)
    i = ppr.find(W + 'ind')
    if i is not None: ppr.remove(i)
    child(ppr, 'jc', ('textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')).set(W + 'val', 'center')
    stats['그림'] += 1

tree.write(str(path), xml_declaration=True, encoding='UTF-8', standalone=True)
print(f'본문 폭 {BODY} dxa', stats)
