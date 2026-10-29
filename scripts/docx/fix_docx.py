"""HWP 변환 docx 서식 정리: 글꼴 통일, ko-KR, exact→atLeast, 음수 자간·장평 제거, A4, compat 15"""
import sys, re, pathlib
root = pathlib.Path(sys.argv[1]) / 'word'
FONT = '맑은 고딕'
RF = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
stats = {}
def sub(pat, rep, s, key):
    s2, n = re.subn(pat, rep, s); stats[key] = stats.get(key, 0) + n; return s2
for f in root.glob('*.xml'):
    if f.name in ('fontTable.xml', 'webSettings.xml'): continue
    s = f.read_text('utf8')
    s = sub(r'<w:rFonts\b[^>]*/>', RF, s, 'rFonts')
    s = sub(r'<w:lang\b[^>]*/>', '<w:lang w:val="en-US" w:eastAsia="ko-KR" w:bidi="ar-SA"/>', s, 'lang')
    s = sub(r'w:lineRule="exact"', 'w:lineRule="atLeast"', s, 'exact→atLeast')
    s = sub(r'<w:spacing w:val="-\d+"\s*/>', '', s, '음수자간')
    s = sub(r'<w:w w:val="\d+"\s*/>', '', s, '장평')
    s = sub(r'<w:pgSz w:w="\d+" w:h="\d+"', '<w:pgSz w:w="11906" w:h="16838"', s, 'pgSz')
    s = sub(r'(w:name="compatibilityMode"[^>]*w:val=")14"', r'\g<1>15"', s, 'compat')
    f.write_text(s, 'utf8')
# fontTable: 맑은 고딕만 남김
ft = root / 'fontTable.xml'; s = ft.read_text('utf8')
fonts = re.findall(r'<w:font w:name="[^"]*".*?</w:font>|<w:font w:name="[^"]*"[^>]*/>', s, re.S)
keep = [x for x in fonts if f'w:name="{FONT}"' in x][:1]
for x in fonts:
    if x not in keep: s = s.replace(x, '', 1)
stats['fontTable 제거'] = len(fonts) - len(keep)
ft.write_text(s, 'utf8')
print(stats)

# 표 폭이 본문 폭을 넘으면 마지막 열을 그만큼 줄임 (중첩 표 없는 문서 기준)
doc = root / 'document.xml'; s = doc.read_text('utf8')
m = re.search(r'<w:pgMar [^>]*w:right="(\d+)"[^>]*w:left="(\d+)"', s)
body = 11906 - int(m.group(1)) - int(m.group(2))
def shrink(tm):
    t = tm.group(0)
    w = re.search(r'<w:tblW w:w="(\d+)" w:type="dxa"', t)
    if not w or int(w.group(1)) <= body: return t
    ex = int(w.group(1)) - body
    t = t.replace(w.group(0), f'<w:tblW w:w="{body}" w:type="dxa"', 1)
    cols = list(re.finditer(r'<w:gridCol w:w="(\d+)"', t)); c = cols[-1]
    t = t[:c.start()] + f'<w:gridCol w:w="{int(c.group(1)) - ex}"' + t[c.end():]
    def row(rm):
        r = rm.group(0); cw = list(re.finditer(r'<w:tcW w:w="(\d+)"', r))
        if not cw: return r
        c = cw[-1]; return r[:c.start()] + f'<w:tcW w:w="{int(c.group(1)) - ex}"' + r[c.end():]
    stats['표 폭 축소'] = stats.get('표 폭 축소', 0) + 1
    return re.sub(r'<w:tr[ >].*?</w:tr>', row, t, flags=re.S)
s = re.sub(r'<w:tbl>.*?</w:tbl>', shrink, s, flags=re.S)
doc.write_text(s, 'utf8')
print(stats)
