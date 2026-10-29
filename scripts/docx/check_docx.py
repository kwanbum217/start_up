"""사업계획서 docx 서식 규칙 점검. 사용: python3 -I check_docx.py 파일.docx (위반 시 종료코드 1)"""
import sys, zipfile, re
FONT = '맑은 고딕'
z = zipfile.ZipFile(sys.argv[1])
parts = {n: z.read(n).decode('utf8') for n in z.namelist()
         if n.startswith('word/') and n.endswith('.xml') and n.count('/') == 1}
allx = ''.join(v for k, v in parts.items() if k != 'word/fontTable.xml')
doc = parts['word/document.xml']
fails = []
def check(name, ok, detail=''):
    print(('PASS ' if ok else 'FAIL ') + name + (f' — {detail}' if detail else ''))
    if not ok: fails.append(name)

fonts = {v for m in re.findall(r'<w:rFonts\b([^>]*)/>', allx)
         for v in re.findall(r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"', m)}
check('글꼴은 맑은 고딕 하나', fonts <= {FONT}, f'사용 글꼴: {sorted(fonts)}')
themes = re.findall(r'w:\w+[Tt]heme="[^"]+"', allx)
check('테마 글꼴 참조 없음', not themes, f'{len(themes)}곳')
langs = re.findall(r'<w:lang\b[^>]*/>', allx)
bad = [l for l in langs if 'w:eastAsia="ko-KR"' not in l]
check('w:lang eastAsia=ko-KR', not bad, f'{len(bad)}/{len(langs)}곳 위반')
n = allx.count('w:lineRule="exact"')
check('줄 간격 exact 없음', n == 0, f'{n}곳')
n = len(re.findall(r'<w:spacing w:val="-\d+"', allx))
check('음수 자간 없음', n == 0, f'{n}곳')
n = len(re.findall(r'<w:w w:val="\d+"', allx))
check('장평 조정 없음', n == 0, f'{n}곳')
pg = set(re.findall(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', doc))
check('A4 11906×16838', pg == {('11906', '16838')}, str(pg))
mar = re.search(r'<w:pgMar [^>]*w:right="(\d+)"[^>]*w:left="(\d+)"', doc)
body = 11906 - int(mar.group(1)) - int(mar.group(2))
tbls = re.findall(r'<w:tbl>.*?</w:tbl>', doc, re.S)
notfixed = sum('<w:tblLayout w:type="fixed"' not in t for t in tbls)
check('표 tblLayout fixed', notfixed == 0, f'{notfixed}/{len(tbls)}개 위반')
widths = [int(w) for w in re.findall(r'<w:tblW w:w="(\d+)" w:type="dxa"', doc)]
nondxa = len(re.findall(r'<w:tblW [^>]*w:type="(?!dxa)', doc))
over = [w for w in widths if w > body]
check(f'표 폭 dxa·본문 폭({body}) 이내', not over and not nondxa,
      f'초과 {over}, dxa 아님 {nondxa}개')

# 레이아웃 통일 (normalize_layout.py 기준)
odd = [w for w in widths if w != body]
check(f'모든 표 폭 = 본문 폭({body})', not odd, f'다른 폭 {sorted(set(odd))}')
grid_bad = [i for i, t in enumerate(tbls)
            if sum(map(int, re.findall(r'<w:gridCol w:w="(\d+)"', t))) != body]
check('열 폭(gridCol) 합 = 본문 폭', not grid_bad, f'표 번호 {grid_bad}')
nocenter = [i for i, t in enumerate(tbls) if not re.search(r'<w:tblPr>.*?<w:jc w:val="center"/>.*?</w:tblPr>', t, re.S)]
check('표 가운데 정렬', not nocenter, f'표 번호 {nocenter}')
n = doc.count('<w:tblInd ')
check('표 들여쓰기(tblInd) 없음', n == 0, f'{n}곳')
mars = {re.search(r'<w:tblCellMar>(.*?)</w:tblCellMar>', t, re.S).group(1) if '<w:tblCellMar>' in t else None for t in tbls}
check('셀 여백 통일(위아래 40, 좌우 80)', mars == {'<w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
      '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>'}, f'{len(mars)}종')
cells = [c for t in tbls for c in re.findall(r'<w:tc>.*?</w:tc>', t, re.S)]
notmid = sum('<w:vAlign w:val="center"/>' not in c for c in cells)
check('셀 세로 가운데 정렬', notmid == 0, f'{notmid}/{len(cells)}칸 위반')
cind = sum(len(re.findall(r'<w:ind\b', c)) for c in cells)
check('셀 안 문단 들여쓰기 없음', cind == 0, f'{cind}곳')
anchors = doc.count('<wp:anchor')
check('이미지 inline', anchors == 0, f'anchor {anchors}개')
emu_body = body * 635
big = [int(c) for c in re.findall(r'<wp:extent cx="(\d+)"', doc) if int(c) > emu_body]
check('이미지 폭 본문 이내', not big, f'초과 {len(big)}개')
outside = re.sub(r'<w:tbl>.*?</w:tbl>', '', doc, flags=re.S)
diff = [int(c) for c in re.findall(r'<wp:extent cx="(\d+)"', outside) if int(c) != emu_body]
check('표 밖 이미지 폭 = 본문 폭', not diff, f'다른 폭 {len(diff)}개')
spill = []
for c in cells:
    w = re.search(r'<w:tcW w:w="(\d+)"', c)
    m = re.search(r'<w:tcMar>(.*?)</w:tcMar>', c, re.S)
    pad = sum(int(x) for x in re.findall(r'<w:(?:left|right) w:w="(\d+)"', m.group(1))) if m else 160
    spill += [int(x) for x in re.findall(r'<wp:extent cx="(\d+)"', c) if w and int(x) > (int(w.group(1)) - pad) * 635]
check('표 칸 안 이미지 폭 ≤ 칸 안쪽 폭', not spill, f'넘침 {len(spill)}개')
tabs = doc.count('<w:tab/>')
spaces = len(re.findall(r'<w:t xml:space="preserve">[^<]* {3,}', doc))
print(f'INFO 탭 {tabs}개, 연속 공백 {spaces}곳 — 줄맞춤 용도인지 눈으로 확인')
print(f'\n결과: {"모두 통과" if not fails else f"{len(fails)}개 규칙 위반"}')
sys.exit(1 if fails else 0)
