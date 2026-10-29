"""docx 본문을 표 구조가 보이는 텍스트로 뽑는다.

사용: python3 -I scripts/docx/docx_to_text.py <입력.docx> <출력.txt>
표는 <표>…</표>와 | 칸 | 형식, 이미지는 [그림]으로 쓴다. 초안 작성·검증용이며 docx는 고치지 않는다.
"""
import sys,zipfile,re
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def ptext(p):
    out=[]
    for el in p.iter():
        if el.tag==W+'t': out.append(el.text or '')
        elif el.tag==W+'tab': out.append(' ')
        elif el.tag==W+'br': out.append(' / ')
        elif el.tag.endswith('}blip'): out.append('[그림]')
    return ''.join(out).strip()
src,dst=sys.argv[1],sys.argv[2]
root=ET.fromstring(zipfile.ZipFile(src).read('word/document.xml'))
body=root.find(W+'body'); lines=[]
for ch in body:
    if ch.tag==W+'p':
        t=ptext(ch)
        if t: lines.append(t)
    elif ch.tag==W+'tbl':
        lines.append('<표>')
        for tr in ch.iter(W+'tr'):
            cells=[' '.join(filter(None,(ptext(p) for p in tc.iter(W+'p')))) for tc in tr.findall(W+'tc')]
            if any(cells): lines.append('| '+' | '.join(cells)+' |')
        lines.append('</표>')
open(dst,'w',encoding='utf-8').write('\n'.join(lines)+'\n')
print(dst,len(lines))
