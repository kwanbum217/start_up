"""계획서에 쓰지 않기로 한 표현이 들어갔는지 본다.

사용: python3 -I scripts/docx/check_forbidden.py <파일.docx>
종료 코드 1: 금지 표현이 있을 때(위치와 문장을 보여 준다). docx는 고치지 않는다.

금지 목록의 정본은 AGENTS.md 「계획서에 쓰지 않는 대표자 정보」다. 여기 목록을 고치면
그 절도 같이 고친다. 양식 원본의 항목 이름(예: 「보유역량 (경력 및 학력 등)」)은
계획서가 아니라 양식의 글자이므로 ALLOW 줄에 걸린 것은 보고하지 않는다.
"""
import sys, zipfile, re
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# (찾을 글자, 설명). 학력은 「고등학교 졸업」처럼 내용으로 들어간 것만 잡는다.
BANNED = [
    ('고등학교', '학력(고등학교)'),
    ('졸업', '학력(졸업)'),
    ('대학교', '학력(대학교)'),
    ('대학원', '학력(대학원)'),
    ('학사', '학력(학사)'),
    ('석사', '학력(석사)'),
    ('박사', '학력(박사)'),
    ('경력·학력', '표 항목 이름에 학력'),
    ('출신 학교', '학력(출신 학교)'),
]
# 양식 원본에서 온 항목 이름이라 그대로 두는 줄
ALLOW = [r'보유\s*역량\s*\(\s*경력\s*및\s*학력']


def ptext(p):
    return ''.join(t.text or '' for t in p.iter(W + 't')).strip()


def lines_of(src):
    root = ET.fromstring(zipfile.ZipFile(src).read('word/document.xml'))
    body = root.find(W + 'body')
    for ch in body:
        if ch.tag == W + 'p':
            t = ptext(ch)
            if t:
                yield t
        elif ch.tag == W + 'tbl':
            for tr in ch.iter(W + 'tr'):
                for tc in tr.findall(W + 'tc'):
                    for p in tc.iter(W + 'p'):
                        t = ptext(p)
                        if t:
                            yield t


def main():
    src = sys.argv[1]
    hits = []
    for line in lines_of(src):
        if any(re.search(a, line) for a in ALLOW):
            continue
        for word, why in BANNED:
            if word in line:
                hits.append((why, line))
    print(f'{src}: 금지 표현 후보 {len(hits)}곳')
    for why, line in hits:
        print(f'  [{why}] {line[:90]}')
    if hits:
        print('AGENTS.md 「계획서에 쓰지 않는 대표자 정보」를 보고 뺀다.')
        sys.exit(1)
    print('금지 표현 없음')


if __name__ == '__main__':
    main()
