"""계획서 docx에서 어렵게 읽히는 문장의 후보를 위치와 함께 보여 준다 (references/plain.md 2절).

사용: python3 -I check_plain.py <계획서.docx> [--all]
뽑는 것: 긴 문장(괄호 안을 뺀 글자 수가 본문 60자, 표 칸 45자 초과), 장(□)마다 풀이 없이 처음 나온 용어,
약어, 「~의」 세 번 연쇄, 조사로 끝난 문장. 후보일 뿐이고 docx는 고치지 않는다.
"""
import collections, re, sys, zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# plain.md 3절 바꿈표의 용어. 장마다 처음 나올 때 괄호 풀이(「쉬운 말(용어)」 또는 「용어(쉬운 말)」)가 있어야 한다.
TERMS = ['적격심사', '낙찰률', '투찰', '예정가격', '기초금액', '별표', '세부기준', '배점표', '기준값', '산식',
         '판본', '개정본', '규칙 판정', '최저가', '최상가', 'AI 예측가', '신인도', '평균 절대오차', '%p',
         '면허', '파일럿', '토큰', 'SaaS', 'B2B', 'API', 'OCR', 'TAM', 'SAM', 'SOM', '지자체', '중간평가', '색인']
ALLOWED_ABBR = {'AI', 'BidBox', 'Narani', 'CI', 'HTTPS', 'PC', 'G2B'}
PARTICLE_END = re.compile(r'(에|로|으로|와|과|를|을|의|에서|까지|하고)$')
LIMIT = {'본문': 60, '표': 45}


def paragraphs(path):
    root = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    body = root.find(W + 'body')
    out = []
    for el in body:
        where = '표' if el.tag == W + 'tbl' else '본문'
        for p in el.iter(W + 'p'):
            t = ''.join(x.text or '' for x in p.iter(W + 't')).strip()
            if t:
                out.append((where, t))
    return out


def sentences(t):
    return [s.strip() for s in re.split(r'(?<=[.。])\s+|(?<=[.。])$', t) if s.strip()]


def explained(term, text):
    e = re.escape(term)
    return re.search(r'\([^()]*' + e + r'[^()]*\)|' + e + r'\(', text) is not None


def main():
    path = sys.argv[1]
    show_all = '--all' in sys.argv
    paras = paragraphs(path)
    hits = collections.defaultdict(list)
    chapter, seen = '(머리)', set()
    for i, (where, p) in enumerate(paras, 1):
        if p.startswith('□'):
            chapter, seen = p[:20], set()
        for s in sentences(p):
            bare = re.sub(r'\([^()]*\)', '', re.sub(r'\([^()]*\)', '', s))  # 괄호 안 풀이·출처는 세지 않는다
            if len(bare) > LIMIT[where]:
                hits['긴 문장'].append((i, f'{where} {len(bare)}자', s[:50]))
            core = s.rstrip('.').rstrip()
            if PARTICLE_END.search(core) and len(core) > 8 and not core.endswith(('가설', '있음', '없음')):
                hits['조사로 끝남'].append((i, core[-6:], s[-40:]))
        for term in TERMS:
            if term in p and term not in seen:
                seen.add(term)
                if not explained(term, p):
                    hits['풀이 없는 첫 용어'].append((i, f'{chapter} · {term}', p[max(0, p.find(term) - 15):p.find(term) + 25]))
        for m in re.finditer(r'\b[A-Za-z]{2,}\b', p):
            if m.group(0) not in ALLOWED_ABBR and m.group(0).isupper():
                hits['약어'].append((i, m.group(0), p[max(0, m.start() - 15):m.end() + 15]))
        for m in re.finditer(r'(\S+의 ){2,}\S+의', p):
            hits['「~의」 연쇄'].append((i, '의×3', m.group(0)[:40]))
    print('문단 %d개' % len(paras))
    total = 0
    for cat in ('긴 문장', '풀이 없는 첫 용어', '약어', '「~의」 연쇄', '조사로 끝남'):
        h = hits.get(cat, [])
        total += len(h)
        print('\n[%s] %d곳' % (cat, len(h)))
        for i, w, ctx in (h if show_all else h[:10]):
            print('   %4d문단  %s  …%s…' % (i, w, ctx))
    print('\n후보 합계 %d곳. 후보일 뿐이다. 숫자·출처·날짜는 바꾸지 않는다(references/plain.md 4절).' % total)


if __name__ == '__main__':
    main()
