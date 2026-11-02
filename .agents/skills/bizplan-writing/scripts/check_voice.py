"""계획서 docx(또는 txt)에서 AI가 쓴 글로 읽히기 쉬운 표현을 찾아 위치와 함께 보여 준다.

사용: python3 -I check_voice.py <계획서.docx|본문.txt> [--all]
출력: 상투어·번역투·근거 없는 수식어 목록, 같은 문형 반복, 문단 길이가 고른 구간.
판정 도구가 아니다. 걸린 문장은 사람이 읽고, 우리 사실(숫자·날짜·출처·현장 장면)로 바꿀 수 있을 때만 고친다.
docx는 고치지 않는다.
"""
import re, sys, zipfile, collections, statistics
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# (분류, 정규식, 바꾸는 방향)
PATTERNS = [
    ('상투어', r'혁신적|획기적|차별화된|최적의|최적화된|극대화|시너지|패러다임|게임\s?체인저|원스톱|선도(하는|적)|새로운 지평|한 단계 도약|경쟁력 강화|가치 창출|생태계 (조성|구축)|니즈를 충족|솔루션을 제공', '무엇이 몇 % 줄었는지, 어떤 기능인지로 바꾼다'),
    ('막연한 수식', r'다양한|효율적(으로|인)|효과적(으로|인)|체계적(으로|인)|지속적(으로|인)|적극적(으로|인)|원활한|신속(하게|한)|고도화|폭넓은|풍부한', '대상·횟수·기간을 쓰거나 지운다'),
    ('번역투', r'에 있어서?|함에 있어|를 통해서?|으로 인해|에 대한 [가-힣]+ 제공|것으로 (판단|예상|기대)됨|할 수 있을 것|가능하게 함|라고 할 수 있|하는 데 기여', '주어·동사를 바로 쓰거나 명사 종결로 줄인다'),
    ('기대 맺음', r'기대됨|기대 효과가 있|될 것으로 보|전망됨|이바지|도모', '측정할 목표값과 측정 방법으로 바꾼다'),
    ('AI 문장부호', r'—|–|:\s|;\s|\*\*|“|”', '대시·콜론·세미콜론은 쉼표나 괄호로, 굵게는 표 머리로'),
    ('나열 공식', r'첫째[,，]|둘째[,，]|셋째[,，]|뿐만 아니라|나아가|더 나아가|또한[,，]? ', '나열은 표나 글머리로, 접속어는 지운다'),
]


def paragraphs(path):
    if path.endswith('.txt'):
        return [l.strip() for l in open(path, encoding='utf-8') if l.strip()]
    root = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    out = []
    for p in root.iter(W + 'p'):
        t = ''.join(el.text or '' for el in p.iter(W + 't')).strip()
        if t:
            out.append(t)
    return out


def main():
    path = sys.argv[1]
    show_all = '--all' in sys.argv
    paras = paragraphs(path)
    hits = collections.defaultdict(list)
    for i, p in enumerate(paras, 1):
        for cat, rx, _ in PATTERNS:
            for m in re.finditer(rx, p):
                if m.group(0).startswith(':') and m.start() < 14:
                    continue  # 「○ 시장: …」 같은 항목 이름 뒤 콜론은 양식 관례라 뺀다
                if m.group(0) in '—–' and len(p) <= 10:
                    continue  # 표 빈칸 표시(「—」, 「— 미확인」)는 뺀다
                s = max(0, m.start() - 18)
                hits[cat].append((i, m.group(0), p[s:m.end() + 18]))
    print('문단 %d개, 글자 %d자' % (len(paras), sum(map(len, paras))))
    for cat, rx, fix in PATTERNS:
        h = hits.get(cat, [])
        print('\n[%s] %d곳 — %s' % (cat, len(h), fix))
        cnt = collections.Counter(w for _, w, _ in h)
        if cnt:
            print('  ', ', '.join('%s×%d' % kv for kv in cnt.most_common(10)))
        for i, w, ctx in (h if show_all else h[:8]):
            print('   %4d문단  …%s…' % (i, ctx))
    # 같은 끝말 반복: 연속 문단 3개 이상이 같은 두 글자로 끝나는 구간
    print('\n[같은 끝말 연속] (3문단 이상)')
    ends = [re.sub(r'[\s\.\)」]+$', '', p)[-2:] for p in paras]
    run = 1
    for k in range(1, len(ends) + 1):
        if k < len(ends) and ends[k] == ends[k - 1] and len(paras[k]) > 15:
            run += 1
            continue
        if run >= 3:
            print('   %d~%d문단이 모두 「%s」로 끝남' % (k - run + 1, k, ends[k - 1]))
        run = 1
    # 문단 길이가 너무 고른 구간(긴 본문 문단 6개 연속, 길이 편차 작음)
    print('\n[길이가 고른 문단 묶음] (40자 이상 6연속, 변동계수 0.15 미만)')
    long_idx = [i for i, p in enumerate(paras) if len(p) >= 40]
    for a in range(len(long_idx) - 5):
        win = long_idx[a:a + 6]
        if win[-1] - win[0] != 5:
            continue
        ls = [len(paras[i]) for i in win]
        if statistics.pstdev(ls) / statistics.mean(ls) < 0.15:
            print('   %d~%d문단 길이 %s' % (win[0] + 1, win[-1] + 1, ls))
    print('\n걸린 곳은 후보일 뿐이다. 숫자·출처·날짜는 바꾸지 않는다.')


if __name__ == '__main__':
    main()
