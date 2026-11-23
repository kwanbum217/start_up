"""계획서 docx(또는 txt)에서 AI가 쓴 글로 읽히기 쉬운 표현을 찾아 위치와 함께 보여 준다.

사용: python3 -I check_voice.py <계획서.docx|본문.txt> [--all]

분류마다 강약을 붙인다. 출력 머리줄은 `[강] 분류 N곳 — 방향` 또는 `[약] 분류 N곳 — 방향` 꼴이다.

- 강(한 번만 보여도 고칠 후보): 상투어, 번역투, 기대 맺음, 작업 메모, 같은 부정 반복, 경구형 제목
- 약(같은 문단에 다른 약 신호가 함께 있을 때만 고칠 후보): 막연한 수식, AI 문장부호,
  나열 공식, 대조 구문, 같은 끝말 연속, 같은 첫말 연속, 길이가 고른 문단 묶음

분류 설명
- 「작업 메모」는 제출본에 남으면 안 되는 작성 과정 흔적이다. 「제출 전 확인」, 「[확인 필요: …]」,
  「[TODO」, 「근거 제외」, 「제3자 비교」, 「추후 보완」, 버전 표기(vNN)를 잡는다.
  사실을 밝히는 말인 「미확인」「확인 전」은 잡지 않는다(예: 「돈을 낼 의사 미확인」).
- 「대조 구문」은 「아님」「아니라」「아닌」으로 무엇을 부정하는 문장이다. 부정하는 쪽이 읽는 사람이
  실제로 믿는 것을 바로잡을 때만 남긴다.
- 「같은 부정 반복」은 같은 부정 문구가 문서 안에서 두 번 이상 나오는 경우다. 앞 2어절이 같으면
  같은 문구로 본다. AGENTS.md가 요구한 「낙찰 보증 없음」 같은 고지는 REQUIRE_LIST에 두어 세지 않는다.
- 「경구형 제목」은 짧은 글머리 제목이 격언처럼 들리는 경우다(「…것이 성능」, 「…이 곧 …」, 「…의 언어」,
  「… 이유」). 표 안이어도 40자 이하면 본다.
- 「같은 첫말 연속」은 15자 이상 문단 3개 이상이 같은 첫 어절로 시작하는 구간이다.
- 「쉼표 나열 항목 수」는 판정이 아니라 참고값이다(괄호 안 쉼표와 숫자 안 쉼표는 뺀다).

분류는 blader/humanizer v4.0.0(MIT)의 구조 패턴을 한국어 계획서에 맞춘 것이다.

판정 도구가 아니다. 걸린 문장은 사람이 읽고, 우리 사실(숫자·날짜·출처·현장 장면)로 바꿀 수 있을 때만 고친다.
docx는 고치지 않는다. 후보 도구이므로 종료 코드는 늘 0이다.
"""
import re, sys, zipfile, collections, statistics
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

STRONG = '강'
WEAK = '약'

# 「X 아님」 꼴이 겹쳐도 고지 문구는 반복으로 세지 않는다(AGENTS.md가 요구한 문장).
REQUIRE_LIST = ('낙찰 보증 없음', '낙찰은 보증')

# (분류, 강약, 정규식, 바꾸는 방향)
WORD_RULES = [
    ('상투어', STRONG, r'혁신적|획기적|차별화된|최적의|최적화된|극대화|시너지|패러다임|게임\s?체인저|원스톱|선도(하는|적)|새로운 지평|한 단계 도약|경쟁력 강화|가치 창출|생태계 (조성|구축)|니즈를 충족|솔루션을 제공|발판|토대|첫걸음|초석|도약|기반 마련|본질|진정한|스마트한|똑똑한|손쉽게', '무엇이 몇 % 줄었는지, 어떤 기능인지로 바꾼다'),
    ('막연한 수식', WEAK, r'다양한|효율적(으로|인)|효과적(으로|인)|체계적(으로|인)|지속적(으로|인)|적극적(으로|인)|원활한|신속(하게|한)|고도화|폭넓은|풍부한', '대상·횟수·기간을 쓰거나 지운다'),
    ('번역투', STRONG, r'에 있어서?|함에 있어|를 통해서?|으로 인해|에 대한 [가-힣]+ 제공|것으로 (판단|예상|기대)됨|할 수 있을 것|가능하게 함|라고 할 수 있|하는 데 기여', '주어·동사를 바로 쓰거나 명사 종결로 줄인다'),
    ('기대 맺음', STRONG, r'기대됨|기대 효과가 있|될 것으로 보|전망됨|이바지|도모', '측정할 목표값과 측정 방법으로 바꾼다'),
    ('작업 메모', STRONG, r'제출 전 확인|\[확인 필요|\[TODO|근거에서 제외|근거 제외|제3자 비교|추후 보완|보완 예정|\bv\d{2}\b', '작성 과정 흔적을 지우고 제출본에는 사실만 남긴다'),
    ('대조 구문', WEAK, r'아님[.,]?|아니라|아닌 ', '부정하는 쪽을 지우고 바로 말한다'),
    ('AI 문장부호', WEAK, r'—|–|:\s|;\s|\*\*|“|”', '대시·콜론·세미콜론은 쉼표나 괄호로, 굵게는 표 머리로'),
    ('나열 공식', WEAK, r'첫째[,，]|둘째[,，]|셋째[,，]|뿐만 아니라|나아가|더 나아가|또한[,，]? ', '나열은 표나 글머리로, 접속어는 지운다'),
]

# 문단 사이를 보는 구조 분류. 곳 수는 잡힌 구간(문구·제목·연속 구간) 수다.
STRUCT_RULES = {
    '같은 부정 반복': (STRONG, '같은 부정 문구는 한 번만 남기고 나머지는 바로 말한다'),
    '경구형 제목': (STRONG, '제목은 격언 대신 그 칸에 담긴 내용을 가리키게 쓴다'),
    '같은 끝말 연속': (WEAK, '3문단 이상 같은 끝말이 이어지면 하나를 바꾼다'),
    '같은 첫말 연속': (WEAK, '15자 이상 3문단이 같은 말로 시작하면 주어나 시작을 바꾼다'),
    '길이가 고른 문단 묶음': (WEAK, '40자 이상 6연속이 고르면 문단 길이를 섞고 짧은 칸은 한 줄로 줄인다'),
}

ORDER = ['상투어', '막연한 수식', '번역투', '기대 맺음', '작업 메모', '대조 구문', 'AI 문장부호',
         '나열 공식', '같은 부정 반복', '경구형 제목', '같은 끝말 연속', '같은 첫말 연속',
         '길이가 고른 문단 묶음']

# 부정 문구에서 「X」를 뜰 때 쓰는 표지와, 제목으로 볼 글머리
NEG_RX = re.compile(r'아님|아니라|아닌')
TITLE_RX = re.compile(r'것이 [가-힣]+$|이 곧|가 곧|의 언어|의 열쇠|[가-힣]+ 이유$')
BULLET_RX = re.compile(r'^[◦○□▪]')
LEAD_RX = re.compile(r'^[◦○□▪\-·△※<>()\s]+')
PARTICLE_RX = re.compile(r'(이|가|은|는|을|를|의|도)$')
SENT_END_RX = re.compile(r'[.。?!]\s*')


def meta(cat):
    for name, strength, _, fix in WORD_RULES:
        if name == cat:
            return strength, fix
    strength, fix = STRUCT_RULES[cat]
    return strength, fix


def paragraphs(path):
    if path.endswith('.txt'):
        with open(path, encoding='utf-8') as f:
            return [l.strip() for l in f if l.strip()]
    root = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    out = []
    for p in root.iter(W + 'p'):
        t = ''.join(el.text or '' for el in p.iter(W + 't')).strip()
        if t:
            out.append(t)
    return out


def word_hits(paras):
    hits = collections.defaultdict(list)
    for i, p in enumerate(paras, 1):
        for cat, _, rx, _ in WORD_RULES:
            for m in re.finditer(rx, p):
                if m.group(0).startswith(':') and m.start() < 14:
                    continue  # 「○ 시장: …」 같은 항목 이름 뒤 콜론은 양식 관례라 뺀다
                if m.group(0) in '—–' and len(p) <= 10:
                    continue  # 표 빈칸 표시(「—」, 「— 미확인」)는 뺀다
                if cat == '작업 메모' and any(h[0] == i for h in hits[cat]):
                    continue  # 한 문단의 작업 메모는 한 곳으로 센다(「제3자 비교 글이라 근거 제외」)
                s = max(0, m.start() - 18)
                hits[cat].append((i, m.group(0), p[s:m.end() + 18]))
    return hits


def repeated_negations(paras):
    """같은 부정 문구(앞 2어절 기준)가 2번 이상 나온 묶음."""
    groups = collections.defaultdict(list)
    shown = {}
    for i, p in enumerate(paras, 1):
        for m in NEG_RX.finditer(p):
            head = p[:m.start()]
            start = 0
            for s in SENT_END_RX.finditer(head):
                start = s.end()
            if any(req in head[start:] for req in REQUIRE_LIST):
                continue  # 고지 문구는 반복으로 세지 않는다
            words = head.split()
            if not words:
                continue
            tail = words[-2:]
            key = (tail[0], PARTICLE_RX.sub('', tail[-1]))
            groups[key].append(i)
            shown.setdefault(key, '%s %s %s' % (key[0], key[1], m.group(0)))
    return [(key, sorted(set(ns)), shown[key]) for key, ns in groups.items() if len(set(ns)) >= 2]


def proverb_titles(paras):
    out = []
    for i, p in enumerate(paras, 1):
        if len(p) > 40 or not BULLET_RX.match(p):
            continue
        m = TITLE_RX.search(p)
        if m:
            out.append((i, p))
    return out


def same_ending_runs(paras):
    """연속 문단 3개 이상이 같은 두 글자로 끝나는 구간."""
    ends = [re.sub(r'[\s\.\)」]+$', '', p)[-2:] for p in paras]
    runs, run = [], 1
    for k in range(1, len(ends) + 1):
        if k < len(ends) and ends[k] == ends[k - 1] and len(paras[k]) > 15:
            run += 1
            continue
        if run >= 3:
            runs.append((k - run + 1, k, ends[k - 1]))
        run = 1
    return runs


def same_opening_runs(paras):
    """15자 이상 문단 3개 이상이 같은 첫 어절로 시작하는 구간."""
    firsts = []
    for p in paras:
        body = LEAD_RX.sub('', p)
        firsts.append(body.split()[0] if body.split() else '')
    runs, run = [], 1
    for k in range(1, len(firsts) + 1):
        joined = (k < len(firsts) and firsts[k] and firsts[k] == firsts[k - 1]
                  and len(paras[k]) > 15 and len(paras[k - 1]) > 15)
        if joined:
            run += 1
            continue
        if run >= 3:
            runs.append((k - run + 1, k, firsts[k - 1]))
        run = 1
    return runs


def flat_length_runs(paras):
    """40자 이상 긴 문단 6개 연속, 길이 변동계수 0.15 미만 구간."""
    runs = []
    long_idx = [i for i, p in enumerate(paras) if len(p) >= 40]
    for a in range(len(long_idx) - 5):
        win = long_idx[a:a + 6]
        if win[-1] - win[0] != 5:
            continue
        ls = [len(paras[i]) for i in win]
        if statistics.pstdev(ls) / statistics.mean(ls) < 0.15:
            runs.append((win[0] + 1, win[-1] + 1, ls))
    return runs


def comma_item_counts(paras):
    """괄호 밖 쉼표로 센 나열 항목 수(2개·3개·4개 이상) 문단 수."""
    dist = {2: 0, 3: 0, '4+': 0}
    for p in paras:
        t = re.sub(r'\([^)]*\)|（[^）]*）', '', p)
        t = re.sub(r'(?<=\d),(?=\d{3})', '', t)  # 5,000,000 같은 숫자 쉼표는 뺀다
        n = t.count(',') + 1
        if n == 2:
            dist[2] += 1
        elif n == 3:
            dist[3] += 1
        elif n >= 4:
            dist['4+'] += 1
    return dist


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    path = argv[0]
    show_all = '--all' in argv
    paras = paragraphs(path)
    hits = word_hits(paras)

    negs = repeated_negations(paras)
    titles = proverb_titles(paras)
    ends = same_ending_runs(paras)
    opens = same_opening_runs(paras)
    flats = flat_length_runs(paras)

    counts = {}
    weak_paras = collections.defaultdict(set)

    print('문단 %d개, 글자 %d자' % (len(paras), sum(map(len, paras))))
    for cat in ORDER:
        strength, fix = meta(cat)
        if cat not in STRUCT_RULES:
            h = hits.get(cat, [])
            counts[cat] = len(h)
            print('\n[%s] %s %d곳 — %s' % (strength, cat, len(h), fix))
            cnt = collections.Counter(w for _, w, _ in h)
            if cnt:
                print('  ', ', '.join('%s×%d' % kv for kv in cnt.most_common(10)))
            for i, w, ctx in (h if show_all else h[:8]):
                print('   %4d문단  …%s…' % (i, ctx))
            if strength == WEAK:
                for i, _, _ in h:
                    weak_paras[i].add(cat)
        elif cat == '같은 부정 반복':
            counts[cat] = len(negs)
            print('\n[%s] %s %d곳 — %s' % (strength, cat, len(negs), fix))
            for key, ns, phrase in negs:
                print('   「%s」 %s문단 (%d회)' % (phrase, '·'.join(map(str, ns)), len(ns)))
        elif cat == '경구형 제목':
            counts[cat] = len(titles)
            print('\n[%s] %s %d곳 — %s' % (strength, cat, len(titles), fix))
            for i, t in titles:
                print('   %4d문단  %s' % (i, t))
        elif cat == '같은 끝말 연속':
            counts[cat] = len(ends)
            print('\n[%s] %s %d곳 — %s' % (strength, cat, len(ends), fix))
            for a, b, tail in ends:
                print('   %d~%d문단이 모두 「%s」로 끝남' % (a, b, tail))
                for i in range(a, b + 1):
                    weak_paras[i].add(cat)
        elif cat == '같은 첫말 연속':
            counts[cat] = len(opens)
            print('\n[%s] %s %d곳 — %s' % (strength, cat, len(opens), fix))
            for a, b, head in opens:
                print('   %d~%d문단이 모두 「%s」로 시작' % (a, b, head))
                for i in range(a, b + 1):
                    weak_paras[i].add(cat)
        else:  # 길이가 고른 문단 묶음
            counts[cat] = len(flats)
            print('\n[%s] %s %d곳 — %s' % (strength, cat, len(flats), fix))
            for a, b, ls in flats:
                print('   %d~%d문단 길이 %s' % (a, b, ls))
                for i in range(a, b + 1):
                    weak_paras[i].add(cat)

    dist = comma_item_counts(paras)
    print('\n[참고] 쉼표 나열 항목 수 — 괄호 안 제외, 2개 %d회 · 3개 %d회 · 4개 이상 %d회 (판정 아님)'
          % (dist[2], dist[3], dist['4+']))

    strong_n = sum(counts[c] for c in ORDER if meta(c)[0] == STRONG)
    weak_n = sum(counts[c] for c in ORDER if meta(c)[0] == WEAK)
    print('\n[요약] 강 합계 %d곳 · 약 합계 %d곳' % (strong_n, weak_n))
    overlap = sorted(i for i, cats in weak_paras.items() if len(cats) >= 2)
    if overlap:
        print('[약 신호 겹침] 약 신호가 2개 이상 겹친 문단 %d개: %s'
              % (len(overlap), ', '.join(map(str, overlap))))
    else:
        print('[약 신호 겹침] 약 신호가 2개 이상 겹친 문단 없음')

    print('\n걸린 곳은 후보일 뿐이다. 숫자·출처·날짜는 바꾸지 않는다.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
