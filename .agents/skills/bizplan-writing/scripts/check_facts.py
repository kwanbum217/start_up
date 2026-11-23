"""계획서의 이전 판과 새 판에서 「사실 토큰」(단위 붙은 숫자·날짜·문서 번호·분수)을 뽑아
사라지거나 새로 생기거나 개수가 달라진 것을 문단 번호와 문맥과 함께 보여 준다.

사용: python3 -I check_facts.py <이전.docx|이전.txt> <새.docx|새.txt> [--all]

목적: 계획서 규칙 「숫자·출처·날짜는 바꾸지 않는다」(AGENTS.md)를 문장을 고친 뒤 기계로 확인한다.
문장을 다듬다가 숫자가 조용히 빠지거나 바뀌는 일을 잡는 것이 목적이고, 숫자의 옳고 그름을 판정하지는 않는다.
차이가 나온 숫자는 반영 기록과 대조하고, AGENTS.md와 제품 저장소에 있는 근거만 쓴다.
docx는 읽기만 하고 고치지 않는다.

종료 코드: 차이 없음 0, 차이 있음 1, 인자 오류·파일 없음 2.

blader/humanizer v4.0.0(MIT) 작업 3단계(고친 초안에서 사실·숫자·날짜·인용이 빠지거나 바뀌지 않았는지 먼저 확인)를
이 프로젝트에 맞춘 것이다.
"""
import collections
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 단위 후보. 긴 단위부터 맞춘다(「조원」이 「원」보다 먼저, 「개월」이 「개」보다 먼저).
UNITS = ['조원', '억원', '만 원', '만원', '원', '건', '%p', '%', '종', '개사', '사',
         '개', '명', '개월', '년차', '년', '쪽', 'ms', '토큰', '시', '배']

# 문서 번호를 날짜보다 먼저 뽑는다. 「제2026-260호」의 「2026-26」이 날짜(YYYY-MM)로 잡히지 않게 하기 위해서다.
DOC_PATTERNS = [r'제\d{4}[-–—]\d+호']

# 긴 날짜부터 맞춘다. 「2026-06-26」이 「2026-06」과 「2026」으로 쪼개지지 않게 하기 위해서다.
DATE_PATTERNS = [r'\d{4}-\d{2}-\d{2}', r'\d{4}-\d{2}', r'\d{4}년\s?\d{1,2}월', r'\d{4}년']

# 분수·비율 꼴. 예: 144/144
FRACTION_PATTERN = r'\d+/\d+'

# 숫자(쉼표·소수점 포함) 뒤 선택 공백 뒤 단위.
UNIT_RE = re.compile(
    r'(?<![\d.,])(?P<num>\d[\d,]*(?:\.\d+)?)\s?(?P<unit>' +
    '|'.join(sorted(UNITS, key=len, reverse=True)) + r')'
)

LIMIT = 20


def paragraphs(path):
    """docx는 word/document.xml의 모든 w:p, txt는 줄 단위. 빈 문단은 뺀다.
    check_voice.py와 같은 방식이며, 그 파일을 import하지 않고 함수를 그대로 복사했다."""
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


def _mask(text, spans):
    """뽑은 자리를 글자 수가 같은 빈칸으로 가린다. 자리 번호가 밀리지 않는다."""
    if not spans:
        return text
    chars = list(text)
    for s, e in spans:
        for i in range(s, e):
            chars[i] = '\x00'
    return ''.join(chars)


def _spans(pattern, text):
    return [(m.start(), m.end()) for m in re.finditer(pattern, text)]


def extract_facts(text):
    """한 문단에서 사실 토큰을 (토큰, 시작, 끝) 목록으로 뽑는다. 문서 번호·날짜·분수를 먼저 뽑고
    그 자리를 가린 뒤 나머지 숫자를 뽑아, 날짜에 든 숫자가 단위 숫자로 다시 세어지지 않게 한다."""
    found = []
    masked = text
    for rx in DOC_PATTERNS:
        spans = _spans(rx, masked)
        for s, e in spans:
            found.append((masked[s:e].replace('–', '-').replace('—', '-'), s, e))
        masked = _mask(masked, spans)
    for rx in DATE_PATTERNS:
        spans = _spans(rx, masked)
        for s, e in spans:
            found.append((' '.join(masked[s:e].split()), s, e))
        masked = _mask(masked, spans)
    spans = _spans(FRACTION_PATTERN, masked)
    for s, e in spans:
        found.append((masked[s:e], s, e))
    masked = _mask(masked, spans)
    for m in UNIT_RE.finditer(masked):
        token = re.sub(r'\s+', '', m.group('num') + m.group('unit'))
        found.append((token, m.start(), m.end()))
    found.sort(key=lambda f: (f[1], f[2]))
    return found


def _context(text, start, end):
    return text[max(0, start - 20):min(len(text), end + 20)]


def collect(paras):
    """문단 목록에서 토큰별 개수와 처음 나온 자리(문단 번호, 앞뒤 20자 문맥)를 모은다."""
    counts = collections.Counter()
    first = {}
    for i, p in enumerate(paras, 1):
        for token, s, e in extract_facts(p):
            counts[token] += 1
            if token not in first:
                first[token] = (i, _context(p, s, e))
    return counts, first


def diff(old_paras, new_paras):
    """두 문서의 토큰을 개수까지 비교한다. 개수만 줄거나 는 것도 사라짐·새로 생김에 넣는다."""
    old_counts, old_first = collect(old_paras)
    new_counts, new_first = collect(new_paras)
    gone, added = [], []
    for token in set(old_counts) | set(new_counts):
        o, n = old_counts[token], new_counts[token]
        if o == n:
            continue
        if n < o:
            gone.append((token, o, n, old_first[token]))
        else:
            added.append((token, o, n, new_first[token]))
    # 읽는 순서(처음 나온 문단)로 늘어놓고, 같은 문단이면 많이 줄거나 는 것부터 보여 준다.
    gone.sort(key=lambda r: (r[3][0], -r[1], r[0]))
    added.sort(key=lambda r: (r[3][0], -r[2], r[0]))
    return old_counts, new_counts, gone, added


def render(old_paras, new_paras, show_all=False):
    """사람이 읽는 보고를 만든다. (보고 글, 사라짐, 새로 생김)을 돌려준다."""
    old_counts, new_counts, gone, added = diff(old_paras, new_paras)
    lines = ['이전 %d문단 · 사실 토큰 %d종 → 새 %d문단 · 사실 토큰 %d종'
             % (len(old_paras), len(old_counts), len(new_paras), len(new_counts))]
    for title, rows in (('[사라짐]', gone), ('[새로 생김]', added)):
        lines.append('')
        lines.append('%s %d종' % (title, len(rows)))
        shown = rows if show_all else rows[:LIMIT]
        for token, o, n, (para, ctx) in shown:
            lines.append('   %s  %d→%d  %d문단  …%s…' % (token, o, n, para, ctx))
        if len(rows) > len(shown):
            lines.append('   … 외 %d종 (--all로 전부 보기)' % (len(rows) - len(shown)))
    lines.append('')
    lines.append('차이 %d종. 의도한 숫자 변경인지 반영 기록과 대조한다. '
                 '숫자는 AGENTS.md와 제품 저장소 근거만 쓴다.' % (len(gone) + len(added)))
    return '\n'.join(lines), gone, added


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    show_all = '--all' in args
    paths = [a for a in args if a != '--all']
    if len(paths) != 2 or any(a.startswith('-') for a in paths):
        print('사용: python3 -I check_facts.py <이전.docx|.txt> <새.docx|.txt> [--all]', file=sys.stderr)
        return 2
    for p in paths:
        if not os.path.isfile(p):
            print('파일 없음: %s' % p, file=sys.stderr)
            return 2
    try:
        old_paras, new_paras = paragraphs(paths[0]), paragraphs(paths[1])
    except Exception as e:  # docx 구조가 아니거나 읽을 수 없는 파일
        print('읽지 못함: %s (%s)' % (', '.join(paths), e), file=sys.stderr)
        return 2
    text, gone, added = render(old_paras, new_paras, show_all)
    print(text)
    return 1 if (gone or added) else 0


if __name__ == '__main__':
    sys.exit(main())
