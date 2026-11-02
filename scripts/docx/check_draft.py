"""초안 md(위치/지우는 원문/넣는 내용/숫자 근거/분량)를 작업본 추출 텍스트와 대조한다.

사용: python3 -I scripts/docx/check_draft.py <초안.md> [추가 근거 파일...]
기준 본문: 환경변수 BASE_TXT. 없으면 AGENTS.md 상태 줄의 작업본과 같은 버전
04_참고자료/_추출/우리_vNN.txt. 추출이 없거나 작업본 docx보다 오래되면 종료 코드 1.
추출 다시 만들기: python3 -I scripts/docx/docx_to_text.py <작업본.docx> <우리_vNN.txt>
출력: 지우는 원문 불일치(설명문 줄도 잡히니 눈으로 거른다), 「~다/요」 종결 의심, 근거 없는 숫자 의심.
"""
import re, sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def working_pair():
    agents = open(os.path.join(ROOT, 'AGENTS.md'), encoding='utf-8').read()
    m = re.search(r'상태\([^)]+\): 작업본 `([^`]+\.docx)`', agents)
    if not m:
        sys.exit('AGENTS.md 상태 줄에서 작업본 docx를 찾지 못했다')
    docx_rel = m.group(1)
    ver = re.search(r'_v(\d+)\.docx$', docx_rel)
    if not ver:
        sys.exit('작업본 파일 이름에서 버전을 찾지 못했다: ' + docx_rel)
    docx_path = docx_rel if os.path.isabs(docx_rel) else os.path.join(ROOT, docx_rel)
    default_txt = os.path.join(ROOT, '04_참고자료', '_추출', '우리_v%s.txt' % ver.group(1))
    txt_path = os.environ.get('BASE_TXT', default_txt)
    regen = 'python3 -I scripts/docx/docx_to_text.py %s %s' % (docx_path, txt_path)
    if not os.path.isfile(txt_path):
        sys.exit('기준 본문 없음: %s\n%s' % (txt_path, regen))
    if os.path.isfile(docx_path) and os.path.getmtime(txt_path) + 1 < os.path.getmtime(docx_path):
        sys.exit('기준 본문이 작업본보다 오래됨: %s\n%s' % (txt_path, regen))
    return open(txt_path, encoding='utf-8').read(), agents, txt_path

base, agents, txt_path = working_pair()
src = base + agents
extra = sys.argv[2:]
for e in extra:
    src += open(e, encoding='utf-8').read()
norm = lambda s: re.sub(r'\s+', '', s)
V = norm(base)
SRC = norm(src)
d = open(sys.argv[1], encoding='utf-8').read()
blocks = re.split(r'\n(?=### )', d)
print('기준', txt_path)
print('블록', len(blocks) - 1, '지우는', d.count('**지우는 원문**'), '넣는', d.count('**넣는 내용**'), '숫자근거', d.count('**숫자 근거**'))
miss_orig = []
bad_end = []
new_nums = []
for b in blocks:
    head = b.split('\n', 1)[0]
    m = re.search(r'\*\*지우는 원문\*\*(.*?)(?=\n(?:- )?\*\*넣는 내용)', b, re.S)
    if m:
        for line in m.group(1).splitlines():
            t = line.strip().lstrip('-> ').strip()
            t = re.sub(r'^「|」$', '', t)
            if len(t) < 8:
                continue
            if norm(t) not in V:
                miss_orig.append((head, t[:90]))
    m = re.search(r'\*\*넣는 내용\*\*(.*?)(?=\n(?:- )?\*\*숫자 근거)', b, re.S)
    if m:
        body = m.group(1)
        for line in body.splitlines():
            cells = [c.strip() for c in line.split('|')] if '|' in line else [line.strip().lstrip('-> ').strip()]
            for c in cells:
                c = re.sub(r'[」\)\*]+$', '', c)
                if re.search(r'(다|요)\.?$', c) and len(c) > 6 and not re.search(r'(이상|이하|미만|초과)다?$', c):
                    bad_end.append((head, c[-40:]))
        for n in set(re.findall(r'\d[\d,\.]*\d(?:%p|%|건|종|원|억원|조원|ms|사|개|개월|년차)?|\d+(?:건|종|원|사|개월|년차)', body)):
            core = re.sub(r'(%p|%|건|종|원|억원|조원|ms|사|개|개월|년차)$', '', n)
            if len(core) >= 2 and core not in src and norm(core) not in SRC:
                new_nums.append((head, n))
print('지우는 원문 불일치', len(miss_orig))
for h, t in miss_orig[:25]:
    print('  ', h[:40], '|', t)
print('「~다/요」 종결 의심', len(bad_end))
for h, t in bad_end[:15]:
    print('  ', h[:40], '|', t)
print('근거 없는 숫자 의심', len(new_nums))
for h, t in sorted(set(new_nums))[:30]:
    print('  ', h[:40], '|', t)
