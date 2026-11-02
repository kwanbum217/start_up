"""커밋 인덱스에 올리면 안 되는 경로가 있는지 본다.

사용: python3 -I scripts/docx/check_push_scope.py
종료 코드 1: 04_참고자료/_추출, 02_과제/양식_원본, 비밀값으로 보이는 파일이 스테이지에 있을 때.
"""
import subprocess, sys

r = subprocess.run(['git', 'diff', '--cached', '--name-only', '-z'], capture_output=True)
if r.returncode != 0:
    sys.exit(r.stderr.decode() or 'git diff 실패')
names = [n for n in r.stdout.decode().split('\0') if n]
bad = []
for n in names:
    low = n.lower()
    base = low.rsplit('/', 1)[-1]
    if n.startswith('04_참고자료/_추출/') or n.startswith('02_과제/양식_원본/'):
        bad.append(n)
    elif base == '.env' or base.startswith('.env.') or base.endswith(('.pem', '.key', '.p12')):
        bad.append(n)
    elif base.endswith('.json') and any(s in base for s in (
            'credential', 'service-account', 'service_account', 'client_secret', 'booming-primer')):
        bad.append(n)
print('스테이지 %d개' % len(names))
if bad:
    print('커밋하면 안 되는 경로:')
    for n in bad:
        print(' ', n)
    sys.exit(1)
print('금지 경로 없음')
