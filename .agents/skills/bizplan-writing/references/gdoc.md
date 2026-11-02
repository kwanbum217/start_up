# 검토용 구글 문서로 옮기기

사용자는 macOS를 쓰고 Word의 맑은 고딕 렌더링을 바로 보기 어렵다. 그래서 작업본 docx를 **항상 같은 구글 문서 링크**에 변환해 넣고, 사용자는 그 링크에서 서식과 디자인을 본다.

- 링크: https://docs.google.com/document/d/1-CqlcQz9wHpX5gDnBmfQJeFSAf3AIOWSgaxV3_veenU/edit
- 이 문서는 검토용이다. 쪽 나눔은 Word와 다를 수 있어 구글 문서·구글 PDF의 쪽 수로 12쪽 판정을 하지 않는다.

## 왜 docx를 통째로 올리나

마크다운이나 문단 단위로 옮기면 표의 셀 병합이 풀리고 글꼴·여백이 달라진다(2026-10-29에 겪음). Drive API `files.update`로 같은 파일 ID에 docx를 올리면 구글이 docx를 변환해 본문 전체를 바꾸고, 링크·공유 설정·댓글 스레드 위치는 유지된다. 이전 본문은 구글 문서의 버전 기록에 남는다.

## 절차

1. 덮어쓰기는 사용자가 보는 문서를 바꾸므로, 올리기 전에 사용자에게 「검토용 구글 문서를 vNN으로 바꿀지」 확인한다(이번 세션에 이미 허락받았으면 생략).
2. 지금 상태를 읽는다(읽기 전용):
   ```bash
   python3 -W ignore .agents/skills/bizplan-writing/scripts/sync_gdoc.py --check
   ```
3. 올리고 보정한다(2026-11-02부터 이 명령 하나로 한다):
   ```bash
   python3 -W ignore .agents/skills/bizplan-writing/scripts/fit_gdoc.py 02_과제/<작업본>.docx
   ```
   `sync_gdoc.py`로 docx만 올리면 구글이 칸 여백·칸 줄 간격·구역 종류를 버리고, 맑은 고딕 대신 굴림으로 그려 v25가 15쪽이 된다. `fit_gdoc.py`는 검토용 사본에만 이렇게 보정한다(자세한 내용은 스크립트 첫머리).
   - 「다음 쪽」 구역을 「이어서」로 바꾼 임시 사본을 올린다.
   - 쪽 여백, 칸 여백, 칸 줄 간격 90%를 docx 값으로 맞추고, 행 쪽 넘김 금지를 푼다.
   - 글꼴을 나눔고딕으로 바꾸고, 빈 표 행을 지우고, 「다음 단락과 함께」는 장 배너에만 남긴다.
   끝에 구글 PDF 쪽 수와 쪽마다 마지막 줄을 출력한다. 마지막 줄이 「< … >」 표 제목이면 그 제목이 쪽 끝에 홀로 남은 것이다.
4. 이 보정 때문에 구글 사본은 제출 docx와 다르다. 글꼴(나눔고딕), 빈 행 없음, 1쪽 뒤 구역 나누기가 없는 점이 다르다. 구글 쪽 수는 검토용 참고값이고, 제출 쪽 수는 `scripts/docx/count_pages.sh`와 Word로 판정한다.
5. 확인할 값(`--check`): 표·그림 수가 docx와 같음, 글꼴이 Nanum Gothic 하나(보정 후), 여백 위아래 2.5·좌우 2.0cm. 다르면 사용자에게 보고하고, 구글 문서 쪽에서 손으로 고치지 않는다(다음 반영 때 덮어써져 사라진다). 고칠 것은 docx에서 고친다.
6. `AGENTS.md` 「현재 사업계획서」의 구글 문서 문장(어느 버전을 넣었는지, 날짜)을 갱신하고 사용자에게 링크를 준다.

## 실행 환경

- `python3 -I`가 아니라 `python3`로 실행한다. 구글 라이브러리(`google-api-python-client`, `google-auth`)가 사용자 site-packages에 있어 `-I`로는 불러오지 못한다. 이 스크립트는 프로젝트 안의 신뢰할 수 있는 파일만 읽는다.
- 인증은 서비스 계정 키 파일이다. 경로는 환경변수 `GOOGLE_APPLICATION_CREDENTIALS`(또는 `SERVICE_ACCOUNT_PATH`)에서, 없으면 `~/.claude.json`의 `google-docs` MCP 설정에서 읽는다. 키 경로나 내용을 저장소 파일에 적지 않는다(저장소가 공개다).
- Grok·Codex처럼 `~/.claude.json`이 없는 에이전트는 환경변수로 키 경로를 넘긴다. 키 파일 위치는 사용자에게 묻는다.
- 403·404가 나면 대상 문서가 키의 `client_email` 계정에 편집자로 공유됐는지 사용자에게 확인을 부탁한다.
- 구글 문서의 댓글은 `google-docs` MCP(`listComments` 등)로 읽을 수 있다. 사용자가 구글 문서에 남긴 의견은 docx 수정 지시로 받아 1절 흐름으로 반영한다.
