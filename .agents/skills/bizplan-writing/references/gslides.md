# 검토용 구글 슬라이드로 옮기기

사용자는 macOS를 쓰고 PowerPoint의 맑은 고딕 렌더링을 바로 보기 어렵다. 그래서 작업본 pptx를 **항상 같은 구글 슬라이드 링크**에 변환해 넣고, 사용자는 그 링크에서 디자인과 시안을 본다.

- 링크: https://docs.google.com/presentation/d/1freHyZ0yO83FKwQ_UziZXhc3czPzsfHGHdEJlUq4Nvw/edit
- 이 슬라이드는 검토용이다. 구글 변환은 글꼴과 도형 위치를 pptx와 조금 다르게 그릴 수 있다. 제출본은 `02_과제/피치덱/`의 pptx다.

## 왜 pptx를 통째로 올리나

장마다 다시 그리면 좌표, 그림, 표가 어긋난다. Drive API `files.update`로 같은 파일 ID에 pptx를 올리면 구글이 pptx를 변환해 장 전체를 바꾸고, 링크와 공유 설정은 유지된다. 이전 장은 구글 슬라이드의 버전 기록에 남는다.

## 절차

1. 덮어쓰기는 사용자가 보는 슬라이드를 바꾸므로, 올리기 전에 사용자에게 「검토용 구글 슬라이드를 vNN으로 바꿀지」 확인한다(이번 세션에 이미 허락받았으면 생략).
2. 지금 상태를 읽는다(읽기 전용):
   ```bash
   python3 -W ignore .agents/skills/bizplan-writing/scripts/sync_gslides.py --check
   ```
3. 올리고 글꼴을 맞춘다:
   ```bash
   python3 -W ignore .agents/skills/bizplan-writing/scripts/sync_gslides.py 02_과제/피치덱/<작업본>.pptx
   ```
   구글은 맑은 고딕을 그리지 못한다. `sync_gslides.py`는 검토용 사본에만 이렇게 한다.
   - 제출 pptx는 열기만 하고 바꾸지 않는다. 글꼴은 맑은 고딕 그대로다.
   - 임시 사본의 `typeface="Malgun Gothic"`과 `typeface="맑은 고딕"`(테마의 한글 script 포함)만 `Nanum Gothic`으로 바꾼 뒤 그 사본을 올린다.
   - 막대는 생성기에서 도형으로 그린다. 네이티브 막대 차트는 변환 때 항목 이름·순서·숫자 형식이 깨진다. 도넛 차트는 변환이 된다. 이 스크립트는 차트나 도형을 다시 그리지 않는다.
4. 이 보정 때문에 구글 사본은 제출 pptx와 글꼴만 다르다(나눔고딕). 구글에서 어긋나 보이는 여백·줄바꿈은 제출본의 결함으로 단정하지 않는다. 고칠 것은 생성기 `scripts/pitchdeck/build_pitchdeck.js`에서 고치고, 버전을 올린 새 pptx를 만든 뒤 다시 올린다. 구글 슬라이드 위에서 손으로 고치지 않는다(다음 반영 때 덮어써져 사라진다).
5. 확인할 값(`--check`와 올리기 출력): 장 수가 pptx와 같음, 한글 글꼴이 Nanum Gothic 또는 나눔고딕, 맑은 고딕·Malgun Gothic이 0. 다르면 사용자에게 보고한다.
6. `AGENTS.md` 「피치덱」에 어느 버전을 넣었는지 한 줄을 갱신하고 사용자에게 링크를 준다.

피치덱을 고칠 때의 사실·디자인 정본은 `AGENTS.md` 「피치덱」이다. 생성기를 고쳐 새 pptx를 만든 뒤에 이 절차로 올린다.

## 실행 환경

- `python3 -I`가 아니라 `python3`로 실행한다. 구글 라이브러리(`google-api-python-client`, `google-auth`)가 사용자 site-packages에 있어 `-I`로는 불러오지 못한다. 이 스크립트는 프로젝트 안의 신뢰할 수 있는 파일만 읽는다.
- 인증은 서비스 계정 키 파일이다. 경로는 `sync_gdoc.py`와 같이 환경변수 `GOOGLE_APPLICATION_CREDENTIALS`(또는 `SERVICE_ACCOUNT_PATH`)에서, 없으면 `~/.claude.json`의 `google-docs` MCP 설정에서 읽는다. 키 경로나 내용을 저장소 파일에 적지 않는다(저장소가 공개다).
- Grok·Codex처럼 `~/.claude.json`이 없는 에이전트는 환경변수로 키 경로를 넘긴다. 키 파일 위치는 사용자에게 묻는다.
- 403·404가 나면 대상 슬라이드가 키의 `client_email` 계정에 편집자로 공유됐는지 사용자에게 확인을 부탁한다.
- 사용자가 링크에서 알려 준 수정은 생성기를 고치는 지시로 받는다.
