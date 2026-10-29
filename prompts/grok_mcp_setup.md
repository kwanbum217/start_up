# Grok MCP 설정 프롬프트

아래 "프롬프트" 부분을 그대로 복사해, 프로젝트 폴더(`/Users/kwanbum/Documents/korea_IT/start_up`)에서 실행한 Grok에 붙여 넣는다.

---

## 프롬프트

이 프로젝트에서 Claude Code가 쓰는 MCP 서버를 너(Grok)도 똑같이 쓸 수 있도록, **이 프로젝트 범위**로 MCP를 설정해 줘. 시작 전에 프로젝트 루트의 `AGENTS.md`를 읽고 그 규칙(한국어 보고, 되돌리기 어려운 작업 전 확인)을 따라.

### 전제

- 프로젝트 경로: `/Users/kwanbum/Documents/korea_IT/start_up`. 원격은 `https://github.com/kwanbum217/start_up.git`의 `main`이다. Grok은 이 폴더를 현재 디렉터리로 두고 작업한다. 설정 파일은 `./.grok/config.toml`(프로젝트 범위)이다.
- 설정은 `grok mcp add --scope project ...` 명령으로 추가한다. 직접 편집이 필요하면 `.grok/config.toml`만 편집한다. `~/.grok/config.toml`(사용자 범위)은 건드리지 않는다. 그 안의 기존 `google-docs` 서버는 그대로 둔다.
- 참고 문서: `~/.grok/docs/user-guide/07-mcp-servers.md`

### 보안 규칙 (반드시 지킬 것)

1. 토큰·키·비밀번호 값을 설정 파일에 직접 쓰지 않는다. 모든 비밀값은 `${VAR}` 형식으로 참조한다.
2. CLI로 추가할 때는 셸이 `${VAR}`를 미리 풀지 않도록 **작은따옴표**로 감싼다. 예: `-e 'SLACK_BOT_TOKEN=${SLACK_BOT_TOKEN}'`
3. `~/.claude.json`, `~/.codex/`, `~/.hermes/` 등 다른 에이전트 설정 파일에서 토큰 값을 읽거나 복사하지 않는다.
4. 환경변수 값이 필요하면 직접 설정하지 말고, 필요한 변수 이름을 나에게 알려 준다.
5. 작업이 끝나면 `.grok/config.toml`을 보여 주고, 비밀값이 평문으로 들어가지 않았는지 확인한 결과를 함께 보고한다.

### 등록할 서버

**A. 원격 서버 (HTTP, OAuth 로그인)**: Claude의 claude.ai 커넥터와 같은 서버다.

| 이름 | URL |
|---|---|
| notion | `https://mcp.notion.com/mcp` |
| slack | `https://mcp.slack.com/mcp` |
| google-drive | `https://drivemcp.googleapis.com/mcp/v1` |
| gmail | `https://gmailmcp.googleapis.com/mcp/v1` |
| google-calendar | `https://calendarmcp.googleapis.com/mcp/v1` |
| railway | `https://mcp.railway.com` |

- 추가 예: `grok mcp add --scope project --transport http notion https://mcp.notion.com/mcp`
- 추가한 뒤에는 내가 `/mcps` 창에서 각 서버를 골라 `i`를 눌러 브라우저로 로그인한다. 로그인은 네가 대신하지 말고 나에게 안내만 해 준다.
- 어떤 서버가 Grok의 OAuth 클라이언트를 거부하거나 로그인에 실패하면, 우회하지 말고 그 서버를 `enabled = false`로 바꾼 뒤 오류 메시지를 그대로 보고한다.

**B. 로컬 서버 (stdio)**

| 이름 | 명령 | 환경변수 |
|---|---|---|
| google-slides | `npx -y mcp-google-slides` | `GOOGLE_APPLICATION_CREDENTIALS`, `GOOGLE_CLOUD_PROJECT` |
| mcp-gsheets | `npx -y mcp-gsheets` | `GOOGLE_APPLICATION_CREDENTIALS`, `GOOGLE_PROJECT_ID` |
| puppeteer | `npx -y @modelcontextprotocol/server-puppeteer` | 없음 |
| chrome-devtools | `npx -y chrome-devtools-mcp@latest --autoConnect --filesystemRoot=/Users/kwanbum/Documents/korea_IT/start_up` | `CHROME_DEVTOOLS_MCP_NO_USAGE_STATISTICS=1` |
| slack-admin | `node /Users/kwanbum/.config/opencode/slack-delete-mcp.mjs --stdio` | `SLACK_BOT_TOKEN` |

- 추가 예:
  ```bash
  grok mcp add --scope project google-slides \
    -e 'GOOGLE_APPLICATION_CREDENTIALS=${GOOGLE_APPLICATION_CREDENTIALS}' \
    -e 'GOOGLE_CLOUD_PROJECT=${GOOGLE_CLOUD_PROJECT}' \
    -- npx -y mcp-google-slides
  ```
- `npx`로 실행하는 서버는 첫 실행이 느리므로 `.grok/config.toml`에서 `startup_timeout_sec = 120`을 넣는다.
- `slack-admin`은 Slack 메시지를 **삭제**하는 도구다. 등록은 하되, 이 서버의 도구를 호출하기 전에는 매번 나에게 확인을 받는다.
- `chrome-devtools`는 Claude 목록에 없는 Grok용 서버다. 이미 켜 둔 Chrome(144 이상, 이 Mac은 156)에 붙는다. 첫 사용 전에 Chrome에서 `chrome://inspect/#remote-debugging`의 원격 디버깅을 켜고, 연결 허용을 누른다.

**C. 등록하지 않을 것**

- Claude Docs, Claude in Chrome: Claude 전용 기능이라 Grok에서 쓸 수 없다.
- harness: Claude에서도 연결이 안 되는 상태(Docker 필요)라 제외한다.

### 진행 순서

1. `grok mcp list`로 현재 상태를 확인한다.
2. A·B의 서버를 프로젝트 범위로 추가한다.
3. `grok mcp list`에서 각 서버가 `(project)`로 표시되는지 확인한다.
4. 각 서버에 `grok mcp doctor <이름>`을 실행한다. 환경변수가 없어서 실패하는 서버는 어떤 변수가 빠졌는지 정리한다.
5. 원격 서버 6개의 로그인 방법(`/mcps` → 서버 선택 → `i`)을 안내한다.
6. 아래 형식으로 보고한다.

### 보고 형식

| 서버 | 종류 | 등록 | doctor 결과 | 남은 조치 |
|---|---|---|---|---|

그리고 다음도 함께 알려 준다.
- 내가 `~/.zshrc`에 export해야 할 환경변수 이름 목록(값은 쓰지 말 것)
- 생성된 `.grok/config.toml` 전문
- 비밀값 평문 포함 여부 점검 결과
