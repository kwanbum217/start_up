# 사업계획서 오케스트레이션 워커

2026-10-28에 확인한 기동 정본이다. 모델 이름이나 기동 방법이 바뀌면 이 파일만 고친다.
코디네이터 Run은 `run_4aa7ff0a34ed`다. 기준 제출본은 `02_과제/정부지원_사업계획서_제출용_별첨1_refac_bid_box_v17.docx`이고, 다음 수정은 v18로 저장한다.

이 폴더는 git 저장소다. 원격은 `https://github.com/kwanbum217/start_up.git`의 `main`이다. 워커는 그래도 현재 폴더에 둔다. `--worktree current`를 쓴다. 자식 워크트리에만 고치면 원격 동기화에서 빠진다. 커밋과 푸시의 정본은 `AGENTS.md`의 「원격 동기화」와 `.grok/skills/git-sync/SKILL.md`다.

## 풀

| 슬롯 | Orca `--agent` | 모델 ID | effort | 확인 |
|---|---|---|---|---|
| opencode | `opencode` | `opencode/deepseek-v4.1-flash-free` | 지정하지 않음 | `opencode models`에 있음 |
| codex | `codex` | `gpt-6-luna` | 지정하지 않음. 카탈로그 기본값은 medium | `~/.codex/models_cache.json` slug. 지원 effort는 low, medium, high, xhigh, max. `~/.codex/config.toml` 기본 모델은 `gpt-5.6-luna`라서 기동 때 모델을 빼면 luna 6이 아니다 |
| antigravity | `antigravity` | `gemini-3.8-flash-medium` | ID에 포함 | `agy models`의 Flash는 high, medium, low 세 ID다. 사용자가 단계만 말했으므로 medium을 기본으로 둔다. high는 `gemini-3.8-flash-high`, low는 `gemini-3.8-flash-low`. 2026-10-28에 토큰이 소진됐다. 다시 있다고 말하기 전에는 새 작성 워커로 띄우지 않는다. 진행 중이던 v17 작성은 qwen Nemotron 3 Ultra로 바꿨다 |
| qwen | Orca 에이전트 ID 없음 | `nvidia/nemotron-3-ultra-550b-a55b:free` | max | qwen 표시 이름은 `NVIDIA: Nemotron 3 Ultra (free)`. `~/.qwen/settings.json`의 `model.name`과 `model.reasoningEffort`가 이 값이다. CLI에 effort 플래그는 없다 |
| cmd-deepseek | 터미널 선생성(`command-code`는 `--model` 미지원) | `deepseek/deepseek-v4.1-flash` | 지정하지 않음. 화면 표시는 max | Command Code v1.88.2 `cmd --list-models`에 있음. 2026-10-29 화면에서 `deepseek-v4.1-flash-(latest) with max effort` 확인 |
| cmd-glm | 터미널 선생성 | `z-ai/glm-5.3-flash` | 지정하지 않음. 화면 표시는 high | `cmd --list-models`에 있음. 2026-10-29 화면에서 `glm-5.3-flash with high effort` 확인 |

`orca agent launch --agent qwen`은 `Unknown TUI agent "qwen"`이다. qwen만 터미널을 먼저 만들고 `worker-start --terminal`로 붙인다.

cmd(Command Code)는 Orca 에이전트 ID `command-code`가 있지만, `worker-start --model`은 이 에이전트에 모델을 넘기지 않는다(자체 설정 모델로 뜸). 그래서 cmd 두 슬롯도 qwen처럼 터미널을 먼저 만들고 붙인다.

2026-10-29 사용자 지정: 작성·리뷰 워커는 cmd 두 슬롯(DeepSeek v4.1 Flash, GLM 5.3 Flash)으로 돌리고, Claude가 코디네이터를 맡는다. 웨이브마다 한쪽이 작성, 다른 쪽이 리뷰다.

## 배정

한 웨이브에서 제출 docx를 고치는 워커는 한 명이다. 리뷰 워커는 그 작성 워커와 다른 에이전트이고, 파일을 고치지 않는다. 누가 작성인지는 웨이브마다 이 풀에서 고른다. 같은 docx를 두 워커가 동시에 수정하지 않는다.

작성 워커는 `AGENTS.md`의 백업, 새 버전, 명사 종결, 확인된 숫자만 쓴다. `02_과제/양식_원본`은 수정하지 않는다. 리뷰 워커의 완료 보고는 발견 사항이다. 고칠 일은 다음 작성 워커에게 넘긴다.

## 기동

세 에이전트는 작업 스펙을 정한 뒤 이렇게 띄운다. `--model`은 사용자가 이 풀로 지정했으므로 매번 넣는다. effort는 사용자가 그 슬롯에 말한 때만 넣는다. 지금은 qwen의 max뿐이다.

```text
orca orchestration worker-start --task <task_id> --worktree current --agent opencode --model opencode/deepseek-v4.1-flash-free --json
orca orchestration worker-start --task <task_id> --worktree current --agent codex --model gpt-6-luna --json
orca orchestration worker-start --task <task_id> --worktree current --agent antigravity --model gemini-3.8-flash-medium --json
```

영수증의 `launch.requested`와 `launch.effective`가 다를 수 있다. 실제로 뜬 모델은 `launch.effective`다.

cmd:

```text
orca terminal create --worktree current --title "cmd deepseek v4.1 flash" --command "cmd -t --skip-onboarding -m deepseek/deepseek-v4.1-flash --permission-mode accept-edits" --json
orca terminal create --worktree current --title "cmd glm 5.3 flash" --command "cmd -t --skip-onboarding -m z-ai/glm-5.3-flash --permission-mode accept-edits" --json
```

`orca terminal read --terminal <handle>`의 `# models:` 줄이 지정 모델일 때만 `worker-start --task <task_id> --terminal <handle> --worktree current --json`으로 붙인다. `accept-edits`는 파일 수정만 자동 승인한다. 셸 명령 승인 창이 뜨면 코디네이터가 화면을 보고 사용자에게 알린다.

qwen:

```text
orca terminal create --worktree current --title "qwen nemotron ultra free" --command "qwen -m nvidia/nemotron-3-ultra-550b-a55b:free" --json
```

화면의 모델이 `nvidia/nemotron-3-ultra-550b-a55b:free`이고 effort가 max인 뒤에만 붙인다. 다르면 그 터미널에 작업을 넣지 않는다. `~/.qwen/settings.json`은 이 작업에서 고치지 않는다.

```text
orca orchestration worker-start --task <task_id> --terminal <handle> --worktree current --json
```

OpenCode가 폴더 작업실에서 `--model`을 거부하면 그 영수증의 `failedStage`를 보고, 복구 안내를 따른다. 같은 명령으로 다시 띄우지 않는다.

## 회수

워커가 맡은 일이 수락되면 그 워커의 작업실을 회수한다. 결과 파일이 프로젝트 폴더에 있는 것을 확인한 뒤에만 한다. 제출 docx, 백업, `02_과제/양식_원본`은 회수 대상이 아니다.

1. 수락된 Dispatch는 `orca orchestration worker-release --dispatch <dispatch_id> --json`으로 닫는다. 출력은 아카이브에 남는다. 같은 터미널에 바로 이을 작업이 있을 때만 release 대신 그 터미널을 재사용한다.
2. 회수 목록은 `orca orchestration worker-list --run run_4aa7ff0a34ed --terminal-state reclaimable --json`이다. 이 목록의 `nextAction`이 release이면 그 명령으로 닫는다. 목록이 비어도 웨이브가 끝난 것은 아니다. 같은 Run의 워커 터미널이 아직 연결되어 있으면 그것까지 닫아야 한다. `orca terminal close --worktree current --all`은 코디네이터까지 닫으므로 쓰지 않는다.
3. 미리 만든 터미널에 `worker-start --terminal`로 붙인 경우 `worker-release`가 `retained`, 이유 `external_terminal`로 끝나고 프로세스는 그대로일 수 있다. 사용자가 그 터미널을 가져간 것이 아니고 이을 작업도 없으면 `orca terminal close --terminal <handle>`로 그 터미널만 닫는다. `ptyKilled`가 false이면 같은 핸들로 한 번 더 닫고, 터미널 목록에서 빠졌는지 확인한다. 사용자가 가져가 `user_takeover`로 남은 터미널은 닫지 않는다. 코디네이터 터미널도 닫지 않는다.
4. 워커는 `--worktree current`다. 이 경로는 프로젝트 폴더 자체이므로 `orca worktree rm --worktree current`를 하지 않는다. 회수 단계에서는 커밋하거나 푸시하지 않는다.
5. 자식 워크트리를 만들어 작업한 경우에만, 결과 파일이 `02_과제`에 복사된 뒤 그 자식만 `orca worktree rm --worktree <자식 선택자> --json`으로 제거한다. `refac_bid_box`의 main 워크트리와 이 프로젝트 폴더는 제거하지 않는다. `worktree rm`은 체크아웃 브랜치 삭제를 시도하므로 대상 경로를 출력으로 확인한 뒤에만 실행한다.
