---
name: git-sync
description: >
  start_up 폴더를 GitHub origin https://github.com/kwanbum217/start_up.git 의 main과 맞춘다.
  커밋, 푸시, 풀, 원격 동기화, git commit, git push, remote sync, /git-sync 를 할 때 쓴다.
  이 폴더의 파일 변경을 마친 뒤 원격에 남길 때도 쓴다.
---

# git-sync

정책의 정본은 `AGENTS.md`의 「원격 동기화」다. 그 절과 이 순서가 다르면 `AGENTS.md`를 따른다. 규칙을 바꿀 때는 두 곳을 같이 고친다.

작업 디렉터리는 `/Users/kwanbum/Documents/korea_IT/start_up`이다. 다른 저장소(제품 저장소 `refac_bid_box` 포함)에는 이 절차를 적용하지 않는다.

## 언제 실행하나

- 사용자가 커밋, 푸시, 풀, 동기화, 원격 반영을 말한 때
- 이 폴더의 추적 대상 파일을 고치는 작업을 마친 때. 사용자가 그 작업은 올리지 말라고 하면 커밋과 푸시를 하지 않는다

올릴 변경이 없으면 커밋하지 않는다. 상태만 확인하고 이미 원격과 같다고 보고한다.

## 순서

1. `git status`, `git diff`, `git diff --cached`, `git log -5 --oneline`으로 올릴 내용을 본다.
2. `git remote get-url origin`이 `https://github.com/kwanbum217/start_up.git`인지 확인한다. `origin`이 없을 때만 그 주소로 추가한다. 다른 주소가 이미 있으면 덮어쓰지 말고 사용자에게 알린다.
3. 비밀값(토큰, 비밀번호, `.env`, 인증서, 개인 키, 자격 증명 JSON)이 스테이지에 있으면 커밋하지 않는다. `.DS_Store`, `.grok/doc-review/`, `.commandcode/`는 올리지 않는다.
4. `origin/main`을 `git fetch origin`으로 받은 뒤, 로컬 `main`보다 앞에 있으면 `git pull --ff-only origin main`으로 맞춘다. fast-forward가 아니면 푸시하지 않고 사용자에게 알린다.
5. 확인한 파일만 스테이지한다. 커밋 메시지는 한국어로, 무엇을 왜 넣는지 쓴다.
6. `git push -u origin main`으로 올린다. `git push --force`, 이력 재작성, 원격 브랜치 삭제는 사용자가 그 명령을 명시하기 전에는 하지 않는다.
7. `git status`가 깨끗한지, `git rev-parse HEAD`와 `git rev-parse origin/main`이 같은지 확인한 뒤에 올렸다고 보고한다.

오케스트레이션 워커는 이 폴더(`--worktree current`)에서 고친다. 자식 워크트리에만 있는 커밋은 이 원격에 올라가지 않는다. `orca worktree rm`으로 이 폴더를 지우지 않는다. 워커 회수 절차는 `prompts/orca_workers.md`에 있고, 그 절차 안에서 푸시하지 않는다. 푸시는 회수가 끝난 뒤 이 스킬로 한다.
