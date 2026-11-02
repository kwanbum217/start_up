# 예시 계획서 기준 서식 맞추기

## 정본

서식 수치(글꼴, 장 제목·소제목·본문·표 글자 크기, 줄·문단 간격, 위아래 2.5cm·좌우 2cm 여백, 자간 -0.15pt, 장평 100%, 글머리 내어쓰기, 표 음영·괘선)의 정본은 `AGENTS.md`의 「1. 서식 규칙」과 「1-1. 시안성 규칙」이다. 값은 그 표에서 읽는다. 값을 바꾸면 그 표와 `scripts/docx/apply_style.py`, `scripts/docx/check_docx.py`를 같이 고친다.

예시 파일:
- `04_참고자료/2026년 앱_다취업.docx` (여백 위아래 1701·좌우 1134, 머리말 720)
- `04_참고자료/2026년 제조_맘맘.docx` (여백 위아래 1417·좌우 1134, 머리말 850)
- 본문 텍스트와 구성 분석: `04_참고자료/_추출/예시_*.txt`, `분석_*.md`

예시는 서식과 구성만 참고한다. 예시의 문장·숫자·사례를 우리 계획서로 옮기지 않는다.

## 적용 순서

```bash
W=<에이전트 임시 폴더>/work && rm -rf "$W" && mkdir -p "$W"
unzip -q <원본.docx> -d "$W" && find "$W" -type l -delete
python3 -I scripts/docx/fix_docx.py "$W"
python3 -I scripts/docx/normalize_layout.py "$W"   # lxml 필요
python3 -I scripts/docx/apply_style.py "$W"
(cd "$W" && zip -qXr <절대경로>/<새파일.docx> .)
python3 -I scripts/docx/check_docx.py <새파일.docx>
```

쪽 수가 12쪽을 넘으면 1-1절의 줄이는 순서(소제목 → 표 제목 → 장 제목)를 한 단계씩 따르고 매번 다시 잰다. 본문 10pt·표 본문 9pt·자간 -3 아래로는 줄이지 않고, 그래도 넘으면 내용을 줄이자고 사용자에게 보고한다.

## 예시에서 가져온 구성 습관

- 개요표(창업 아이템 개요)만 읽어도 고객·문제·해결·차별·자금이 보이게 칸 안에 작은 표를 둔다.
- 1장 문제 인식은 짧게, 2·3장(실현 가능성, 성장전략)에 분량을 준다.
- 표마다 「< 표 제목 >」을 달고, 그림 아래 9pt 캡션.
- 4장 팀은 강의 4회차 방식: 대표자 경력 연수·주요 업무 근거, 팀원은 필요 분야·역량·고용 예정 시기.

## 예시를 다시 재야 할 때

예시 파일이 바뀌거나 새 예시를 받으면 docx XML에서 직접 잰다(다른 도구의 진단은 그대로 믿지 않는다).

```bash
unzip -p "<예시.docx>" word/document.xml | grep -o '<w:pgMar[^>]*>' | sort | uniq -c      # 여백
unzip -p "<예시.docx>" word/document.xml | grep -o '<w:sz w:val="[0-9]*"/>' | sort | uniq -c | sort -rn | head   # 글자 크기(반 pt)
unzip -p "<예시.docx>" word/document.xml | grep -o '<w:spacing w:val="-\?[0-9]*"/>' | sort | uniq -c | sort -rn | head   # 자간(1/20 pt)
unzip -p "<예시.docx>" word/document.xml | grep -o 'w:line="[0-9]*" w:lineRule="[a-zA-Z]*"' | sort | uniq -c | sort -rn | head   # 줄 간격
```

잰 값이 1-1절과 다르면 바로 고치지 말고, 차이와 쪽 수 영향을 사용자에게 보고한 뒤 결정을 받는다.

## macOS에서 볼 때 주의

맑은 고딕이 macOS에 기본으로 없어서 미리보기·Pages·Quick Look의 줄밀림과 표 폭은 믿지 않는다(AGENTS.md 3절). 사용자가 눈으로 볼 곳은 검토용 구글 문서(`references/gdoc.md`)이고, 최종 쪽 수와 줄바꿈은 Word(가능하면 Windows)나 Word에서 내보낸 PDF로 확인한다.
