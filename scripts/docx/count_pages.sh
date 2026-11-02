#!/bin/zsh
# 사업계획서 docx의 쪽 수를 한글이 실제로 그려지는 상태로 센다(macOS).
# 사용: zsh scripts/docx/count_pages.sh <파일.docx> [더 많은 docx ...]
#
# 이 Mac의 LibreOffice는 시스템 글꼴을 찾지 못해 한글을 빈칸(폭 거의 0)으로 그린다.
# 그 상태로 센 쪽 수는 실제보다 2~3쪽 적다(2026-11-02 확인: v23 11쪽 → 14쪽).
# 그래서 임시 LibreOffice 프로필에 Apple SD Gothic Neo를 넣고, 사본의 「맑은 고딕」을
# 그 글꼴로 바꿔 센다. 맑은 고딕과 글자 폭·줄 높이가 똑같지는 않으니 Word 쪽 수의 근사값이다.
set -e
SO=$(command -v soffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)
T=$(mktemp -d)
mkdir -p "$T/profile/user/fonts" "$T/out"
cp /System/Library/Fonts/AppleSDGothicNeo.ttc "$T/profile/user/fonts/"
for f in "$@"; do
  n=$(basename "$f" .docx)
  rm -rf "$T/x"; mkdir -p "$T/x"
  unzip -q "$f" -d "$T/x"; find "$T/x" -type l -delete
  for x in "$T"/x/word/*.xml; do sed -i '' 's/맑은 고딕/Apple SD Gothic Neo/g' "$x"; done
  (cd "$T/x" && zip -qXr "$T/$n.docx" .)
  "$SO" -env:UserInstallation=file://$T/profile --headless --convert-to pdf --outdir "$T/out" "$T/$n.docx" >/dev/null 2>&1
  echo "$n: $(pdfinfo "$T/out/$n.pdf" | awk '/Pages/{print $2}')쪽 (PDF: $T/out/$n.pdf)"
done
