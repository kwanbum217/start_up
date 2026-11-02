"""작업본 docx를 검토용 구글 문서에 올리고, 구글 변환에서 생기는 차이를 보정한 뒤 쪽 수를 잰다.

사용:
  python3 -W ignore fit_gdoc.py <계획서.docx>            # 올리기 + 보정 + 쪽 수
  python3 -W ignore fit_gdoc.py --fix-only               # 이미 올린 문서에 보정만 다시

구글 문서는 맑은 고딕을 그리지 못하고(PDF에서 굴림으로 대체), docx의 칸 여백·칸 줄 간격·
구역 종류 일부를 버린다. 그대로 두면 v25 기준 15쪽이 된다. 이 스크립트는 「검토용 사본에만」
아래 보정을 한다. 제출 docx는 바꾸지 않는다.

1. 올리기 전 임시 사본에서 「다음 쪽」 구역 나누기를 「이어서」로 바꾼다(일반현황 뒤 반 쪽 빈칸).
   구글 API로는 올린 뒤에 구역 종류를 바꾸거나 지울 수 없다.
2. 쪽 여백을 docx pgMar 값으로(위아래 1417·좌우 1134 dxa, 머리말·꼬리말 850).
3. 표 칸 여백을 칸마다 위아래 2pt·좌우 4pt(docx tblCellMar 40/80 dxa). 표 전체 범위로 한 번에
   넣으면 무시되므로 칸마다 넣는다.
4. 표 칸 문단 줄 간격 90%(docx 216 auto), 표 행의 「쪽 넘김 금지」 해제(docx cantSplit).
5. 글꼴을 Nanum Gothic으로(구글이 실제로 그리는 한글 글꼴. 맑은 고딕은 이름만 남고 굴림으로 그려짐).
6. 표의 빈 행 삭제(양식의 「공란 유지」는 제출 docx에서 지킨다. 검토용 사본에서만 지움).
7. 「다음 단락과 함께」를 풀고, 장 배너(그림 문단)에만 다시 건다.
그 뒤 PDF로 내보내 쪽 수와 쪽마다 마지막 줄을 출력한다.
"""
import os, re, subprocess, sys, tempfile, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sync_gdoc as g
from googleapiclient.errors import HttpError

pt = lambda v: {'magnitude': v, 'unit': 'PT'}


def variant(docx):
    """「다음 쪽」 구역을 「이어서」로 바꾼 임시 사본(파일 이름은 그대로, 제목이 같게)."""
    out = os.path.join(tempfile.mkdtemp(), os.path.basename(docx))
    with zipfile.ZipFile(docx) as zi, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == 'word/document.xml':
                s = data.decode('utf-8')
                s = re.sub(r'<w:sectPr\b.*?</w:sectPr>',
                           lambda m: m.group(0) if '<w:type ' in m.group(0)
                           else m.group(0).replace('<w:pgSz', '<w:type w:val="continuous"/><w:pgSz', 1),
                           s, flags=re.S)
                data = s.encode('utf-8')
            zo.writestr(item, data)
    return out


def text(el):
    s = ''
    for e in el['paragraph']['elements']:
        s += e.get('textRun', {}).get('content', '')
    return s.strip()


def cells_text(cont):
    s = ''
    for el in cont:
        if 'paragraph' in el:
            s += text(el) + ('[IMG]' if any('inlineObjectElement' in e for e in el['paragraph']['elements']) else '')
        elif 'table' in el:
            s += '[TBL]'
    return s


def batch(docs, reqs):
    for i in range(0, len(reqs), 400):
        docs.documents().batchUpdate(documentId=g.DOC_ID, body={'requests': reqs[i:i + 400]}).execute()


def fix(docs):
    get = lambda: docs.documents().get(documentId=g.DOC_ID).execute()
    d = get()
    body = d['body']['content']
    # 2. 쪽 여백
    m = {'marginTop': pt(1417 / 20), 'marginBottom': pt(1417 / 20), 'marginLeft': pt(1134 / 20), 'marginRight': pt(1134 / 20)}
    reqs = [{'updateDocumentStyle': {'documentStyle': dict(m, marginHeader=pt(42.5), marginFooter=pt(42.5)),
                                     'fields': 'marginTop,marginBottom,marginLeft,marginRight,marginHeader,marginFooter'}}]
    for el in body:
        if 'sectionBreak' in el and 'startIndex' in el:
            reqs.append({'updateSectionStyle': {'range': {'startIndex': el['startIndex'], 'endIndex': el['endIndex']},
                                                'sectionStyle': m, 'fields': ','.join(m)}})
    # 3·4. 칸 여백, 칸 줄 간격, 행 쪽 넘김
    for el in body:
        if 'table' not in el:
            continue
        loc = {'index': el['startIndex']}
        reqs.append({'updateTableRowStyle': {'tableStartLocation': loc, 'tableRowStyle': {'preventOverflow': False},
                                             'fields': 'preventOverflow'}})
        for i, row in enumerate(el['table']['tableRows']):
            for j, c in enumerate(row['tableCells']):
                reqs.append({'updateTableCellStyle': {'tableRange': {'tableCellLocation': {'tableStartLocation': loc, 'rowIndex': i, 'columnIndex': j},
                                                                     'rowSpan': 1, 'columnSpan': 1},
                                                      'tableCellStyle': {'paddingTop': pt(2), 'paddingBottom': pt(2), 'paddingLeft': pt(4), 'paddingRight': pt(4)},
                                                      'fields': 'paddingTop,paddingBottom,paddingLeft,paddingRight'}})
                for p in c['content']:
                    if 'paragraph' in p:
                        reqs.append({'updateParagraphStyle': {'range': {'startIndex': p['startIndex'], 'endIndex': p['endIndex']},
                                                              'paragraphStyle': {'lineSpacing': 90}, 'fields': 'lineSpacing'}})
    # 5. 글꼴
    reqs.append({'updateTextStyle': {'range': {'startIndex': 1, 'endIndex': body[-1]['endIndex'] - 1},
                                     'textStyle': {'weightedFontFamily': {'fontFamily': 'Nanum Gothic'}}, 'fields': 'weightedFontFamily'}})
    batch(docs, reqs)
    # 6. 빈 행 삭제(뒤에서부터)
    d = get()
    reqs = []
    for el in reversed(d['body']['content']):
        if 'table' in el:
            rows = el['table']['tableRows']
            for i in reversed(range(len(rows))):
                if not any(cells_text(c['content']) for c in rows[i]['tableCells']):
                    reqs.append({'deleteTableRow': {'tableCellLocation': {'tableStartLocation': {'index': el['startIndex']}, 'rowIndex': i, 'columnIndex': 0}}})
    if reqs:
        batch(docs, reqs)
    print('빈 행 삭제', len(reqs))
    # 7. 다음 단락과 함께: 모두 풀고 장 배너에만
    d = get()
    body = d['body']['content']
    reqs = []

    def walk(cont):
        for el in cont:
            if 'table' in el:
                for r in el['table']['tableRows']:
                    for c in r['tableCells']:
                        walk(c['content'])
            elif 'paragraph' in el and el['paragraph'].get('paragraphStyle', {}).get('keepWithNext'):
                reqs.append({'updateParagraphStyle': {'range': {'startIndex': el['startIndex'], 'endIndex': el['endIndex']},
                                                      'paragraphStyle': {'keepWithNext': False}, 'fields': 'keepWithNext'}})
    walk(body)
    for k, el in enumerate(body):
        if 'paragraph' not in el:
            continue
        img = any('inlineObjectElement' in e for e in el['paragraph']['elements'])
        nxt = body[k + 1] if k + 1 < len(body) else {}
        if img and not text(el) and 'paragraph' in nxt and text(nxt)[:1] in ('○', '<', '◦'):
            reqs.append({'updateParagraphStyle': {'range': {'startIndex': el['startIndex'], 'endIndex': el['endIndex']},
                                                  'paragraphStyle': {'keepWithNext': True}, 'fields': 'keepWithNext'}})
    batch(docs, reqs)


def pages(drive):
    pdf = os.path.join(tempfile.mkdtemp(), 'gdoc.pdf')
    open(pdf, 'wb').write(drive.files().export(fileId=g.DOC_ID, mimeType='application/pdf').execute())
    n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
    print(f'구글 PDF {n}쪽 ({pdf})')
    for p in range(1, n + 1):
        t = subprocess.run(['pdftotext', '-f', str(p), '-l', str(p), pdf, '-'], capture_output=True, text=True).stdout
        last = [l for l in t.splitlines() if l.strip()]
        print(f'  {p}쪽 끝: {last[-1][:40] if last else "(빈 쪽)"}')


def main():
    drive, docs = g.services()
    if sys.argv[1:] and sys.argv[1] != '--fix-only':
        v = variant(sys.argv[1])
        subprocess.run([sys.executable, '-W', 'ignore', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sync_gdoc.py'), v], check=True)
    fix(docs)
    pages(drive)


if __name__ == '__main__':
    main()
