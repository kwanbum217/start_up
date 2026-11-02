"""사업계획서 docx를 검토용 구글 문서(같은 링크)에 변환해 넣고, 결과를 확인한다.

사용:
  python3 -I sync_gdoc.py --check                  # 읽기만: 지금 구글 문서의 제목·표·그림·글꼴·여백
  python3 -I sync_gdoc.py <계획서.docx>             # 같은 문서 ID에 docx를 덮어 변환한 뒤 확인
  python3 -I sync_gdoc.py <계획서.docx> --doc <ID>  # 다른 문서 ID

키 파일 경로는 환경변수 GOOGLE_APPLICATION_CREDENTIALS 또는 SERVICE_ACCOUNT_PATH에서,
없으면 ~/.claude.json의 google-docs MCP 설정(env)에서 읽는다. 키 내용은 출력하지 않는다.
대상 문서는 그 키의 client_email에 편집자로 공유돼 있어야 한다.
덮어쓰면 구글 문서의 이전 본문은 버전 기록(파일 > 버전 기록)에만 남는다.
"""
import argparse, collections, json, os, sys, zipfile

DOC_ID = '1-CqlcQz9wHpX5gDnBmfQJeFSAf3AIOWSgaxV3_veenU'
DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


def key_path():
    for v in ('GOOGLE_APPLICATION_CREDENTIALS', 'SERVICE_ACCOUNT_PATH'):
        if os.environ.get(v):
            return os.environ[v]
    try:
        cfg = json.load(open(os.path.expanduser('~/.claude.json'), encoding='utf-8'))
        env = cfg.get('mcpServers', {}).get('google-docs', {}).get('env', {})
        for v in ('GOOGLE_APPLICATION_CREDENTIALS', 'SERVICE_ACCOUNT_PATH'):
            if env.get(v):
                return os.path.expanduser(env[v])
    except (OSError, ValueError):
        pass
    sys.exit('키 파일 경로를 찾지 못했습니다. GOOGLE_APPLICATION_CREDENTIALS를 지정하세요.')


def services():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        key_path(), scopes=['https://www.googleapis.com/auth/drive',
                            'https://www.googleapis.com/auth/documents'])
    return (build('drive', 'v3', credentials=creds, cache_discovery=False),
            build('docs', 'v1', credentials=creds, cache_discovery=False))


def docx_counts(path):
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8')
    return {'표': x.count('<w:tbl>'), '그림': x.count('<a:blip ')}


def walk(elements, acc):
    for el in elements or []:
        if 'table' in el:
            acc['표'] += 1
            for row in el['table'].get('tableRows', []):
                for cell in row.get('tableCells', []):
                    walk(cell.get('content'), acc)
        elif 'paragraph' in el:
            for pe in el['paragraph'].get('elements', []):
                if 'inlineObjectElement' in pe:
                    acc['그림'] += 1
                tr = pe.get('textRun')
                if tr and tr.get('content', '').strip():
                    f = tr.get('textStyle', {}).get('weightedFontFamily', {}).get('fontFamily', '(기본)')
                    acc['글꼴'][f] += len(tr['content'])


def report(docs, doc_id):
    d = docs.documents().get(documentId=doc_id).execute()
    acc = {'표': 0, '그림': 0, '글꼴': collections.Counter()}
    walk(d['body']['content'], acc)
    st = d.get('documentStyle', {})
    cm = lambda k: round(st.get(k, {}).get('magnitude', 0) / 28.3465, 2)
    print('제목:', d.get('title'))
    print('표:', acc['표'], '/ 그림:', acc['그림'])
    print('글꼴(글자 수):', dict(acc['글꼴'].most_common(5)))
    print('여백 cm 위·아래·좌·우:', cm('marginTop'), cm('marginBottom'), cm('marginLeft'), cm('marginRight'))
    print('링크: https://docs.google.com/document/d/%s/edit' % doc_id)
    return acc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx', nargs='?')
    ap.add_argument('--doc', default=DOC_ID)
    ap.add_argument('--check', action='store_true')
    a = ap.parse_args()
    drive, docs = services()
    if a.check or not a.docx:
        report(docs, a.doc)
        return
    from googleapiclient.http import MediaFileUpload
    title = os.path.splitext(os.path.basename(a.docx))[0]
    res = drive.files().update(
        fileId=a.doc, body={'name': title},
        media_body=MediaFileUpload(a.docx, mimetype=DOCX_MIME, resumable=False),
        supportsAllDrives=True, fields='id,name,mimeType,modifiedTime').execute()
    print('UPDATED', res['name'], res['mimeType'], res['modifiedTime'])
    want = docx_counts(a.docx)
    got = report(docs, a.doc)
    for k in ('표', '그림'):
        mark = '같음' if want[k] == got[k] else '다름(확인 필요)'
        print('docx %s %d vs 구글 %d: %s' % (k, want[k], got[k], mark))


if __name__ == '__main__':
    main()
