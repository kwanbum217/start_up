"""피치덱 pptx를 검토용 구글 슬라이드(같은 링크)에 변환해 넣고, 결과를 확인한다.

사용:
  python3 -W ignore sync_gslides.py --check                 # 읽기만: 제목·장 수·글꼴
  python3 -W ignore sync_gslides.py <피치덱.pptx>            # 같은 파일 ID에 pptx를 덮어 변환한 뒤 확인
  python3 -W ignore sync_gslides.py <피치덱.pptx> --deck <ID>

키 파일 경로는 sync_gdoc.py와 같다. 키 내용은 출력하지 않는다.
대상 슬라이드는 그 키의 client_email에 편집자로 공유돼 있어야 한다.
제출 pptx는 바꾸지 않는다. 올리는 임시 사본만 Malgun Gothic·맑은 고딕을 Nanum Gothic으로 바꾼다.
덮어쓰면 이전 장은 구글 슬라이드의 버전 기록에만 남는다.
절차의 정본은 references/gslides.md다.
"""
import argparse, collections, os, re, sys, tempfile, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sync_gdoc as g

DECK_ID = '1freHyZ0yO83FKwQ_UziZXhc3czPzsfHGHdEJlUq4Nvw'
PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
SRC_FONTS = ('Malgun Gothic', '맑은 고딕')
OK_FONTS = ('Nanum Gothic', '나눔고딕')


def services():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        g.key_path(), scopes=['https://www.googleapis.com/auth/drive',
                              'https://www.googleapis.com/auth/presentations'])
    return (build('drive', 'v3', credentials=creds, cache_discovery=False),
            build('slides', 'v1', credentials=creds, cache_discovery=False))


def pptx_meta(path):
    slides, charts = [], []
    fonts = collections.Counter()
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.fullmatch(r'ppt/slides/slide\d+\.xml', name):
                slides.append(name)
            elif name.startswith('ppt/charts/') and not name.endswith('/'):
                charts.append(name)
            if name.endswith('.xml'):
                text = z.read(name).decode('utf-8', 'replace')
                for face in re.findall(r'typeface="([^"]+)"', text):
                    if face in SRC_FONTS or face in OK_FONTS:
                        fonts[face] += 1
    return {'장': len(slides), '차트': len(charts), '글꼴': fonts}


def variant(pptx):
    """한글 글꼴만 Nanum Gothic으로 바꾼 임시 사본. 원본 파일은 열기만 한다."""
    out = os.path.join(tempfile.mkdtemp(), os.path.basename(pptx))
    replaced = 0
    with zipfile.ZipFile(pptx) as zi, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename.endswith('.xml'):
                text = data.decode('utf-8')
                for src in SRC_FONTS:
                    old = 'typeface="%s"' % src
                    new = 'typeface="Nanum Gothic"'
                    replaced += text.count(old)
                    text = text.replace(old, new)
                data = text.encode('utf-8')
            zo.writestr(item.filename, data)
    return out, replaced


def font_families(obj, acc):
    if isinstance(obj, dict):
        face = obj.get('fontFamily')
        if isinstance(face, str) and face:
            acc[face] += 1
        for value in obj.values():
            font_families(value, acc)
    elif isinstance(obj, list):
        for value in obj:
            font_families(value, acc)


def report(slides, deck_id):
    d = slides.presentations().get(presentationId=deck_id).execute()
    fonts = collections.Counter()
    font_families(d, fonts)
    print('제목:', d.get('title'))
    print('장:', len(d.get('slides', [])))
    print('글꼴(텍스트 런 수):', dict(fonts.most_common(8)))
    print('링크: https://docs.google.com/presentation/d/%s/edit' % deck_id)
    return {'장': len(d.get('slides', [])), '글꼴': fonts}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx', nargs='?')
    ap.add_argument('--deck', default=DECK_ID)
    ap.add_argument('--check', action='store_true')
    a = ap.parse_args()
    drive, slides = services()
    if a.check or not a.pptx:
        report(slides, a.deck)
        return
    src = pptx_meta(a.pptx)
    copy, replaced = variant(a.pptx)
    copied = pptx_meta(copy)
    left = sum(copied['글꼴'][name] for name in SRC_FONTS)
    print('임시 사본 글꼴 바꿈:', replaced, '곳 / 남은 맑은 고딕:', left, '/ 차트 부품:', src['차트'])
    if left:
        sys.exit('임시 사본에 맑은 고딕이 남았습니다. 제출 pptx는 그대로 두고 사본만 확인하세요.')
    from googleapiclient.http import MediaFileUpload
    title = os.path.splitext(os.path.basename(a.pptx))[0]
    res = drive.files().update(
        fileId=a.deck, body={'name': title},
        media_body=MediaFileUpload(copy, mimetype=PPTX_MIME, resumable=False),
        supportsAllDrives=True, fields='id,name,mimeType,modifiedTime').execute()
    print('UPDATED', res['name'], res['mimeType'], res['modifiedTime'])
    got = report(slides, a.deck)
    mark = '같음' if src['장'] == got['장'] else '다름(확인 필요)'
    print('pptx 장 %d vs 구글 %d: %s' % (src['장'], got['장'], mark))
    bad = sum(got['글꼴'][name] for name in SRC_FONTS)
    good = sum(got['글꼴'][name] for name in OK_FONTS)
    font_mark = '같음' if bad == 0 and good else '다름(확인 필요)'
    print('구글 한글 글꼴 Nanum Gothic/나눔고딕 %d, 맑은 고딕 %d: %s' % (good, bad, font_mark))


if __name__ == '__main__':
    main()
