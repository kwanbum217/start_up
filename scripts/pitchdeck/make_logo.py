"""BidBox 로고(제품 저장소의 logo.png, 남색 바탕 포함)에서 피치덱용 투명 로고를 만든다.

사용: python3 -I scripts/pitchdeck/make_logo.py [원본 logo.png]
결과: scripts/pitchdeck/assets/
  bidbox_logo_dark.png   어두운 장용. 원래 색(흰 BID, 파란 BOX·상자) 그대로, 바탕만 투명
  bidbox_logo_light.png  밝은 장용. 흰 부분(BID 글자, 체크, 화살표)을 남색 글자색으로 바꿈
원본 바탕색과 거리가 가까울수록 투명하게 해서(부드러운 가장자리), 바탕 섞인 색을 되돌린다.
"""
import os, sys
from PIL import Image

SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser(
    '~/Documents/korea_IT/lanhchain_ai_vision/refac_bid_box/src/app/static/images/logo.png')
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
BG = (2, 33, 84)          # 원본 바탕(가장 많은 색)
INK = (15, 23, 42)        # 피치덱 글자색 #0F172A
LO, HI = 25, 110          # 바탕과의 거리: LO 이하 완전 투명, HI 이상 완전 불투명

im = Image.open(SRC).convert('RGB')
W, H = im.size
src = im.load()
dark = Image.new('RGBA', (W, H), (0, 0, 0, 0)); light = Image.new('RGBA', (W, H), (0, 0, 0, 0))
dp, lp = dark.load(), light.load()
for y in range(H):
    for x in range(W):
        p = src[x, y]
        d = sum((a - b) ** 2 for a, b in zip(p, BG)) ** 0.5
        a = min(1.0, max(0.0, (d - LO) / (HI - LO)))
        if a == 0:
            continue
        # 바탕이 섞인 가장자리 색을 되돌린다
        c = tuple(max(0, min(255, round((v - bg * (1 - a)) / a))) for v, bg in zip(p, BG))
        A = round(a * 255)
        dp[x, y] = c + (A,)
        r, g, b = c
        whiteish = min(c) > 150 and max(c) - min(c) < 70   # 흰색에 가까운 부분(BID, 체크, 화살표)
        lp[x, y] = (INK if whiteish else c) + (A,)
box = dark.getbbox()
pad = 12
box = (max(0, box[0] - pad), max(0, box[1] - pad), min(W, box[2] + pad), min(H, box[3] + pad))
os.makedirs(OUT, exist_ok=True)
for name, img in (('bidbox_logo_dark.png', dark), ('bidbox_logo_light.png', light)):
    img.crop(box).save(os.path.join(OUT, name))
print('crop', box, 'size', (box[2] - box[0], box[3] - box[1]), '->', OUT)
