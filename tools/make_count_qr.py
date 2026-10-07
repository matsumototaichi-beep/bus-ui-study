# -*- coding: utf-8 -*-
"""計数研究（車内を1から数える vs 乗った人・降りた人を数える）の QR を作る

    python tools/make_count_qr.py

★乗降人数計算/ に書き出す。
  QRポスター_計数研究.pdf … A4縦×2枚（ゆき＝人数アプリ / かえり＝乗降アプリ）
  乗降アプリQR.png        … チャットで送る用
"""
import os, io, pathlib
import segno
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(os.path.dirname(ROOT), "★乗降人数計算")
OUT = os.path.join(OUTDIR, "QRポスター_計数研究.pdf")
PNG = os.path.join(OUTDIR, "乗降アプリQR.png")
FONTDIR = pathlib.Path(os.environ['TEMP']) / 'jpfonts_countqr'
FONTDIR.mkdir(parents=True, exist_ok=True)

BASE = "https://matsumototaichi-beep.github.io/bus-ui-study/"
POSTERS = [
    # (右上の見分け, アプリ名, URL, 開いたらすること)
    ("計数研究　ゆき", "人数アプリ", BASE + "?study=count", [
        "「ログイン」タブを選ぶ（「新規登録」ではありません）",
        "カードの ID と パスワード を入力",
        "ログインできたら実験者に見せてください",
    ]),
    ("計数研究　かえり", "乗降アプリ", BASE + "counter.html", [
        "ログインは要りません",
        "画面の指示に従って操作してください",
    ]),
]
TITLE = "このQRを読み取ってください"
SUB = "スマホのカメラアプリで読み取ってください"
H2 = "開いたらすること"
WARN1 = "※ LINE や Discord の中のブラウザでは開かないでください。"
WARN2 = "　 カメラアプリから開いてください。"
W, H = 595.0, 842.0   # A4


def build_font(ttc_path, out_name, chars):
    out = FONTDIR / out_name
    tmp = FONTDIR / ('_full_' + out_name)
    if not tmp.exists():
        TTCollection(ttc_path).fonts[0].save(str(tmp))
    opts = subset.Options()
    opts.drop_tables += ['DSIG']
    opts.notdef_outline = True
    font = subset.load_font(str(tmp), opts)
    ss = subset.Subsetter(options=opts)
    ss.populate(unicodes={ord(c) for c in chars if ord(c) > 31})
    ss.subset(font)
    subset.save_font(font, str(out), opts)
    font.close()
    return str(out)


used = set(TITLE + SUB + H2 + WARN1 + WARN2)
for mark, name, url, steps in POSTERS:
    used |= set(mark + name + url + "".join(steps))
used |= set(chr(c) for c in range(0x20, 0x7f))
REG = build_font(r'C:\Windows\Fonts\YuGothM.ttc', 'reg.ttf', used)
BLD = build_font(r'C:\Windows\Fonts\YuGothB.ttc', 'bld.ttf', used)
FR = pymupdf.Font(fontfile=REG)
FB = pymupdf.Font(fontfile=BLD)


def wid(font, text, size):
    return font.text_length(text, fontsize=size)


def qr_png(url, scale=20):
    buf = io.BytesIO()
    segno.make(url, error="m").save(buf, kind="png", scale=scale, border=2)
    return buf.getvalue()


os.makedirs(OUTDIR, exist_ok=True)
doc = pymupdf.open()
for mark, name, url, steps in POSTERS:
    page = doc.new_page(width=W, height=H)
    page.insert_font(fontname="R", fontfile=REG)
    page.insert_font(fontname="B", fontfile=BLD)
    page.insert_text((W - 40 - wid(FB, mark, 20), 54), mark, fontname="B", fontsize=20, color=(0.2, 0.2, 0.2))
    page.insert_text(((W - wid(FB, TITLE, 26)) / 2, 118), TITLE, fontname="B", fontsize=26)
    page.insert_text(((W - wid(FR, SUB, 12)) / 2, 144), SUB, fontname="R", fontsize=12, color=(0.35, 0.35, 0.35))
    side = 330.0
    page.insert_image(pymupdf.Rect((W - side) / 2, 166, (W + side) / 2, 166 + side), stream=qr_png(url))
    page.insert_text(((W - wid(FB, name, 16)) / 2, 520), name, fontname="B", fontsize=16)
    page.insert_text(((W - wid(FR, url, 8.5)) / 2, 538), url, fontname="R", fontsize=8.5, color=(0.4, 0.4, 0.4))
    page.draw_rect(pymupdf.Rect(70, 562, W - 70, 762), color=(0.75, 0.75, 0.75), width=0.8)
    page.insert_text((92, 594), H2, fontname="B", fontsize=15)
    y = 626
    for i, t in enumerate(steps, start=1):
        page.insert_text((92, y), "%d." % i, fontname="B", fontsize=12, color=(0.1, 0.3, 0.6))
        page.insert_text((112, y), t, fontname="R", fontsize=12)
        y += 28
    page.insert_text((92, 740), WARN1, fontname="R", fontsize=10, color=(0.65, 0, 0))
    page.insert_text((92, 754), WARN2, fontname="R", fontsize=10, color=(0.65, 0, 0))
doc.save(OUT, garbage=4, deflate=True)
with open(PNG, "wb") as f:
    f.write(qr_png(BASE + "counter.html", scale=12))
print("書き出しました: %s（%dページ）" % (OUT, doc.page_count))
print("書き出しました: %s" % PNG)
