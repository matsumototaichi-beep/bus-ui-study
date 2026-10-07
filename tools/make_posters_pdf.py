# -*- coding: utf-8 -*-
"""QRポスターをA4のPDFにする（印刷を人に頼むとき用）

    pip install segno
    python tools/make_posters_pdf.py

★UI研究/QRポスター.pdf を作る。A4縦×3枚（1回目 / 2回目 / 計数研究）。
docs/qr_posters.html と同じ内容。
"""
import os, io, pathlib
import segno
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(os.path.dirname(ROOT), "★UI研究")
OUT = os.path.join(OUTDIR, "QRポスター.pdf")
FONTDIR = pathlib.Path(os.environ['TEMP']) / 'jpfonts_posters'
FONTDIR.mkdir(parents=True, exist_ok=True)

BASE = "https://matsumototaichi-beep.github.io/bus-ui-study/"

# ★ポスターに条件の中身は書かない。参加者が条件に気づくと反応が混ざるため。
# ★2026-10-07: 4枚→3枚。ゆき/かえりはアプリ内で切り替える（実験者は南口で離脱し、桃山台にQRを出す人がいないため）。
# 入力方式・「わからない」ボタンはアプリが (回目, ゆき/かえり, 組) から決める。
POSTERS = [
    # (右上の見分け, URL, 最後の手順)
    ("1回目",    "?round=1",     "ログインしたら「ゆき」を押して、実験者に見せてください"),
    ("2回目",    "?round=2",     "ログインしたら「ゆき」を押して、実験者に見せてください"),
    ("計数研究", "?study=count", "ログインできたら実験者に見せてください"),
]

TITLE = "このQRを読み取ってください"
SUB = "スマホのカメラアプリで読み取ってください"
H2 = "開いたらすること"
STEPS = [
    "「ログイン」タブを選ぶ（「新規登録」ではありません）",
    "カードの ID と パスワード を入力",
    "うまくいかないときは「パスワードを表示して確かめる」にチェック",
]
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


used = set(TITLE + SUB + H2 + WARN1 + WARN2 + "".join(STEPS) + BASE)
for mark, qs, last in POSTERS:
    used |= set(mark) | set(qs) | set(last)
used |= set(chr(c) for c in range(0x20, 0x7f)) | set("1234567890.")

REG = build_font(r'C:\Windows\Fonts\YuGothM.ttc', 'reg.ttf', used)
BLD = build_font(r'C:\Windows\Fonts\YuGothB.ttc', 'bld.ttf', used)

FR = pymupdf.Font(fontfile=REG)   # 文字幅を測って中央に置くため
FB = pymupdf.Font(fontfile=BLD)


def wid(font, text, size):
    return font.text_length(text, fontsize=size)


doc = pymupdf.open()
for mark, qs, last in POSTERS:
    url = BASE + qs
    page = doc.new_page(width=W, height=H)
    page.insert_font(fontname="R", fontfile=REG)
    page.insert_font(fontname="B", fontfile=BLD)

    # 右上の見分け
    page.insert_text((W - 40 - wid(FB, mark, 20), 54),
                     mark, fontname="B", fontsize=20, color=(0.2, 0.2, 0.2))

    # 見出し
    page.insert_text(((W - wid(FB, TITLE, 26)) / 2, 118),
                     TITLE, fontname="B", fontsize=26)
    page.insert_text(((W - wid(FR, SUB, 12)) / 2, 144),
                     SUB, fontname="R", fontsize=12, color=(0.35, 0.35, 0.35))

    # QR
    qr = segno.make(url, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=20, border=2)
    buf.seek(0)
    side = 330.0
    page.insert_image(pymupdf.Rect((W - side) / 2, 166, (W + side) / 2, 166 + side),
                      stream=buf.read())

    # URL
    page.insert_text(((W - wid(FR, url, 8.5)) / 2, 520),
                     url, fontname="R", fontsize=8.5, color=(0.4, 0.4, 0.4))

    # 開いたらすること
    page.draw_rect(pymupdf.Rect(70, 552, W - 70, 762),
                   color=(0.75, 0.75, 0.75), width=0.8)
    page.insert_text((92, 584), H2, fontname="B", fontsize=15)
    y = 616
    for i, t in enumerate(STEPS + [last], start=1):
        page.insert_text((92, y), "%d." % i, fontname="B", fontsize=12,
                         color=(0.1, 0.3, 0.6))
        page.insert_text((112, y), t, fontname="R", fontsize=12)
        y += 28
    page.insert_text((92, 740), WARN1, fontname="R", fontsize=10, color=(0.65, 0, 0))
    page.insert_text((92, 754), WARN2, fontname="R", fontsize=10, color=(0.65, 0, 0))

doc.save(OUT, garbage=4, deflate=True)
print("書き出しました: %s" % OUT)
print("  %dページ ／ %d バイト" % (doc.page_count, os.path.getsize(OUT)))
