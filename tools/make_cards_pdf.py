# -*- coding: utf-8 -*-
"""配布カードをA4のPDFにする（印刷を人に頼むとき用）

    python tools/make_cards_pdf.py

★UI研究/accounts.csv を読んで、★UI研究/配布カード.pdf を作る。
A4に8枚（2列×4行）。切り取り線つき。
"""
import os, csv, pathlib
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(os.path.dirname(ROOT), "★UI研究")
CSV = os.path.join(OUTDIR, "accounts.csv")
OUT = os.path.join(OUTDIR, "配布カード.pdf")
FONTDIR = pathlib.Path(os.environ['TEMP']) / 'jpfonts_cards2'
FONTDIR.mkdir(parents=True, exist_ok=True)

NOTE1 = "このIDとパスワードは2回目にもう一度使います。"
NOTE2 = "カードを無くさないでください。"

# A4 = 595 x 842pt
M = 28                      # 外側の余白
GAP = 12                    # カードどうしの間
COLS, ROWS = 2, 4
CW = (595 - M * 2 - GAP) / COLS
CH = (842 - M * 2 - GAP * (ROWS - 1)) / ROWS


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


with open(CSV, encoding="utf-8-sig", newline="") as f:
    rows = list(csv.DictReader(f))

used = set("ID パスワード 予備" + NOTE1 + NOTE2)
used |= set(chr(c) for c in range(0x20, 0x7f))
for r in rows:
    used |= set(r["名前"]) | set(r["ID"]) | set(r["パスワード"])

REG = build_font(r'C:\Windows\Fonts\YuGothM.ttc', 'reg.ttf', used)
BLD = build_font(r'C:\Windows\Fonts\YuGothB.ttc', 'bld.ttf', used)

doc = pymupdf.open()
for i, r in enumerate(rows):
    if i % (COLS * ROWS) == 0:
        page = doc.new_page(width=595, height=842)
        page.insert_font(fontname="R", fontfile=REG)
        page.insert_font(fontname="B", fontfile=BLD)

    k = i % (COLS * ROWS)
    x = M + (k % COLS) * (CW + GAP)
    y = M + (k // COLS) * (CH + GAP)

    # 枠（切り取り線）
    page.draw_rect(pymupdf.Rect(x, y, x + CW, y + CH),
                   color=(0.55, 0.55, 0.55), width=0.7, dashes="[3 3] 0")

    yobi = (r.get("予備") or "").strip()

    # 名前
    page.insert_text((x + 16, y + 34), r["名前"], fontname="B", fontsize=22)
    if yobi:
        page.insert_text((x + CW - 46, y + 24), "予備", fontname="B", fontsize=10,
                         color=(0.7, 0, 0))

    # 区切り線
    page.draw_line(pymupdf.Point(x + 16, y + 46), pymupdf.Point(x + CW - 16, y + 46),
                   color=(0.8, 0.8, 0.8), width=0.6)

    # ID / パスワード
    page.insert_text((x + 16, y + 72), "ID", fontname="R", fontsize=9, color=(0.4, 0.4, 0.4))
    page.insert_text((x + 16, y + 92), r["ID"], fontname="B", fontsize=19)

    page.insert_text((x + 16, y + 120), "パスワード", fontname="R", fontsize=9, color=(0.4, 0.4, 0.4))
    page.insert_text((x + 16, y + 140), r["パスワード"], fontname="B", fontsize=19)

    # 注意書き
    page.insert_text((x + 16, y + CH - 26), NOTE1, fontname="R", fontsize=8, color=(0.3, 0.3, 0.3))
    page.insert_text((x + 16, y + CH - 14), NOTE2, fontname="R", fontsize=8, color=(0.3, 0.3, 0.3))

doc.save(OUT, garbage=4, deflate=True)
print("書き出しました: %s" % OUT)
print("  %d人分 ／ %dページ ／ %d バイト" % (len(rows), doc.page_count, os.path.getsize(OUT)))
