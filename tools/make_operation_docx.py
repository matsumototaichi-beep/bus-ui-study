# -*- coding: utf-8 -*-
"""実験の進め方を Word にする

    python tools/make_operation_docx.py

docs/実験の進め方.docx が出る。** で挟んだところが太字になる。
"""
import os
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Mm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "実験の進め方.docx")
FONT = "Yu Gothic"


def jp(run, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if color:
        run.font.color.rgb = RGBColor(*color)


def p(doc, text="", size=10, bold=False, before=0, after=4):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = Pt(after)
    for i, chunk in enumerate(text.split("**")):
        if chunk:
            jp(par.add_run(chunk), size, bold or i % 2 == 1)
    return par


def h1(doc, text):
    p(doc, text, size=17, bold=True, after=10)


def h2(doc, text):
    par = p(doc, text, size=13, bold=True, before=14, after=4)
    bd = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "8")
    bot.set(qn("w:space"), "2"); bot.set(qn("w:color"), "1A4B8C")
    bd.append(bot)
    par._p.get_or_add_pPr().append(bd)


def h3(doc, text):
    p(doc, text, size=11, bold=True, before=8, after=3)


def table(doc, head, body, widths=None):
    t = doc.add_table(rows=0, cols=len(head or body[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    for ri, row in enumerate(([head] if head else []) + body):
        cells = t.add_row().cells
        for ci, val in enumerate(row):
            if widths:
                cells[ci].width = Mm(widths[ci])
            par = cells[ci].paragraphs[0]
            par.paragraph_format.space_before = Pt(2)
            par.paragraph_format.space_after = Pt(2)
            is_head = head and ri == 0
            for i, chunk in enumerate(str(val).split("**")):
                if chunk:
                    jp(par.add_run(chunk), 9.5, is_head or i % 2 == 1)
    p(doc, "", size=4, after=0)
    return t


def steps(doc, items):
    for i, t in enumerate(items, start=1):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(3)
        par.paragraph_format.left_indent = Mm(5)
        jp(par.add_run("%d.  " % i), 10, True, (0x1A, 0x4B, 0x8C))
        for j, chunk in enumerate(t.split("**")):
            if chunk:
                jp(par.add_run(chunk), 10, j % 2 == 1)


doc = Document()
st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
for s in doc.sections:
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.top_margin = s.bottom_margin = Mm(17)
    s.left_margin = s.right_margin = Mm(17)

h1(doc, "実験の進め方")

# ---------------------------------------------------------------- 全体
h2(doc, "研究は2つある")
table(doc, ["", "調べること", "1人あたり", "アンケート"],
      [["**UI研究**", "どの入れ方だと答えやすいか", "**2回**", "毎回あり"],
       ["**計数研究**", "人数をどう数えるのが正確か", "**1回**", "なし"]],
      widths=[26, 72, 22, 36])
p(doc, "別の研究として数える。同じ人が両方に出てもよい。")

# ---------------------------------------------------------------- UI研究
h2(doc, "UI研究")
p(doc, "1回＝ＪＲ吹田駅（南口）からの**往復2本**。")
table(doc, ["", "ゆき", "かえり", "わからないボタン", "使うポスター"],
      [["**1回目**", "**数字キー**", "**＋−ボタン**", "なし", "ゆき D ／ かえり B"],
       ["**2回目**", "**＋−ボタン**", "**数字キー**", "**あり**", "ゆき A ／ かえり C"]],
      widths=[20, 30, 30, 34, 42])

h3(doc, "乗ってもらい方")
steps(doc, [
    "1回目と2回目は、**昼と夕方で1回ずつ**乗ってもらう",
    "日付は別の日でよい。**順番も自由**（夕方が先でもよい）",
    "都合がつかなければ同じ時間帯が2回でもよい",
    "往復が終わったら、その場でアンケートに答えてもらう",
])

h3(doc, "使う便")
table(doc, ["", "集合", "ゆき 南口発", "桃山台着", "かえり 桃山台発", "南口着・解散"],
      [["**昼**", "11:35", "**12:03**", "12:48", "**13:48**", "14:32"],
       ["**夕方**", "15:25", "**15:55**", "16:40", "**17:18**", "18:03"]],
      widths=[18, 20, 34, 24, 40, 32])
p(doc, "すべて阪急バス 吹田市内線2系統。南口のりば2からは **[2]** と **[3]** の両方が"
       "「桃山台駅ゆき」で出る。**乗るのは [2]。**")

h3(doc, "QRポスター")
table(doc, ["記号", "入れ方", "わからないボタン"],
      [["**A**", "＋−ボタン", "あり"],
       ["**B**", "＋−ボタン", "なし"],
       ["**C**", "数字キー", "あり"],
       ["**D**", "数字キー", "なし"]],
      widths=[18, 44, 44])

# ---------------------------------------------------------------- 計数研究
h2(doc, "計数研究")
p(doc, "実験日程の**最後のほう**に行う。1回＝往復2本。")
table(doc, ["", "やってもらうこと", "使うアプリ"],
      [["**ゆき**", "停まるたびに**車内の人数を1から数える**", "人数アプリ"],
       ["**かえり**", "停まるたびに**乗った人・降りた人を数える**", "乗降アプリ"]],
      widths=[20, 92, 44])
p(doc, "参加回数は UI研究 とは別に数える。この人にとっての1回目になる。")

# ---------------------------------------------------------------- 乗降アプリ
h2(doc, "乗降アプリの使い方")
p(doc, "https://matsumototaichi-beep.github.io/bus-ui-study/counter.html")
p(doc, "**ログインは要らない。**記録は押すたびに端末へ保存される。電波も要らない。")
steps(doc, [
    "役割を選ぶ（**乗車係 ／ 降車係 ／ 数え直し係**）",
    "方向を選ぶ",
    "便を書く（発車時刻など）",
    "「はじめる」を押す",
    "停まるたびに、**タップ1回＝1人**",
    "押しすぎたら「1つ戻す」",
    "終わったら記録を書き出して松本に送る",
])

h3(doc, "計数役の配置")
table(doc, ["係", "人数", "すること"],
      [["**乗車係**", "1", "その停留所で乗ってきた人を数える"],
       ["**降車係**", "1", "その停留所で降りた人を数える"],
       ["**数え直し係**", "1", "発車したあと、車内にいる人を最初から数える"]],
      widths=[28, 16, 112])

# ---------------------------------------------------------------- 持ちもの
h2(doc, "持っていくもの")
table(doc, None,
      [["配布カード", "Aさん〜Lさん。ID と パスワード"],
       ["QRポスター", "A・B・C・D の4枚。その日に使うのは2枚"],
       ["参加者向け案内", "集合時に配る"],
       ["対応表", "accounts.csv。人目に触れない場所に"],
       ["記録用紙", "乗降アプリが使えないときの予備"]],
      widths=[32, 124])

doc.save(OUT)
print("書き出しました: " + OUT)
