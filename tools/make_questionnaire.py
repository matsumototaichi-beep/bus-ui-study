# -*- coding: utf-8 -*-
"""事後アンケートの紙の用紙（Word）を作る（2026-10-06 改訂）

★設問は docs/questionnaire/設問一覧.xlsx が唯一の置き場所。
  Googleフォーム（tools/make_forms.py）も同じ Excel から作る。
  **文言を直すときは Excel だけを直せばよい。二重管理にならない。**

    python tools/make_questionnaire.py

  docs/questionnaire/ に3ファイル：
    アンケート_前半.docx        … 各実施日の解散直後。全員・版の違いなし
    アンケート_後半_A.docx      … 前半を回収してから渡す。v1・v2 の人
    アンケート_後半_B.docx      … 同上。v3・v4 の人

  紙はGoogleフォームが使えないときの予備。普段はフォームで回答してもらう。

★なぜ前半と後半で分かれているか
  後半の S4 で「わからない」ボタンの話を先に見せると、
  前半の S3（今日どう数えたか）が、ボタンを意識した答えになってしまう。
  **前半を回収してから後半を渡す。** 印刷も別々にする。
"""
import os, sys
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(ROOT, "docs", "questionnaire")
XLSX = os.path.join(OUTDIR, "設問一覧.xlsx")
FONT = "Yu Gothic"
NCOL = 9

BUTTONS = [
    ("「わからない」ボタン あり", "「わからない」というボタンが置かれていた便"),
    ("「わからない」ボタン なし", "同じ場所に何も置かれていなかった便"),
]


# ---------------------------------------------------------------- 書式の道具
def setup(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for s in doc.sections:
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.top_margin = s.bottom_margin = Mm(18)
        s.left_margin = s.right_margin = Mm(18)


def p(doc, text="", size=10.5, bold=False, color=None, before=0, after=3, align=None):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = Pt(after)
    if align:
        par.alignment = align
    run = par.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return par


def box(doc, label, desc, label_pt=11):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.cell(0, 0).width = Mm(174)
    cp = t.cell(0, 0).paragraphs[0]
    r1 = cp.add_run("【%s】　" % label); r1.bold = True; r1.font.size = Pt(label_pt)
    r2 = cp.add_run(desc); r2.font.size = Pt(10)
    for r in (r1, r2):
        r.font.name = FONT
        r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    p(doc, "", size=4, after=2)


def lines(doc, n=2):
    for _ in range(n):
        p(doc, "　" + "＿" * 44, size=10, color=(0x99, 0x99, 0x99), after=5)


# ---------------------------------------------------------------- 設問を1つ描く
def emit(doc, row, button_order=None):
    youshi, setsu, no, kata, bun, sel, must, hosoku, kurikaeshi = row
    choices = [x.strip() for x in (sel or "").split("／") if x.strip()]
    k = (kurikaeshi or "").strip()

    if k == "ボタンごと":
        p(doc, bun + ("　" + hosoku if hosoku else ""), size=9.5, before=4, after=3)
        for nm, d in (button_order or BUTTONS):
            box(doc, nm, d)
        return
    if k == "ボタン順":
        names = [b[0] for b in (button_order or BUTTONS)]
        vals = ["ボタンがあった便" if "あり" in n else "ボタンがなかった便" for n in names]
        p(doc, "%s  %s" % (no, bun), before=4, after=1)
        p(doc, "　" + "　　".join("□ " + c for c in vals + ["変わらない"]), size=10, after=3)
        return

    if kata == "ページ":
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        p(doc, bun, size=13, bold=True, before=4, after=2)
        if hosoku:
            p(doc, hosoku, size=8.5, color=(0x55, 0x55, 0x55), after=5)
        return
    if kata == "説明":
        p(doc, bun, size=13, bold=True, before=10, after=2)
        if hosoku:
            p(doc, hosoku, size=8.5, color=(0x55, 0x55, 0x55), after=5)
        return

    title = ("%s  %s" % (no, bun)).strip() if no else bun
    p(doc, title, bold=(must == "○"), before=4, after=1)
    if hosoku:
        p(doc, "　" + hosoku, size=8.5, color=(0x55, 0x55, 0x55), after=2)

    if kata in ("単一選択", "複数選択"):
        p(doc, "　" + "　　".join("□ " + c for c in choices), size=10, after=3)
    elif kata == "5段階":
        left = choices[0] if choices else ""
        right = choices[1] if len(choices) > 1 else ""
        p(doc, "　　1 ・ 2 ・ 3 ・ 4 ・ 5　（1＝%s … 5＝%s）" % (left, right), size=10, after=3)
    elif kata == "日付":
        p(doc, "　　＿＿月＿＿日", size=10, after=4)
    elif kata == "短文":
        lines(doc, 1)
    elif kata == "段落":
        lines(doc, 2)


# ---------------------------------------------------------------- 用紙を1つ作る
def build(rows, title, lead, button_order=None, tail=None):
    doc = Document()
    setup(doc)
    p(doc, title, size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    p(doc, lead, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=(0x55, 0x55, 0x55), after=8)
    for row in rows:
        emit(doc, row, button_order)
    if tail:
        p(doc, tail, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER,
          color=(0x55, 0x55, 0x55), before=10)
    p(doc, "ご協力ありがとうございました。", size=11, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, before=8)
    return doc


def main():
    if not os.path.exists(XLSX):
        sys.exit("%s がありません。先に python tools/make_forms.py --bank を流してください。" % XLSX)
    ws = load_workbook(XLSX)["設問"]
    rows = []
    for r in range(6, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, NCOL + 1)]
        if not any(vals):
            continue
        rows.append([("" if v is None else str(v)) for v in vals])

    front = [r for r in rows if r[0] == "①前半"]
    back = [r for r in rows if r[0] == "①後半"]

    out = []
    d = build(front, "バス車内の人数を答えるアプリについて　― 今日の分（前半）",
              "今日の乗車についてお答えください。所要7分ほどです。",
              tail="ここまでを実験者にお渡しください。続きの用紙をお渡しします。")
    path = os.path.join(OUTDIR, "アンケート_前半.docx")
    d.save(path); out.append(path)

    for name, order in (("A", BUTTONS), ("B", list(reversed(BUTTONS)))):
        d = build(back, "バス車内の人数を答えるアプリについて　― 今日の分（後半）%s" % name,
                  "前半をお渡しいただいた方にお配りしています。引き続きお答えください。",
                  button_order=order)
        path = os.path.join(OUTDIR, "アンケート_後半_%s.docx" % name)
        d.save(path); out.append(path)

    for x in out:
        print("書き出しました: " + x)
    print("\n前半 … 各実施日の解散直後。全員・版の違いなし")
    print("後半A … v1・v2 の人　／　後半B … v3・v4 の人")
    print("★前半を回収してから後半を渡すこと。印刷も分けること。")


if __name__ == "__main__":
    main()
