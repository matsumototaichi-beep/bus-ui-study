# -*- coding: utf-8 -*-
"""紙のアンケート用紙を作る（Googleフォームが使えないときの予備）

    python tools/make_questionnaire.py

設問の正本は docs/questionnaire/設問一覧.xlsx。
docs/questionnaire/ に アンケート_1回目.docx と アンケート_2回目.docx が出る。
"""
import os, sys
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(ROOT, "docs", "questionnaire")
XLSX = os.path.join(OUTDIR, "設問一覧.xlsx")
FONT = "Yu Gothic"
NCOL = 8


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


def lines(doc, n=2):
    for _ in range(n):
        p(doc, "　" + "＿" * 44, size=10, color=(0x99, 0x99, 0x99), after=5)


def emit(doc, row, first):
    youshi, setsu, no, kata, bun, sel, must, hosoku = row
    choices = [x.strip() for x in (sel or "").split("／") if x.strip()]

    if kata in ("ページ", "説明"):
        if kata == "ページ" and not first[0]:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        first[0] = False
        p(doc, bun, size=13, bold=True, before=4, after=2)
        if hosoku:
            p(doc, hosoku, size=9, color=(0x55, 0x55, 0x55), after=6)
        return

    first[0] = False
    title = ("%s  %s" % (no, bun)).strip() if no else bun
    p(doc, title, bold=(must == "○"), before=4, after=1)
    if hosoku:
        p(doc, "　" + hosoku, size=9, color=(0x55, 0x55, 0x55), after=2)

    if kata in ("単一選択", "複数選択"):
        p(doc, "　" + "　　".join("□ " + c for c in choices), size=10, after=3)
    elif kata == "5段階":
        left = choices[0] if choices else ""
        right = choices[1] if len(choices) > 1 else ""
        p(doc, "　　1 ・ 2 ・ 3 ・ 4 ・ 5　（1＝%s　5＝%s）" % (left, right), size=10, after=3)
    elif kata == "日付":
        p(doc, "　　＿＿月＿＿日", size=10, after=4)
    elif kata == "短文":
        lines(doc, 1)
    elif kata == "段落":
        lines(doc, 2)


def build(rows, title):
    doc = Document()
    setup(doc)
    p(doc, title, size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
    first = [True]
    for row in rows:
        emit(doc, row, first)
    p(doc, "ご協力ありがとうございました。", size=11, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, before=10)
    return doc


def main():
    if not os.path.exists(XLSX):
        sys.exit("%s がありません。先に python tools/make_forms.py --bank を流してください。" % XLSX)
    ws = load_workbook(XLSX)["設問"]
    rows = []
    for r in range(5, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, NCOL + 1)]
        if not any(vals):
            continue
        rows.append([("" if v is None else str(v)) for v in vals])

    for youshi, fname in (("1回目", "アンケート_1回目.docx"), ("2回目", "アンケート_2回目.docx")):
        sub = [r for r in rows if r[0] == youshi]
        path = os.path.join(OUTDIR, fname)
        build(sub, "バスの人数アプリ　%sのアンケート" % youshi).save(path)
        print("書き出しました: %s（%d問）" % (path, len([r for r in sub if r[2]])))


if __name__ == "__main__":
    main()
