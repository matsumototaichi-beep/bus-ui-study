# -*- coding: utf-8 -*-
"""紙のアンケート用紙を作る（Googleフォームが使えないときの予備）

    python tools/make_questionnaire.py

設問の正本は docs/questionnaire/設問一覧.xlsx。
docs/questionnaire/ に フォームと同じ4つが出る。
  アンケート_1回目_桃山台.docx / アンケート_1回目_南口.docx / アンケート_2回目_桃山台.docx / アンケート_2回目_南口.docx
"""
import os, sys
sys.dont_write_bytecode = True          # tools/ に .pyc を増やさない
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(ROOT, "docs", "questionnaire")
FONT = "Yu Gothic"
OWARI = "ご協力ありがとうございました。"     # 「回答後」の行がないときの結び


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
    youshi, setsu, no, kata, bun, sel, must, hosoku = row[:8]
    choices = [x.strip() for x in (sel or "").split("／") if x.strip()]

    if kata in ("表紙", "回答後"):
        return                      # 題名・説明文・結びは build() で出す
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
    elif kata in ("短文", "ID"):
        lines(doc, 1)
    elif kata == "段落":
        lines(doc, 2)


def build(rows):
    doc = Document()
    setup(doc)
    # ★2026-10-07: 題名と説明文は Excel の「表紙」の行から取る（Googleフォームと同じ文言にする）
    hyoshi = [r for r in rows if r[3] == "表紙"]
    if not hyoshi:
        sys.exit("%s の「表紙」の行がありません。" % rows[0][0])
    p(doc, hyoshi[0][4], size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    if hyoshi[0][7]:
        p(doc, hyoshi[0][7], size=10, color=(0x55, 0x55, 0x55), align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
    first = [True]
    for row in rows:
        emit(doc, row, first)
    # ★2026-10-08: 結びは「回答後」の行（フォームで送信したあとに出る文）と同じにする
    owari = [r[4] for r in rows if r[3] == "回答後" and r[4]]
    p(doc, owari[0] if owari else OWARI, size=11, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, before=10)
    return doc


def main():
    import make_forms as mf
    rows = mf.read_bank()
    for youshi in mf.YOUSHI:
        sub = [r for r in rows if r[0] == youshi]
        path = os.path.join(OUTDIR, "アンケート_%s.docx" % youshi)
        build(sub).save(path)
        print("書き出しました: %s（%d問）" % (path, mf.count_q(rows, youshi)))


if __name__ == "__main__":
    main()
