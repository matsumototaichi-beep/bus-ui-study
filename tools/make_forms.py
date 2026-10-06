# -*- coding: utf-8 -*-
"""事後アンケートを作る

    python tools/make_forms.py --bank --force   設問一覧.xlsx を作り直す
    python tools/make_forms.py                  Excel から GoogleForms生成.gs を作る

設問の正本は docs/questionnaire/設問一覧.xlsx。文言を直すときはそこを直す。

アンケートは2種類。
  1回目 … ゆき＝数字キー ／ かえり＝＋−ボタン
  2回目 … ゆき＝＋−ボタン ／ かえり＝数字キー ＋ 両方に「わからない」ボタン
"""
import os, sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(ROOT, "docs", "questionnaire")
XLSX = os.path.join(OUTDIR, "設問一覧.xlsx")
GS = os.path.join(OUTDIR, "GoogleForms生成.gs")

COLS = ["用紙", "節", "番号", "形式", "設問文", "選択肢（／区切り）", "必須", "補足（説明文）"]

R1 = "1回目"
R2 = "2回目"

LIKERT = "まったくそう思わない／とてもそう思う"


def hyouka(youshi, prefix, nokori):
    """ゆき／かえりの評価5問。中身は同じで、見出しだけ変わる"""
    return [
        (youshi, "評価", "%s-1" % prefix, "5段階", "全体として答えやすかった", LIKERT, "○", ""),
        (youshi, "評価", "%s-2" % prefix, "5段階", "素早く答えられた", LIKERT, "", ""),
        (youshi, "評価", "%s-3" % prefix, "5段階", "押し間違えにくかった", LIKERT, "", ""),
        (youshi, "評価", "%s-4" % prefix, "5段階", "思ったとおりの人数を入れられた", LIKERT, "", ""),
        (youshi, "評価", "%s-5" % prefix, "5段階", "頭を使う・疲れると感じた", LIKERT, "", ""),
    ]


def kazoekata(youshi):
    """今日の数え方。1回目も2回目も同じ"""
    return [
        (youshi, "数え方", "", "ページ", "今日、人数を数えたときのこと", "", "",
         "正解・不正解を見るものではありません。思ったとおりにお答えください。"),
        (youshi, "数え方", "数-1", "単一選択", "人数を答えるとき、一人ずつ数えていましたか",
         "毎回数えた／だいたい数えた／目分量が多かった／ほとんど目分量", "○", ""),
        (youshi, "数え方", "数-2", "複数選択", "数えにくかったのはどんなときですか（いくつでも）",
         "混んでいた／自分が立っていた／揺れた／降りる人が多かった／暗かった／その他", "", ""),
        (youshi, "数え方", "数-3", "単一選択", "数えきれなかったとき、どうしましたか",
         "だいたいの数を入れて送った／空欄のまま送った／数えきれないことはなかった／覚えていない", "○", ""),
        (youshi, "数え方", "数-4", "単一選択", "座っていましたか、立っていましたか",
         "ほぼ座っていた／半々／ほぼ立っていた", "", ""),
    ]


def kurabete(youshi, yuki, kaeri):
    return [
        (youshi, "くらべて", "", "ページ", "ゆきとかえりをくらべて", "", "", ""),
        (youshi, "くらべて", "比-1", "単一選択", "どちらが答えやすかったですか",
         "ゆきの%s／かえりの%s／変わらない" % (yuki, kaeri), "○", ""),
        (youshi, "くらべて", "比-2", "段落", "そう思った理由があれば教えてください", "", "", ""),
    ]


def saigo(youshi):
    return [
        (youshi, "最後に", "", "ページ", "最後に", "", "", ""),
        (youshi, "最後に", "終-1", "段落", "人数を答えるとき、いちばん面倒だったことは何ですか", "", "", ""),
        (youshi, "最後に", "終-2", "段落", "こうしたら答えやすくなる、という案があれば", "", "", ""),
    ]


SUS = [
    "このアプリを何度も使いたいと思う",
    "このアプリは必要以上に複雑だと思う",
    "このアプリは簡単に使えると思う",
    "このアプリを使うには、詳しい人の助けが必要だと思う",
    "このアプリのいろいろな機能は、うまくまとまっていると思う",
    "このアプリには一貫性のないところが多いと思う",
    "たいていの人は、このアプリの使い方をすぐ覚えられると思う",
    "このアプリはとても使いにくいと思う",
    "このアプリを自信をもって使えた",
    "このアプリを使う前に、いろいろ覚える必要があった",
]

ROWS = (
    # ================================ 1回目 ================================
    [
        (R1, "はじめに", "", "説明", "1回目のアンケート", "", "",
         "ゆきとかえりの2本に乗ったあとにお答えください。5分ほどで終わります。"),
        (R1, "はじめに", "初-1", "短文", "カードに書いてあるID", "", "○", "p01 のような形です"),
        (R1, "はじめに", "初-2", "日付", "今日の日付", "", "○", ""),
        (R1, "はじめに", "初-3", "単一選択", "今日乗ったのはどちらですか", "昼／夕方", "○", ""),

        (R1, "あなた", "", "ページ", "あなたについて", "", "", "1回目だけおうかがいします。"),
        (R1, "あなた", "個-1", "単一選択", "年代", "10代／20代／30代／40代以上", "", ""),
        (R1, "あなた", "個-2", "単一選択", "普段バスに乗る頻度",
         "ほぼ毎日／週に数回／月に数回／ほとんど乗らない", "", ""),
        (R1, "あなた", "個-3", "単一選択", "この路線に乗ったことがありますか",
         "よく乗る／数回ある／初めて", "", ""),
        (R1, "あなた", "個-4", "単一選択", "使ったスマホ", "iPhone／Android／その他", "", ""),
        (R1, "あなた", "個-5", "単一選択", "スマホはどのように持って操作しましたか",
         "片手（親指）／両手／その時による", "", ""),
    ]
    + kazoekata(R1)
    + [(R1, "評価", "", "ページ", "ゆき（数字キー）について", "", "",
        "数字のキーを押して人数を直接入力する画面です。")]
    + hyouka(R1, "ゆ", 0)
    + [(R1, "評価", "", "ページ", "かえり（＋−ボタン）について", "", "",
        "「−5」「−1」「＋1」「＋5」のボタンで数を合わせる画面です。")]
    + hyouka(R1, "か", 0)
    + kurabete(R1, "数字キー", "＋−ボタン")
    + saigo(R1)

    # ================================ 2回目 ================================
    + [
        (R2, "はじめに", "", "説明", "2回目のアンケート", "", "",
         "ゆきとかえりの2本に乗ったあとにお答えください。8分ほどで終わります。"),
        (R2, "はじめに", "初-1", "短文", "カードに書いてあるID", "", "○", "1回目と同じIDを入れてください"),
        (R2, "はじめに", "初-2", "日付", "今日の日付", "", "○", ""),
        (R2, "はじめに", "初-3", "単一選択", "今日乗ったのはどちらですか", "昼／夕方", "○", ""),
    ]
    + kazoekata(R2)
    + [(R2, "評価", "", "ページ", "ゆき（＋−ボタン）について", "", "",
        "「−5」「−1」「＋1」「＋5」のボタンで数を合わせる画面です。")]
    + hyouka(R2, "ゆ", 0)
    + [(R2, "評価", "", "ページ", "かえり（数字キー）について", "", "",
        "数字のキーを押して人数を直接入力する画面です。")]
    + hyouka(R2, "か", 0)
    + kurabete(R2, "＋−ボタン", "数字キー")
    + [
        (R2, "わからない", "", "ページ", "「わからない」ボタンについて", "", "",
         "今日は、人数を入れる欄のすぐ下に「わからない」というボタンがありました。"),
        (R2, "わからない", "分-1", "単一選択", "「わからない」ボタンを押したことがありますか",
         "何度もある／数回ある／一度もない／気づかなかった", "○", ""),
        (R2, "わからない", "分-2", "5段階", "「わからない」ボタンがあったほうが答えやすかった", LIKERT, "○", ""),
        (R2, "わからない", "分-3", "単一選択",
         "ボタンが無かった1回目とくらべて、正直に答えられましたか",
         "今日のほうが正直に答えられた／変わらない／1回目のほうが正直に答えられた", "○", ""),
        (R2, "わからない", "分-4", "段落", "そう思った理由があれば教えてください", "", "", ""),

        (R2, "使いやすさ", "", "ページ", "アプリ全体について", "", "",
         "2回使ってみた全体の印象をお答えください。"),
    ]
    + [(R2, "使いやすさ", "全-%d" % i, "5段階", t, LIKERT, "", "")
       for i, t in enumerate(SUS, start=1)]
    + saigo(R2)
)


# ---------------------------------------------------------------- Excel を作る
def write_bank(force=False):
    if os.path.exists(XLSX) and not force:
        sys.exit("既に %s があります。上書きするなら --force を付けてください。" % XLSX)
    os.makedirs(OUTDIR, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "設問"
    ws["A1"] = "事後アンケートの設問一覧"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "「設問文」「選択肢」「補足」を直してから python tools/make_forms.py を流すと作り直されます。"
    ws["A2"].font = Font(size=9, color="555555")

    thin = Side(style="thin", color="AAAAAA")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, c in enumerate(COLS, start=1):
        cell = ws.cell(row=4, column=i, value=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DCE6F1")
        cell.border = box
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate([9, 11, 7, 10, 48, 44, 6, 38], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r, row in enumerate(ROWS, start=5):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = box
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (5, 6, 8)))
        if row[3] in ("説明", "ページ"):
            for c in range(1, len(COLS) + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="F2F2F2")

    dv = DataValidation(type="list",
                        formula1='"説明,ページ,短文,段落,単一選択,複数選択,5段階,日付"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("D5:D%d" % (4 + len(ROWS)))
    ws.freeze_panes = "A5"

    wb.save(XLSX)
    print("書き出しました: " + XLSX)


# ---------------------------------------------------------------- .gs を作る
def js(s):
    return "'" + str(s).replace("\\", "\\\\").replace("'", "\\'").replace("\n", "\\n") + "'"


def emit_item(row, out):
    youshi, setsu, no, kata, bun, sel, must, hosoku = row
    label = ("%s %s" % (no, bun)).strip()
    choices = [x.strip() for x in (sel or "").split("／") if x.strip()]

    if kata == "ページ":
        out.append("  it = f.addPageBreakItem().setTitle(%s);" % js(bun))
        if hosoku:
            out.append("  it.setHelpText(%s);" % js(hosoku))
    elif kata == "説明":
        out.append("  it = f.addSectionHeaderItem().setTitle(%s);" % js(bun))
        if hosoku:
            out.append("  it.setHelpText(%s);" % js(hosoku))
    elif kata == "短文":
        out.append("  it = f.addTextItem().setTitle(%s);" % js(label))
    elif kata == "段落":
        out.append("  it = f.addParagraphTextItem().setTitle(%s);" % js(label))
    elif kata == "日付":
        out.append("  it = f.addDateItem().setTitle(%s);" % js(label))
    elif kata == "単一選択":
        out.append("  it = f.addMultipleChoiceItem().setTitle(%s).setChoiceValues([%s]);"
                   % (js(label), ", ".join(js(c) for c in choices)))
    elif kata == "複数選択":
        out.append("  it = f.addCheckboxItem().setTitle(%s).setChoiceValues([%s]);"
                   % (js(label), ", ".join(js(c) for c in choices)))
    elif kata == "5段階":
        left = choices[0] if choices else "1"
        right = choices[1] if len(choices) > 1 else "5"
        out.append("  it = f.addScaleItem().setTitle(%s).setBounds(1, 5).setLabels(%s, %s);"
                   % (js(label), js(left), js(right)))
    else:
        return
    if kata not in ("ページ", "説明"):
        if hosoku:
            out.append("  it.setHelpText(%s);" % js(hosoku))
        out.append("  it.setRequired(%s);" % ("true" if str(must).strip() == "○" else "false"))


def write_gs():
    if not os.path.exists(XLSX):
        sys.exit("%s がありません。先に python tools/make_forms.py --bank を流してください。" % XLSX)
    ws = load_workbook(XLSX)["設問"]
    rows = []
    for r in range(5, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(COLS) + 1)]
        if not any(vals):
            continue
        rows.append([("" if v is None else str(v)) for v in vals])

    out = ["// 事後アンケートの Google フォームを作るスクリプト",
           "// docs/questionnaire/設問一覧.xlsx から自動生成。Excel を直して作り直すこと。",
           "//",
           "// 1. Google ドライブ →「新規 → その他 → Google Apps Script」",
           "// 2. この中身をぜんぶ貼り付けて保存",
           "// 3. 関数 createAll を実行",
           "// 4. 実行ログに2つのURLが出る",
           "//",
           "// ★先に whoami を実行して、どのアカウントで動くか確かめること。",
           "//   フォームは、ここに出たアカウントの Google ドライブに作られる。",
           "",
           "function whoami() {",
           "  Logger.log('いま動いているアカウント: ' + Session.getEffectiveUser().getEmail());",
           "}",
           "",
           "function createAll() {",
           "  var a = make1(), b = make2();",
           "  Logger.log('1回目\\n  回答用: ' + a[0] + '\\n  編集用: ' + a[1]);",
           "  Logger.log('2回目\\n  回答用: ' + b[0] + '\\n  編集用: ' + b[1]);",
           "}",
           ""]

    def form_fn(fn, name, desc, youshi):
        o = ["function %s() {" % fn,
             "  var f = FormApp.create(%s);" % js(name),
             "  f.setDescription(%s);" % js(desc),
             "  try { f.setCollectEmail(false); } catch (e) {}",
             "  var it;"]
        for row in rows:
            if row[0] == youshi:
                emit_item(row, o)
        o += ["  return [f.getPublishedUrl(), f.getEditUrl()];", "}", ""]
        return o

    out += form_fn("make1", "バスの人数アプリ ― 1回目",
                   "ゆきとかえりの2本に乗ったあとにお答えください。5分ほどで終わります。", R1)
    out += form_fn("make2", "バスの人数アプリ ― 2回目",
                   "ゆきとかえりの2本に乗ったあとにお答えください。8分ほどで終わります。", R2)

    with open(GS, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))
    n1 = len([r for r in rows if r[0] == R1])
    n2 = len([r for r in rows if r[0] == R2])
    print("書き出しました: " + GS)
    print("  1回目 %d 行 ／ 2回目 %d 行" % (n1, n2))


if __name__ == "__main__":
    if "--bank" in sys.argv:
        write_bank(force=("--force" in sys.argv))
    else:
        write_gs()
