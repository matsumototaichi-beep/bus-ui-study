# -*- coding: utf-8 -*-
"""事後アンケートを Google フォームで作るための道具（2026-10-05）

松本が**文言を自分で直せる**ようにするため、2段構えにしてある。

  1. 設問は Excel（docs/questionnaire/設問一覧.xlsx）に1行1問で入っている。
     **言い回しを変えたいときは、この Excel の「設問文」「選択肢」を書き換える。**
  2. そのあと下のコマンドを流すと、Excel の中身から
     Google Apps Script（docs/questionnaire/GoogleForms生成.gs）が作り直される。
     それを Google フォームのスクリプトエディタに貼って実行すると、フォームが5つできる。

    python tools/make_forms.py            # Excel を読んで .gs を作り直す
    python tools/make_forms.py --bank     # Excel 自体を作る（初回だけ。既にあれば拒否）
    python tools/make_forms.py --bank --force   # Excel を作り直す（手で直した分は消える）

★なぜフォームが5つなのか
  ・①の前半（S0・S1・S3）… 版による違いが無いので1つ
  ・①の後半（S4）        … 「あり→なし」と「なし→あり」の2つ
  ・②まとめ（S2・S5・S6）… 「ステッパー先」と「テンキー先」の2つ

  前半と後半を**別のフォームに分けている**のが肝。紙なら「ここで止めてください」で
  足りるが、Google フォームはセクションを分けても戻るボタンで前に戻れてしまう。
  別フォームにして、前半を送信してもらってから後半のURLを渡せば、確実に壁になる。
  （S4 で「わからない」ボタンの話を先に見せると、S3 の答えがそれを意識したものになる）
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

COLS = ["用紙", "節", "番号", "形式", "設問文", "選択肢（／区切り）", "必須", "補足（説明文）", "繰り返し"]

# 用紙の区分
F1A = "①前半"
F1B = "①後半"
F2 = "②まとめ"

METHODS = [
    ("ステッパー方式", "「−5」「−1」「＋1」「＋5」のボタンで、数を増やしたり減らしたりして合わせる"),
    ("テンキー方式", "数字のキーを押して、人数を直接入力する"),
]
BUTTONS = [
    ("「わからない」ボタン あり", "「わからない」というボタンが置かれていた便"),
    ("「わからない」ボタン なし", "同じ場所に何も置かれていなかった便"),
]

# ---------------------------------------------------------------- 設問の中身
# 形式：説明 / 短文 / 段落 / 単一選択 / 複数選択 / 5段階 / 日付 / ページ
ROWS = [
    # ---- ① 前半 ----------------------------------------------------------
    (F1A, "S0", "", "説明", "はじめに", "", "", "今日の乗車についてお答えください。所要5分ほどです。", ""),
    (F1A, "S0", "0-1", "短文", "配布カードに書かれたログインID（p01 など）", "", "○", "カードの左上に書いてあります", ""),
    (F1A, "S0", "0-2", "日付", "今日の日付", "", "○", "", ""),
    (F1A, "S0", "0-3", "単一選択", "今日は何日目ですか", "1日目／2日目", "○", "", ""),
    (F1A, "S0", "0-4", "単一選択", "この用紙の版", "v1／v2／v3／v4", "○", "配布カードに書いてあります", ""),
    (F1A, "S0", "0-5", "短文", "今日乗った1本目の発車時刻", "", "", "覚えている範囲で。空欄でも構いません", ""),
    (F1A, "S0", "0-6", "短文", "今日乗った2本目の発車時刻", "", "", "覚えている範囲で。空欄でも構いません", ""),

    (F1A, "S1", "", "ページ", "普段のことについて", "", "", "★1日目の方だけお答えください。2日目の方はそのまま次へ進んでください。", ""),
    (F1A, "S1", "1-1", "単一選択", "年代", "10代／20代／30代／40代以上", "", "", ""),
    (F1A, "S1", "1-2", "単一選択", "普段バスを利用する頻度", "ほぼ毎日／週に数回／月に数回／ほとんど乗らない", "", "", ""),
    (F1A, "S1", "1-3", "単一選択", "今回の路線に乗ったことがありましたか", "よく乗る／数回ある／初めて", "", "", ""),
    (F1A, "S1", "1-4", "単一選択", "実験で使った端末", "iPhone／Android／その他", "", "", ""),
    (F1A, "S1", "1-5", "単一選択", "利き手", "右／左", "", "", ""),
    (F1A, "S1", "1-6", "単一選択", "実験中、スマホは主にどちらの手で操作しましたか", "片手（親指）／両手／その時による", "", "", ""),

    (F1A, "S3", "", "ページ", "今日、人数を答えるときのことについて", "", "",
     "正解・不正解を見るものではありません。思ったとおりにお答えください。", ""),
    (F1A, "S3", "3-1", "単一選択", "今日、人数を答えるとき、実際に一人ずつ数えていましたか",
     "毎回数えた／だいたい数えた／目分量が多かった／ほとんど目分量", "○", "", ""),
    (F1A, "S3", "3-2", "複数選択", "数えるのが難しいと感じたのはどんなときですか（いくつでも）",
     "混んでいた／自分が立っていた／揺れた／降車が近かった／暗かった／その他", "", "", ""),
    (F1A, "S3", "3-3", "単一選択", "今日、数えきれていないのに、それらしい数を入れて送ったことがありますか",
     "何度もある／数回ある／ない／覚えていない", "○", "", ""),
    (F1A, "S3", "3-4", "段落", "（3-3で「ある」と答えた方）そうしたのはどんなときですか", "", "", "", ""),
    (F1A, "S3", "3-5", "段落", "人数を入れずに空欄のまま送ったことがありますか。あればその理由も", "", "", "", ""),
    (F1A, "S3", "3-6", "5段階", "今日の2便のあいだで、答えやすさは違いましたか", "違わない／大きく違った", "", "", ""),
    (F1A, "S3", "3-6b", "段落", "（3-6について）どう違ったか、よければ教えてください", "", "", "", ""),
    (F1A, "S3", "3-7", "単一選択", "今日、座っていた割合はどのくらいですか", "ほぼ座席／半々／ほぼ立席", "", "", ""),

    # ---- ① 後半（S4）-----------------------------------------------------
    (F1B, "S0", "0-1", "短文", "配布カードに書かれたログインID（p01 など）", "", "○", "前半と同じIDを入れてください", ""),
    (F1B, "S4", "", "説明", "今日の「わからない」ボタンについて", "", "",
     "今日の1本目と2本目で、画面に1か所だけ違いがありました。", ""),
    (F1B, "S4", "4-1", "単一選択", "今日の1本目と2本目で、画面に違いがあったことに気づきましたか",
     "はっきり気づいた／なんとなく気づいた／気づかなかった", "○", "", ""),
    (F1B, "S4", "", "説明", "違いはここです", "", "", "人数を入れる欄のすぐ下に――", "ボタンごと"),
    (F1B, "S4", "4-2", "単一選択", "「わからない」ボタンがあったほうの乗車を覚えていますか",
     "1本目／2本目／覚えていない", "", "", ""),
    (F1B, "S4", "4-3", "単一選択", "ボタンがあったとき、数えきれないときはどうしていましたか",
     "ボタンを押した／それらしい数を入れた／空欄で送った／覚えていない", "", "", ""),
    (F1B, "S4", "4-4", "単一選択", "ボタンがなかったとき、数えきれないときはどうしていましたか",
     "それらしい数を入れた／空欄で送った／クラスだけ選んで送った／覚えていない", "", "", ""),
    (F1B, "S4", "4-5", "単一選択", "どちらのほうが正直に答えられたと感じますか",
     "ボタンがあった便／ボタンがなかった便／変わらない", "", "", "ボタン順"),
    (F1B, "S4", "4-6", "段落", "そう思う理由", "", "", "", ""),
    (F1B, "S4", "4-7", "5段階", "「わからない」と答えるのに、抵抗や後ろめたさはありましたか",
     "まったくなかった／とてもあった", "", "", ""),
    (F1B, "S4", "4-7b", "段落", "（4-7について）よければ理由を教えてください", "", "", "", ""),

    # ---- ② まとめ --------------------------------------------------------
    (F2, "S0", "", "説明", "はじめに", "", "",
     "2日間すべての乗車を終えた方にお答えいただきます。所要5分ほどです。", ""),
    (F2, "S0", "0-1", "短文", "配布カードに書かれたログインID（p01 など）", "", "○", "", ""),
    (F2, "S0", "0-2", "単一選択", "この用紙の版", "v1／v2／v3／v4", "○", "配布カードに書いてあります", ""),

    (F2, "S2", "", "ページ", "2つの入力方式について", "", "",
     "2日間で、人数の入れ方が日によって違いました。次の2つです。", ""),
    (F2, "S2", "2-0", "単一選択", "1日目に使ったのはどちらの方式でしたか",
     "ステッパー方式／テンキー方式／覚えていない", "", "思い出せる範囲で構いません", ""),
    (F2, "S2", "2-1", "5段階", "{方式}は押しやすかった", "まったくそう思わない／とてもそう思う", "", "", "方式ごと"),
    (F2, "S2", "2-2", "5段階", "{方式}は素早く答えられた", "まったくそう思わない／とてもそう思う", "", "", "方式ごと"),
    (F2, "S2", "2-3", "5段階", "{方式}では自分が思った通りの人数を入力できた", "まったくそう思わない／とてもそう思う", "", "", "方式ごと"),
    (F2, "S2", "2-4", "5段階", "{方式}は頭を使う・疲れると感じた", "まったくそう思わない／とてもそう思う", "", "", "方式ごと"),
    (F2, "S2", "2-5", "単一選択", "2つのうち、答えやすかったのはどちらですか",
     "{方式1}／{方式2}／どちらとも言えない", "", "", ""),
    (F2, "S2", "2-6", "段落", "それはなぜですか", "", "", "", ""),
    (F2, "S2", "2-7", "単一選択", "揺れている車内では、どちらが押し間違えにくかったですか",
     "{方式1}／{方式2}／どちらとも言えない", "", "", ""),
    (F2, "S2", "2-8", "段落", "それはなぜですか", "", "", "", ""),

    (F2, "S5", "", "ページ", "アプリ全体の使いやすさ", "", "",
     "国際的に使われている標準の10項目です。項目の追加・削除・並べ替えはしないでください。", ""),
] + [
    (F2, "S5", "5-%d" % i, "5段階", t, "まったくそう思わない／とてもそう思う", "", "", "")
    for i, t in enumerate([
        "このアプリを頻繁に使いたいと思う",
        "このアプリは不必要に複雑だと思う",
        "このアプリは簡単に使えると思う",
        "このアプリを使うには、詳しい人のサポートが必要だと思う",
        "このアプリのさまざまな機能は、うまくまとまっていると思う",
        "このアプリには一貫性のないところが多いと思う",
        "たいていの人はこのアプリの使い方をすぐに覚えられると思う",
        "このアプリはとても使いにくいと思う",
        "このアプリを自信をもって使えた",
        "このアプリを使い始める前に、いろいろ覚える必要があった"], start=1)
] + [
    (F2, "S6", "", "ページ", "自由にお書きください", "", "", "", ""),
    (F2, "S6", "6-1", "段落", "人数を答えるとき、いちばん面倒だと感じたことは何ですか", "", "", "", ""),
    (F2, "S6", "6-2", "段落", "「次に停まるバス停」を選ぶとき、困ったことはありましたか", "", "", "", ""),
    (F2, "S6", "6-3", "段落", "このアプリをこう変えたら答えやすくなる、という案があれば", "", "", "", ""),
    (F2, "S6", "6-4", "段落", "実験全体について、気づいたことがあれば", "", "", "", ""),
]


# ---------------------------------------------------------------- Excel を作る
def write_bank(force=False):
    if os.path.exists(XLSX) and not force:
        sys.exit("既に %s があります。手で直した内容を消さないため、上書きしません。\n"
                 "本当に作り直すなら --force を付けてください。" % XLSX)
    os.makedirs(OUTDIR, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "設問"
    ws["A1"] = "事後アンケートの設問一覧"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("「設問文」「選択肢」「補足」を書き換えてから python tools/make_forms.py を流すと、"
                "Googleフォーム生成スクリプトが作り直されます。")
    ws["A2"].font = Font(size=9, color="555555")
    ws["A3"] = ("※「形式」「繰り返し」「番号」は仕組みに関わるので、意味が分かるとき以外は触らないでください。"
                "{方式} {方式1} {方式2} は自動で置き換わる場所です。")
    ws["A3"].font = Font(size=9, color="C00000")

    thin = Side(style="thin", color="AAAAAA")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, c in enumerate(COLS, start=1):
        cell = ws.cell(row=5, column=i, value=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DCE6F1")
        cell.border = box
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate([10, 6, 7, 10, 52, 46, 6, 34, 10], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r, row in enumerate(ROWS, start=6):
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
    dv.add("D6:D%d" % (5 + len(ROWS)))
    ws.freeze_panes = "A6"

    ws2 = wb.create_sheet("読み方")
    for i, (a, b) in enumerate([
        ("用紙", "①前半＝解散直後の1枚目／①後半＝それを出してから渡すS4／②まとめ＝2日目の最後"),
        ("節", "S0 はじめに／S1 普段のこと／S3 今日の答え方／S4 わからないボタン／S2 2方式の比較／S5 使いやすさ／S6 自由記述"),
        ("形式", "説明＝文章だけ／ページ＝ここから新しいページ／短文／段落／単一選択／複数選択／5段階／日付"),
        ("選択肢", "「／」で区切る。5段階のときは「左端のラベル／右端のラベル」を書く"),
        ("必須", "○ を入れると回答必須になる"),
        ("繰り返し", "方式ごと＝ステッパーとテンキーの2回くり返す／ボタン順＝版によって順番が入れ替わる"),
        ("{方式}", "「ステッパー方式」か「テンキー方式」に自動で置き換わる"),
        ("{方式1} {方式2}", "その版で先に出すほう／後に出すほうに自動で置き換わる"),
    ], start=2):
        ws2.cell(row=i, column=1, value=a).font = Font(bold=True)
        ws2.cell(row=i, column=2, value=b)
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 96
    ws2["A1"] = "列の読み方"
    ws2["A1"].font = Font(bold=True, size=13)

    wb.save(XLSX)
    print("書き出しました: " + XLSX)


# ---------------------------------------------------------------- .gs を作る
def js(s):
    return "'" + str(s).replace("\\", "\\\\").replace("'", "\\'").replace("\n", "\\n") + "'"


def emit_item(row, out, subs):
    youshi, setsu, no, kata, bun, sel, must, hosoku, kurikaeshi = row
    for k, v in subs.items():
        bun = bun.replace(k, v)
        sel = (sel or "").replace(k, v)
        hosoku = (hosoku or "").replace(k, v)
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
    wb = load_workbook(XLSX)
    ws = wb["設問"]
    rows = []
    for r in range(6, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(COLS) + 1)]
        if not any(vals):
            continue
        rows.append([("" if v is None else str(v)) for v in vals])

    def pick(youshi):
        return [r for r in rows if r[0] == youshi]

    out = ["// 事後アンケートの Google フォームを作るスクリプト",
           "// docs/questionnaire/設問一覧.xlsx から自動生成。直接ここを直さず、Excel を直して作り直すこと。",
           "//",
           "// 使い方：",
           "//   1. Google ドライブで「新規 → その他 → Google Apps Script」を開く",
           "//   2. このファイルの中身をぜんぶ貼り付ける",
           "//   3. 関数 createAll を選んで実行する（初回は権限の確認が出る）",
           "//   4. 実行ログに5つのフォームのURLが出るので控える",
           "//",
           "// うまくいかないときは、createAll ではなく下の5つを1つずつ実行してもよい：",
           "//   makeDaily1 / makeDaily2_AriFirst / makeDaily2_NashiFirst",
           "//   makeFinal_StepperFirst / makeFinal_NumpadFirst",
           "",
           "function createAll() {",
           "  var urls = [];",
           "  urls.push(['①前半（全員・解散直後）', makeDaily1()]);",
           "  urls.push(['①後半 S4 … v1 v2 の人に渡す', makeDaily2_AriFirst()]);",
           "  urls.push(['①後半 S4 … v3 v4 の人に渡す', makeDaily2_NashiFirst()]);",
           "  urls.push(['②まとめ … v1 v3 の人に渡す', makeFinal_StepperFirst()]);",
           "  urls.push(['②まとめ … v2 v4 の人に渡す', makeFinal_NumpadFirst()]);",
           "  for (var i = 0; i < urls.length; i++) {",
           "    Logger.log(urls[i][0] + '\\n  回答用: ' + urls[i][1][0] + '\\n  編集用: ' + urls[i][1][1]);",
           "  }",
           "}",
           ""]

    def form_fn(fn, name, desc, rows_, subs, button_order=None, method_order=None):
        o = ["function %s() {" % fn,
             "  var f = FormApp.create(%s);" % js(name),
             "  f.setDescription(%s);" % js(desc),
             "  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある",
             "  var it;"]
        for row in rows_:
            k = (row[8] or "").strip()
            if k == "ボタンごと":
                o.append("  f.addSectionHeaderItem().setTitle(%s).setHelpText(%s);"
                         % (js(row[4]), js(row[7])))
                for nm, d in (button_order or BUTTONS):
                    o.append("  f.addSectionHeaderItem().setTitle(%s).setHelpText(%s);"
                             % (js("【%s】" % nm), js(d)))
                continue
            if k == "ボタン順":
                names = [b[0] for b in (button_order or BUTTONS)]
                vals = ["ボタンがあった便" if "あり" in n else "ボタンがなかった便" for n in names]
                o.append("  it = f.addMultipleChoiceItem().setTitle(%s).setChoiceValues([%s]);"
                         % (js(("%s %s" % (row[2], row[4])).strip()),
                            ", ".join(js(v) for v in vals + ["変わらない"])))
                o.append("  it.setRequired(false);")
                continue
            if k == "方式ごと":
                continue
            if row[2] == "2-0":
                emit_item(row, o, subs)
                # 2-0 の直後に、方式ごとの塊を版の順で2回出す
                reps = [x for x in rows_ if (x[8] or "").strip() == "方式ごと"]
                for nm, d in (method_order or METHODS):
                    o.append("  f.addSectionHeaderItem().setTitle(%s).setHelpText(%s);"
                             % (js("【%s】" % nm), js(d)))
                    for rr in reps:
                        emit_item(rr, o, {"{方式}": nm})
                continue
            emit_item(row, o, subs)
        o.append("  return [f.getPublishedUrl(), f.getEditUrl()];")
        o.append("}")
        o.append("")
        return o

    lead1 = "今日の乗車についてお答えください。所要5分ほどです。"
    lead2 = "前半を送信された方にお渡しするものです。引き続きお答えください。"
    lead3 = "2日間すべての乗車を終えた方にお答えいただきます。所要5分ほどです。"

    out += form_fn("makeDaily1", "バス車内の人数を答えるアプリについて ― 今日の分（前半）",
                   lead1, pick(F1A), {})
    out += form_fn("makeDaily2_AriFirst", "バス車内の人数を答えるアプリについて ― 今日の分（後半）A",
                   lead2, pick(F1B), {}, button_order=BUTTONS)
    out += form_fn("makeDaily2_NashiFirst", "バス車内の人数を答えるアプリについて ― 今日の分（後半）B",
                   lead2, pick(F1B), {}, button_order=list(reversed(BUTTONS)))
    out += form_fn("makeFinal_StepperFirst", "バス車内の人数を答えるアプリについて ― まとめ A",
                   lead3, pick(F2),
                   {"{方式1}": METHODS[0][0], "{方式2}": METHODS[1][0]}, method_order=METHODS)
    out += form_fn("makeFinal_NumpadFirst", "バス車内の人数を答えるアプリについて ― まとめ B",
                   lead3, pick(F2),
                   {"{方式1}": METHODS[1][0], "{方式2}": METHODS[0][0]},
                   method_order=list(reversed(METHODS)))

    with open(GS, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))
    print("書き出しました: " + GS)
    print("  設問 %d 行から、フォーム5つ分のスクリプトを作りました。" % len(rows))
    print("\n次の手順：")
    print("  1. Google ドライブ →「新規 → その他 → Google Apps Script」")
    print("  2. %s の中身を貼り付け" % os.path.basename(GS))
    print("  3. 関数 createAll を実行 → 実行ログに5つのURLが出ます")


if __name__ == "__main__":
    if "--bank" in sys.argv:
        write_bank(force=("--force" in sys.argv))
    else:
        write_gs()
