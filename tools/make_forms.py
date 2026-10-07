# -*- coding: utf-8 -*-
"""事後アンケートを作る

    python tools/make_forms.py --bank --force   設問一覧.xlsx を作り直す
    python tools/make_forms.py                  Excel から GoogleForms生成.gs を作る

設問の正本は docs/questionnaire/設問一覧.xlsx。文言を直すときはそこを直す。

アンケートは2種類。往復（ゆき・かえり）が終わったら答える。
  1回目 … 問1〜7 ＋ 自由記述（8問）
  2回目 … 問1〜7（1回目とまったく同じ）＋「わからない」ボタンの3問 ＋ 自由記述（11問）
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

# ★2026-10-07:「何が分かるか」列を足した。研究で知りたいこと（入れ方・わからないボタン・混雑）の
#   どれに使う設問かを1行で書く。フォームと紙には出ない。
COLS = ["用紙", "節", "番号", "形式", "設問文", "選択肢（／区切り）", "必須", "補足（説明文）", "何が分かるか"]
KATA = ["表紙", "説明", "ページ", "ID", "短文", "段落", "単一選択", "複数選択", "5段階", "日付"]

R1 = "1回目"
R2 = "2回目"

# 5段階はすべて「1＝少ない・弱い、5＝多い・強い」の向きにそろえる
# ★2026-10-07:「混んでいましたか」「答えやすかったですか」と片側だけで聞くと「はい」側に寄るので、
#   問2・3・6・7 は「〜の混み具合」「〜の答えやすさ」と聞き、向きは両端のラベルだけで示す。
YASUI = "答えにくかった／答えやすかった"
KONDA = "空いていた／混んでいた"
HINDO = "なかった／何度もあった"
# ★2026-10-07: 選択肢は「ボタンがあった回」ではなく回の番号にする。ボタンが答えの理由だと匂わせず、
#   先頭に出る選択肢がボタンのある回に偏らない（組で逆になる）。覚え違いがあっても組から実際の条件に戻せる。
DOTCHI = "1回目／2回目／変わらない"

# ★2026-10-07: 88回答えた直後に26問・35問は重すぎる（監査）。研究で知りたい3つ
#   （入れ方・わからないボタン・混雑）に対応しない設問と、アプリの記録で分かる設問を消した。
#   任意の設問は面倒なものだけ抜けて集計が偏るので、すべて必須にした。
#   問1〜7 は両方の回でまったく同じにして、回ごとの差をそのまま比べる。
#   入れ方（問6・7）を両方の回で聞くのは、1回目は数字キーが先・2回目は＋−ボタンが先だから。
#   片方の回だけだと、全員が同じ順番の答えしか取れず「あとで使ったほうが慣れていた」と区別できない。
#   「わからない」ボタンは組によって1回目にあったり2回目にあったりするので、1回目には一切出さない。


def kyotsu(youshi):
    """問1〜7。1回目と2回目でまったく同じ"""
    return [
        (youshi, "はじめに", "問1", "ID", "カードに書いてあるID（例：A）", "", "○", "",
         "アンケートとアプリの記録をつなぐ"),

        # 混雑に関係する設問
        # ★2026-10-07:「今日の」だと、1回目と2回目を同じ日にした人は2回目にどちらの便か迷うので「今回の往復」にした
        (youshi, "混み具合", "", "説明", "今回の往復について", "", "", "", ""),
        (youshi, "混み具合", "問2", "5段階", "ゆき（吹田駅 → 桃山台駅）のバスの混み具合", KONDA, "○", "",
         "便ごとの混み具合。その便で使った入れ方の点数と並べて見る"),
        (youshi, "混み具合", "問3", "5段階", "かえり（桃山台駅 → 吹田駅）のバスの混み具合", KONDA, "○", "",
         "同上"),
        # ★2026-10-07:「混んでいて、」を外した。立っていた・降りる人が多かったなど別の理由で数えきれなかった人が
        #   答えられず、混雑との関係も答える側に決めさせてしまうため。混雑との関係は問2・3 と並べて見る。
        (youshi, "混み具合", "問4", "5段階", "人数を数えきれないことがありましたか", HINDO, "○", "",
         "数えきれないことがどれだけあったか。問2・3 と並べて、混んでいると数えにくいかを見る"),
        (youshi, "混み具合", "問5", "5段階", "数えきれないまま、だいたいの人数を入れて送ったことがありましたか", HINDO, "○", "",
         "数えられないときに人数を入れてしまったか（記録では分からない）。ボタンがあった回となかった回で比べる"),

        # 混雑に関係しない設問
        (youshi, "入れ方", "", "説明", "人数の入れ方について", "", "", "", ""),
        (youshi, "入れ方", "問6", "5段階", "数字キーの答えやすさ", YASUI, "○",
         "0〜9 のキーで人数を打つ画面", "どちらの入れ方が答えやすいか"),
        (youshi, "入れ方", "問7", "5段階", "＋−ボタンの答えやすさ", YASUI, "○",
         "−5・−1・＋1・＋5 で人数を合わせる画面", "同上"),
    ]


def jiyu(youshi, no):
    return [(youshi, "最後に", no, "段落", "答えにくかったことがあれば書いてください（なければ「なし」）", "", "○", "",
             "点数の理由")]


ROWS = (
    # ================================ 1回目 ================================
    [(R1, "表紙", "", "表紙", "バスの人数アプリ　1回目のアンケート", "", "",
      "往復が終わったら答えてください（2分ほど）。思ったとおりに選んでください。", "")]
    + kyotsu(R1)
    + jiyu(R1, "問8")

    # ================================ 2回目 ================================
    + [(R2, "表紙", "", "表紙", "バスの人数アプリ　2回目のアンケート", "", "",
        "往復が終わったら答えてください（3分ほど）。思ったとおりに選んでください。", "")]
    + kyotsu(R2)
    + [
        # ★2026-10-07: 問5 などを答え終えてから見せるよう、ページを分ける。
        #   どちらの回にボタンがあったかは書かない（組によって逆なので決めつけない）。
        (R2, "わからない", "", "ページ", "「わからない」ボタンについて", "", "",
         "1回目と2回目のどちらか一方だけ、人数を入れるところの下に「わからない」ボタンがありました。", ""),
        (R2, "わからない", "問8", "単一選択", "「わからない」ボタンがあったのは、どちらの回ですか",
         "1回目／2回目／覚えていない", "○", "",
         "ボタンに気づいていたか"),
        (R2, "わからない", "問9", "単一選択", "人数を答えやすかったのは、どちらの回ですか", DOTCHI, "○", "",
         "ボタンがあると答えやすいか（組から、ボタンがあった回を選んだかを見る）"),
        # ★2026-10-07:「正直に答えられた」は、もう一方の回は正直でなかったと認めさせる聞き方で「変わらない」に寄る。
        #   ボタンの狙い（数えられないときに数を作らない）をそのまま言葉にした。
        (R2, "わからない", "問10", "単一選択", "数えたとおりの人数を答えられたのは、どちらの回ですか", DOTCHI, "○", "",
         "ボタンがあると正直に答えられるか（同上）"),
    ]
    + jiyu(R2, "問11")
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
    for i, w in enumerate([9, 11, 7, 10, 48, 36, 6, 38, 44], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for r, row in enumerate(ROWS, start=5):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = box
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (5, 6, 8, 9)))
        if row[3] in ("表紙", "説明", "ページ"):
            for c in range(1, len(COLS) + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="F2F2F2")

    dv = DataValidation(type="list", formula1='"%s"' % ",".join(KATA), allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("D5:D%d" % (4 + len(ROWS)))
    ws.freeze_panes = "A5"

    wb.save(XLSX)
    print("書き出しました: " + XLSX)


# ---------------------------------------------------------------- .gs を作る
def js(s):
    return "'" + str(s).replace("\\", "\\\\").replace("'", "\\'").replace("\n", "\\n") + "'"


def emit_item(row, out):
    youshi, setsu, no, kata, bun, sel, must, hosoku = row[:8]
    label = ("%s %s" % (no, bun)).strip()
    choices = [x.strip() for x in (sel or "").split("／") if x.strip()]

    if kata == "表紙":
        return                      # フォームの題名と説明文に使う（form_fn）
    elif kata == "ページ":
        out.append("  it = f.addPageBreakItem().setTitle(%s);" % js(bun))
        if hosoku:
            out.append("  it.setHelpText(%s);" % js(hosoku))
    elif kata == "説明":
        out.append("  it = f.addSectionHeaderItem().setTitle(%s);" % js(bun))
        if hosoku:
            out.append("  it.setHelpText(%s);" % js(hosoku))
    elif kata == "ID":
        # ★2026-10-07: IDを打ち間違えるとアプリの記録とつなげないので、A〜X の1文字だけ受け付ける
        out.append("  it = f.addTextItem().setTitle(%s);" % js(label))
        out.append("  it.setValidation(FormApp.createTextValidation().setHelpText(%s)"
                   ".requireTextMatchesPattern(%s).build());"
                   % (js("カードのIDを1文字で入れてください（例：A）"), js(r"^\s*[A-Xa-xＡ-Ｘａ-ｘ]\s*$")))
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
           "// 5. 回答用URLを Google にログインしていないブラウザで開き、答えられるか確かめる",
           "//",
           "// ★先に whoami を実行して、どのアカウントで動くか確かめること。",
           "//   フォームは、ここに出たアカウントの Google ドライブに作られる。",
           "//",
           "// ★2026-10-07 に設問を減らした（1回目8問・2回目11問）。前に作った2つのフォームはゴミ箱に入れること。",
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

    def form_fn(fn, youshi):
        hyoshi = [r for r in rows if r[0] == youshi and r[3] == "表紙"]
        if not hyoshi:
            sys.exit("%s の「表紙」の行がありません。" % youshi)
        name, desc = hyoshi[0][4], hyoshi[0][7]
        o = ["function %s() {" % fn,
             "  var f = FormApp.create(%s);" % js(name),
             "  f.setDescription(%s);" % js(desc),
             "  try { f.setCollectEmail(false); } catch (e) {}",
             # ★2026-10-07: 大学アカウントで作ると「大学の人だけ回答できる」になることがある。
             #   参加者はスマホで大学アカウントにログインしているとは限らないので外す。
             #   公開状態の設定が無い・効かない環境もあるので try で囲み、手順5で実際に開いて確かめる。
             "  try { f.setRequireLogin(false); } catch (e) {}",
             "  try { f.setPublished(true); } catch (e) {}",
             "  var it;"]
        for row in rows:
            if row[0] == youshi:
                emit_item(row, o)
        o += ["  return [f.getPublishedUrl(), f.getEditUrl()];", "}", ""]
        return o

    out += form_fn("make1", R1)
    out += form_fn("make2", R2)

    with open(GS, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))
    n1 = len([r for r in rows if r[0] == R1 and r[2]])
    n2 = len([r for r in rows if r[0] == R2 and r[2]])
    print("書き出しました: " + GS)
    print("  1回目 %d 問 ／ 2回目 %d 問" % (n1, n2))


if __name__ == "__main__":
    if "--bank" in sys.argv:
        write_bank(force=("--force" in sys.argv))
    else:
        write_gs()
