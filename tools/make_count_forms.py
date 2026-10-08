# -*- coding: utf-8 -*-
"""計数研究（車内を1から数える vs 乗った人・降りた人を数える）のアンケートを作る

    python tools/make_count_forms.py

設問の正本はこのファイルの ROWS。文言を直すときはここを直して流し直す。
★乗降人数計算/ に書き出す。
  計数研究アンケート生成.gs … Google フォームを作る Apps Script（関数 createCount）
  計数研究アンケート.docx   … 紙の予備

往復（ゆき＝人数アプリ、かえり＝乗降アプリ）が終わったら、南口で1回だけ答える。
答えた人数・時間・タップ・位置はアプリの記録で分かるので聞かない。聞くのは主観だけ。
"""
import os, sys
sys.dont_write_bytecode = True          # tools/ に .pyc を増やさない
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from make_forms import emit_item, js    # .gs の1問ぶんの書き方は事後アンケートと同じ
import make_questionnaire as mq         # 紙の見た目（字・大きさ・余白）も同じ

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(os.path.dirname(ROOT), "★乗降人数計算")
GS = os.path.join(OUTDIR, "計数研究アンケート生成.gs")
DOCX = os.path.join(OUTDIR, "計数研究アンケート.docx")

Y = "計数研究"
# 選択肢は参加者向け案内の「回数」の表（「ゆき：車内の人数を1から数える／かえり：乗った人・降りた人を数える」）と一字一句同じ。
# ゆきが必ず人数アプリなので、向きとやり方を両方書いておけば覚え違いが起きにくい。
# 「乗った人」「降りた人」は乗降アプリのボタン（「乗った人　＋1」「降りた人　＋1」）とも同じ。
# ★「ゆき：1から数える」だけだと、何を1から数えるのか（どちらのやり方も1人ずつ数える）が分からないので「車内の人数を」を省かない。
# 並びは経験した順（ゆき→かえり）。UI研究の「1回目／2回目／変わらない」と同じ考え方。
YUKI = "ゆき：車内の人数を1から数える"
KAERI = "かえり：乗った人・降りた人を数える"
DOTCHI = "／".join([YUKI, KAERI, "変わらない"])
# 混雑は「混んでいましたか」と聞かず、混んだときに起きること（数えきれない）を聞く。
# 混み具合はアプリの人数で分かるので、それと並べて見る。
# 人数アプリには「わからない」ボタンがあるが、乗降アプリには無い（＋1を押さなかった人は記録に残らない）。
DEKINAI = "／".join([YUKI, KAERI, "両方", "どちらもなかった"])

# make_forms.py の ROWS と同じ並び:
# (用紙, 節, 番号, 形式, 設問文, 選択肢（／区切り）, 必須, 補足, 何が分かるか)
# 任意の設問は面倒なものだけ抜けて偏るので、すべて必須にする。
ROWS = [
    (Y, "表紙", "", "表紙", "バスの人数アプリ　計数研究のアンケート", "", "",
     "往復が終わったら答えてください（2分ほど）。思ったとおりに選んでください。", ""),
    (Y, "", "問1", "ID", "カードに書いてあるID（例：A）", "", "○", "",
     "アンケートとアプリの記録をつなぐ"),
    (Y, "", "問2", "単一選択", "数えるのが楽だったのは、どちらのやり方ですか", DOTCHI, "○", "",
     "どちらのやり方が楽か"),
    (Y, "", "問3", "単一選択", "正確に数えられたと思うのは、どちらのやり方ですか", DOTCHI, "○", "",
     "どちらのやり方が正確だと本人が思うか（記録では分からない）"),
    (Y, "", "問4", "単一選択", "数えきれないことがあったのは、どちらのやり方ですか", DEKINAI, "○", "",
     "混んだときに数えきれなくなるのはどちらか。アプリの人数（混み具合）と並べて見る"),
    (Y, "", "問5", "段落", "数えにくかったことがあれば書いてください（なければ「なし」）", "", "○", "",
     "問2〜4 の理由"),
]


# ---------------------------------------------------------------- .gs を作る
def write_gs():
    hyoshi = [r for r in ROWS if r[3] == "表紙"][0]
    out = ["// 計数研究のアンケートの Google フォームを作るスクリプト",
           "// bus-ui-study/tools/make_count_forms.py から自動生成。設問を直すときは make_count_forms.py を直して作り直すこと。",
           "//",
           "// 1. Google ドライブ →「新規 → その他 → Google Apps Script」",
           "// 2. この中身をぜんぶ貼り付けて保存",
           "// 3. 関数 createCount を実行",
           "// 4. 実行ログに2つのURLが出る",
           "// 5. 回答用URLを Google にログインしていないブラウザで開き、答えられるか確かめる",
           "//",
           "// ★先に whoami を実行して、どのアカウントで動くか確かめること。",
           "//   フォームは、ここに出たアカウントの Google ドライブに作られる。",
           "",
           "function whoami() {",
           "  Logger.log('いま動いているアカウント: ' + Session.getEffectiveUser().getEmail());",
           "}",
           "",
           "function createCount() {",
           "  var f = FormApp.create(%s);" % js(hyoshi[4]),
           "  f.setDescription(%s);" % js(hyoshi[7]),
           "  try { f.setCollectEmail(false); } catch (e) {}",
           # 大学アカウントで作ると「大学の人だけ回答できる」になることがある。事後アンケートと同じく外す
           "  try { f.setRequireLogin(false); } catch (e) {}",
           "  try { f.setPublished(true); } catch (e) {}",
           "  var it;"]
    for row in ROWS:
        emit_item(row, out)
    out += ["  Logger.log('計数研究\\n  回答用: ' + f.getPublishedUrl() + '\\n  編集用: ' + f.getEditUrl());",
            "  return [f.getPublishedUrl(), f.getEditUrl()];",
            "}",
            ""]
    with open(GS, "w", encoding="utf-8", newline="\n") as fp:
        fp.write("\n".join(out))
    print("書き出しました: %s（%d問）" % (GS, len([r for r in ROWS if r[2]])))


# ---------------------------------------------------------------- 紙を作る
def emit_doc(doc, row, first):
    # 選択肢が長い（「かえり：乗った人・降りた人を数える」）ので、1行に並べると途中で折り返す。
    # 選択肢だけ1つずつ行を分け、ほかは make_questionnaire.py とまったく同じに出す。
    if row[3] != "単一選択":
        mq.emit(doc, row, first)
        return
    first[0] = False
    no, bun, sel, must = row[2], row[4], row[5], row[6]
    mq.p(doc, "%s  %s" % (no, bun), bold=(must == "○"), before=4, after=1)
    choices = [x.strip() for x in sel.split("／") if x.strip()]
    for i, c in enumerate(choices):
        mq.p(doc, "　　□ " + c, size=10, after=(3 if i == len(choices) - 1 else 1))


def write_docx():
    doc = mq.Document()
    mq.setup(doc)
    hyoshi = [r for r in ROWS if r[3] == "表紙"][0]
    mq.p(doc, hyoshi[4], size=15, bold=True, align=mq.WD_ALIGN_PARAGRAPH.CENTER, after=4)
    mq.p(doc, hyoshi[7], size=10, color=(0x55, 0x55, 0x55), align=mq.WD_ALIGN_PARAGRAPH.CENTER, after=8)
    first = [True]
    for row in ROWS:
        emit_doc(doc, row, first)
    mq.p(doc, "ご協力ありがとうございました。", size=11, bold=True,
         align=mq.WD_ALIGN_PARAGRAPH.CENTER, before=10)
    doc.save(DOCX)
    print("書き出しました: %s（%d問）" % (DOCX, len([r for r in ROWS if r[2]])))


if __name__ == "__main__":
    os.makedirs(OUTDIR, exist_ok=True)
    write_gs()
    write_docx()
