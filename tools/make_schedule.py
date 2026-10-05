# -*- coding: utf-8 -*-
"""参加者スケジュール表を Excel で作る（2026-09-26）

「誰がいつ来るのか」を当日その場で迷わないための表。
**Excel なので、作ったあとは自由に書き換えられる。**
日程が動いても、この生成スクリプトを回し直す必要はない（回すと上書きされるので注意）。

    python tools/make_schedule.py            # schedule.xlsx を作る
    python tools/make_schedule.py 別名.xlsx  # 名前を指定して作る

シート構成
  1. 参加者名簿 … 番号・組・群・アンケート版は埋めてある。氏名/連絡先/日付を書き込む
  2. 日程     … 使える平日16日 × 昼枠・夕方枠。便の時刻まで入っている
  3. 当日の流れ … 昼の日・夕方の日のタイムラインと持ち物
  4. 便と条件  … どの便でどちらのURLを配るか（群1/群2）
  5. 準備チェック … 当日までにやることの一覧（docs/preflight_checklist.md と対）

★ 氏名や連絡先が入るので .gitignore 済み。public リポジトリなので commit しないこと。
"""
import sys, os, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEEK = "月火水木金土日"

# 使える平日（10/12 スポーツの日、11/3 文化の日、土日を除く。operation_plan.md §3）
DATES = [(10, 13), (10, 14), (10, 15), (10, 16), (10, 19), (10, 20), (10, 23),
         (11, 2), (11, 4), (11, 5), (11, 6), (11, 9), (11, 10), (11, 11), (11, 12), (11, 13)]

# 枠ごとの便（operation_plan.md §1。2026年4月1日改正のダイヤで照合済み）
SLOTS = [
    {"name": "昼（空いている）",   "meet": "11:35", "go": "12:03", "arr": "12:48",
     "back": "13:48", "ret": "14:32", "hold": "約3時間"},
    {"name": "夕方（混んでいる）", "meet": "15:25", "go": "15:55", "arr": "16:40",
     "back": "17:18", "ret": "18:03", "hold": "約2時間40分"},
]

HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
SUB_FILL  = PatternFill("solid", fgColor="F2F2F2")
SPARE_FILL= PatternFill("solid", fgColor="FBF4E6")
THIN      = Side(style="thin", color="AAAAAA")
BOX       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header(ws, row, cols):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sheet_members(wb):
    ws = wb.create_sheet("参加者名簿")
    ws["A1"] = "参加者名簿"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("番号・組・群・アンケート版は決まっています。氏名／連絡先／実施日を埋めてください。"
                "カードは必ず番号順に3枚ずつ渡すこと（順番を崩すと群とアンケートの釣り合いが崩れます）。")
    ws["A2"].font = Font(size=9, color="555555")

    cols = ["番号", "氏名", "連絡先", "組", "群", "アンケート版",
            "昼の日", "夕方の日", "状態", "備考"]
    header(ws, 4, cols)
    widths(ws, [7, 14, 20, 6, 6, 11, 12, 12, 10, 26])

    for i in range(1, 25):
        r = 4 + i
        team  = (i - 1) // 3 + 1
        group = 1 if ((i - 1) // 3) % 2 == 0 else 2
        form  = ((i - 1) % 4) + 1
        vals = ["p%02d" % i, "", "", "組%d" % team, "群%d" % group, "v%d" % form, "", "", "", ""]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BOX
            if c in (1, 4, 5, 6):
                cell.alignment = Alignment(horizontal="center")
            if i > 12:                      # 組5〜8は予備
                cell.fill = SPARE_FILL
        if i == 13:
            ws.cell(row=r, column=10, value="↓ ここから下は予備（欠員が出たとき用）")

    dv = DataValidation(type="list", formula1='"未連絡,打診中,確定,欠席"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("I5:I28")

    ws.freeze_panes = "A5"
    return ws


def sheet_dates(wb):
    ws = wb.create_sheet("日程")
    ws["A1"] = "日程（使える平日16日 × 昼枠・夕方枠）"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("実施は4日（昼2日・夕方2日）。残りは予備日です。"
                "土日はこの4便が土休日ダイヤに存在しないので使えません。"
                "参加者の都合を聞いて「状態」を埋めてください。")
    ws["A2"].font = Font(size=9, color="555555")

    cols = ["日付", "曜日", "枠", "集合", "1本目 南口発", "桃山台着",
            "2本目 桃山台発", "南口着＝解散", "拘束", "担当する組", "参加者（3名）",
            "計数役（2名）", "状態", "備考"]
    header(ws, 4, cols)
    widths(ws, [10, 6, 17, 7, 13, 10, 15, 14, 11, 11, 22, 18, 9, 22])

    r = 5
    for m, d in DATES:
        dt = datetime.date(2026, m, d)
        for si, s in enumerate(SLOTS):
            vals = ["%d/%d" % (m, d), WEEK[dt.weekday()], s["name"], s["meet"],
                    s["go"], s["arr"], s["back"], s["ret"], s["hold"], "", "", "", "", ""]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BOX
                if c in (1, 2, 4, 5, 6, 7, 8, 9, 10, 13):
                    cell.alignment = Alignment(horizontal="center")
                if si == 1:
                    cell.fill = SUB_FILL
            r += 1

    dv = DataValidation(type="list", formula1='"未定,候補,確定,中止"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("M5:M%d" % (r - 1))

    dv2 = DataValidation(type="list", formula1='"組1,組2,組3,組4,組5,組6,組7,組8"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add("J5:J%d" % (r - 1))

    ws.freeze_panes = "A5"
    ws.cell(row=r + 1, column=1,
            value="※ 便はすべて阪急バス 吹田市内線2系統。4便とも始発便なので車内0人から数えられます。").font = Font(size=9, color="555555")
    ws.cell(row=r + 2, column=1,
            value="※ 南口のりば2からは [2] と [3] の両方が「桃山台駅ゆき」で出ます。行先番号が 2 のバスに乗ること。").font = Font(size=9, color="C00000")
    ws.cell(row=r + 3, column=1,
            value="※ 桃山台では「[2] ＪＲ吹田駅（南口）ゆき」に乗り、終点の南口まで乗ります。「北口ゆき」と出ているのは [5] という別系統です。").font = Font(size=9, color="C00000")
    return ws


def sheet_dayflow(wb):
    ws = wb.create_sheet("当日の流れ")
    ws["A1"] = "当日の流れ"
    ws["A1"].font = Font(bold=True, size=14)
    widths(ws, [10, 44, 10, 46])

    rows = [
        ("", "", "", ""),
        ("■ 昼の日（空いている）", "", "■ 夕方の日（混んでいる）", ""),
        ("11:35", "ＪＲ吹田駅（南口）集合。説明・同意取得", "15:25", "ＪＲ吹田駅（南口）集合。説明（2日目の方は短縮）"),
        ("", "カードを番号順に3枚渡す（ID・パスワード入り）", "", "カードを渡す（2日目の方は1日目と同じ番号）"),
        ("", "その便のQRポスターを掲げ、全員ログインまで確認", "", "同左"),
        ("11:50", "のりば2へ移動。のりばで練習2回", "15:42", "のりば2へ移動。のりばで練習2回"),
        ("12:03", "1本目 発車 → 桃山台 12:48着。松本は離脱", "15:55", "1本目 発車 → 桃山台 16:40着。松本は離脱"),
        ("12:48", "折り返し待ち 60分", "16:40", "折り返し待ち 38分"),
        ("13:48", "2本目 発車（桃山台始発）", "17:18", "2本目 発車（北口まで実測39人の便）"),
        ("14:32", "ＪＲ吹田駅（南口）着・解散", "18:03", "ＪＲ吹田駅（南口）着・解散"),
        ("", "拘束 約3時間", "", "拘束 約2時間40分"),
        ("", "", "", ""),
        ("■ 持ち物", "", "", ""),
        ("", "配布カード（accounts_cards.html をA4に8面で印刷）", "", ""),
        ("", "QRポスター2枚（ボタンあり用／なし用）", "", ""),
        ("", "記録用紙（counting_sheet.html）1日4枚。22停留所版", "", ""),
        ("", "事業者の承諾を示す書面", "", ""),
        ("", "参加者への説明・同意書", "", ""),
        ("", "対応表 accounts.csv（人目に触れない場所に）", "", ""),
        ("", "", "", ""),
        ("■ 注意", "", "", ""),
        ("", "南口のりば2は [2] と [3] がどちらも「桃山台駅ゆき」。行先番号で見分ける", "", ""),
        ("", "松本はどちらの日も送り出しで離脱する（要求特性を避けるため）", "", ""),
        ("", "計数役2名は両方の便に乗る。参加者とは別々に乗車し、関わらない", "", ""),
        ("", "事後アンケートは全乗車を終えたあと", "", ""),
    ]
    for i, (a, b, c, d) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
        ws.cell(row=i, column=3, value=c)
        ws.cell(row=i, column=4, value=d)
        for col in (1, 3):
            if ws.cell(row=i, column=col).value and str(ws.cell(row=i, column=col).value).startswith("■"):
                ws.cell(row=i, column=col).font = Font(bold=True)
    return ws


def sheet_conditions(wb):
    ws = wb.create_sheet("便と条件")
    ws["A1"] = "どの便でどのQRポスターを掲げるか"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "同じ便に乗る3名は必ず同じ群にすること。群を混ぜると、同じバスの中で画面が違う人が並びます。"
    ws["A2"].font = Font(size=9, color="C00000")
    ws["A3"] = "入力方式は実施日ごとに固定（群1は昼＝ステッパー／夕方＝テンキー、群2はその逆）。ボタンの有無は便ごと。"
    ws["A3"].font = Font(size=9, color="555555")
    widths(ws, [16, 24, 24, 24, 24])

    header(ws, 4, ["", "昼の日 1本目（12:03）", "昼の日 2本目（13:48）",
                   "夕方の日 1本目（15:55）", "夕方の日 2本目（17:18）"])
    data = [("群1（組1・組3）", "A", "B", "D", "C"),
            ("群2（組2・組4）", "D", "C", "A", "B")]
    for i, row in enumerate(data, start=5):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = BOX
            cell.alignment = Alignment(horizontal="center")
            if c == 1:
                cell.font = Font(bold=True)

    ws["A8"] = "QRポスターと配布URL（docs/qr_posters.html をA4に4枚印刷。1日に使うのは2枚）"
    ws["A8"].font = Font(bold=True)
    base = "https://matsumototaichi-beep.github.io/bus-ui-study/"
    posters = [("A", "ステッパー ＋ ボタンあり", "?input=stepper"),
               ("B", "ステッパー ＋ ボタンなし", "?input=stepper&unknown=off"),
               ("C", "テンキー ＋ ボタンあり",   "?input=numpad"),
               ("D", "テンキー ＋ ボタンなし",   "?input=numpad&unknown=off")]
    for i, (mark, cond, qs) in enumerate(posters, start=9):
        ws.cell(row=i, column=1, value="ポスター " + mark).font = Font(bold=True)
        ws.cell(row=i, column=2, value=cond)
        ws.cell(row=i, column=3, value=base + qs)
    ws["A14"] = ("※ ポスターには条件の中身を書いていません。右上の A / B / C / D だけで見分けます。"
                 "条件の存在を参加者に気づかれると、条件そのものへの反応（要求特性）が入るためです。")
    ws["A14"].font = Font(size=9, color="C00000")
    ws["A15"] = ("※ 配布カードに刷ってあるQRはパラメータなしの素のURLです。"
                 "条件が付かないので、乗車時は必ずポスターのQRを読ませてください。")
    ws["A15"].font = Font(size=9, color="C00000")
    ws["A17"] = "この割り付けで、どの参加者も4乗車で「入力方式2 × ボタン2」の4通りを1回ずつ経験します。"
    ws["A17"].font = Font(size=9, color="555555")
    return ws


def sheet_preflight(wb):
    """当日までの準備。詳しい理由は docs/preflight_checklist.md にある。"""
    ws = wb.create_sheet("準備チェック")
    ws["A1"] = "実験当日までの準備"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "最速の実施日は10/13。詳しい理由と中身は docs/preflight_checklist.md を見てください。"
    ws["A2"].font = Font(size=9, color="555555")

    header(ws, 4, ["#", "区分", "やること", "なぜ要るか／どうやるか", "期限の目安", "状態", "備考"])
    widths(ws, [5, 14, 30, 52, 16, 10, 24])

    rows = [
        ("最優先", "参加者12名の確保と日程確定",
         "名簿と日程シートがまだ空。1便3名×4組、各組が昼と夕方の2日に出る＝実施8日", "今週中"),
        ("最優先", "参加者への説明・同意書を作る",
         "★実体がまだ無い。持ち物リストには載っているが文面が存在しない", "今週中"),
        ("最優先", "回答する8か所を決める",
         "区間が22停留所に伸びたので決め直し。同じ便の3人が同じ地点で答えないとICCが測れない", "今週中"),
        ("最優先", "阪急バス吹田営業所へ確定日を連絡",
         "承諾条件④。営業所から運転士へ連携されるので、実施の1週間前までに", "実施日が決まり次第"),
        ("最優先", "参加者アカウント24個を作る",
         "python tools/create_accounts.py（service_role キーが要る）", "1週間前"),
        ("最優先", "管理者アカウントを作り直す",
         "python tools/create_accounts.py --admin。いま admin でログインできない件の対処", "1週間前"),
        ("最優先", "配布カードを印刷",
         "accounts_cards.html をA4に8面×3枚", "1週間前"),
        ("最優先", "QRポスターを印刷",
         "docs/qr_posters.html をA4に4枚（A・B・C・D）。1日に使うのは2枚", "1週間前"),
        ("最優先", "記録用紙を印刷",
         "docs/counting_sheet.html を1日4枚。22停留所版", "1週間前"),
        ("最優先", "計数役2名の確保と事前練習",
         "★いちばん難しい役。本番前に1便、練習で乗ってもらう。終点で0に閉じる突き合わせまで通す", "1週間前"),
        ("直後に要る", "事後アンケートの実施方法を決める",
         "用紙（Word 4版）は作成済み。紙で配るか Google Forms にするかだけ", "随時"),
        ("直後に要る", "デブリーフィング文を書く",
         "アンケート後に渡す。「わからない」ボタンを比べていたことをここで明かす", "随時"),
        ("データが出てから", "分析スクリプト",
         "骨組みだけ先に作ると取り忘れに気づける。混合効果モデル・ICC・丸め率・欠損率", "随時"),
        ("待ち", "乗降計測専用アプリ",
         "★教授の返答待ち。A案（記録用紙に1欄）で足りるなら開発ゼロ。返事まで着手しない", "返答待ち"),
        ("待ち", "交通費の事務手続き",
         "謝礼なし・実費精算で確定済み。残るのは費目・証憑・事前申請の確認。総額 約1.8〜2.0万円", "随時"),
        ("前日", "Supabase の Confirm email を確認",
         "/auth/v1/settings の mailer_autoconfirm が true であること", "前日"),
        ("前日", "4つのURLからテスト送信",
         "松本のアカウントで、当日使うURL全部から実際に1件ずつ送れるか", "前日"),
        ("前日", "参加者へ事前連絡",
         "集合はＪＲ吹田駅（南口）。北口ではない。スマホの充電とICカード／小銭を持ってくる", "前日"),
    ]
    for i, (cat, what, why, due) in enumerate(rows, start=1):
        r = 4 + i
        for c, v in enumerate([i, cat, what, why, due, "", ""], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BOX
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (3, 4, 7)))
            if c in (1, 2, 5, 6):
                cell.alignment = Alignment(horizontal="center", vertical="top")
            if cat == "最優先":
                cell.font = Font(bold=(c == 3))
        ws.row_dimensions[r].height = 30

    dv = DataValidation(type="list", formula1='"未着手,着手,完了,不要になった"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("F5:F%d" % (4 + len(rows)))

    r = 4 + len(rows) + 2
    ws.cell(row=r, column=1, value="■ 要らなくなったもの（引きずらないこと）").font = Font(bold=True)
    for j, t in enumerate([
        "停留所ID割当表 … 要因Aを実施日ごとの固定に変えたので、条件が回答回数に依存しなくなった",
        "uiStudy-0.8 の条件割当の改修 … 同上。EXPERIMENT_PLAN §3.1 の「アプリの改修が必要」は解消済み",
        "アンケート用の画面4枚の撮影 … 写真を使わないアンケートに作り替えたため",
        "事後アンケート用紙の作成 … 作成済み（docs/questionnaire/アンケート用紙_v1〜v4.docx）",
        "アプリの改修全般 … uiStudy-0.8.1 で凍結。実験の途中で版が変わると条件が揃わない",
    ], start=1):
        ws.cell(row=r + j, column=1, value="・" + t).font = Font(size=9, color="555555")
    ws.freeze_panes = "A5"
    return ws


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "schedule.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    sheet_members(wb)
    sheet_dates(wb)
    sheet_dayflow(wb)
    sheet_conditions(wb)
    sheet_preflight(wb)
    wb.save(out)
    print("書き出しました: " + out)
    print("Excel で開いて自由に編集できます。日程が変わってもこのスクリプトを回し直す必要はありません")
    print("（回すと上書きされるので注意）。")
    print("\n★ 氏名・連絡先が入るので .gitignore 済みです。commit しないこと。")


if __name__ == "__main__":
    main()
