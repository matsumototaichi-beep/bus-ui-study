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
  6. 実施割当   … 全10日（本実験8＋計数方法2）の中身。日付だけ埋めれば完成
  7. アンケートURL … 版ごとにどのGoogleフォームを渡すか
     （URLは questionnaire_urls.txt から読む。.gitignore 済み。無ければ「未作成」と出る）

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
    """候補16日を昼枠・夕方枠に割った一覧。

    「回」の列に 1〜10 を入れると、組・入力方式・ポスター・参加者が
    「実施割当」シートから自動で引かれる。当日の取り違えを減らすため。
    """
    ws = wb.create_sheet("日程")
    ws["A1"] = "日程（候補16日 × 昼枠・夕方枠）"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("「回」に 1〜10 を入れると、右の列が「実施割当」シートから自動で入ります。"
                "本実験は回1〜8、計数方法の実験は回9〜10です。")
    ws["A2"].font = Font(size=9, color="555555")
    ws["A3"] = ("※ 回9・10（計数方法）は1日で昼と夕方の4便を回すので、同じ日の両方の行に入れてください。")
    ws["A3"].font = Font(size=9, color="555555")

    cols = ["日付", "曜日", "枠", "集合", "1本目 南口発", "桃山台着",
            "2本目 桃山台発", "南口着＝解散", "拘束",
            "回", "区分", "組", "入力方式", "1本目", "2本目", "参加者", "状態", "備考"]
    header(ws, 5, cols)
    widths(ws, [9, 5, 16, 7, 12, 10, 14, 13, 11,
                5, 13, 7, 11, 7, 7, 14, 9, 20])

    # 自動で引く列 → （この表の列番号, 実施割当シートの列番号）
    LOOK = [(11, 2), (12, 5), (13, 7), (14, 8), (15, 9), (16, 10)]

    r = 6
    for m, d in DATES:
        dt = datetime.date(2026, m, d)
        for si, s_ in enumerate(SLOTS):
            vals = ["%d/%d" % (m, d), WEEK[dt.weekday()], s_["name"], s_["meet"],
                    s_["go"], s_["arr"], s_["back"], s_["ret"], s_["hold"],
                    "", "", "", "", "", "", "", "", ""]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BOX
                if c in (1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 17):
                    cell.alignment = Alignment(horizontal="center")
                if si == 1:
                    cell.fill = SUB_FILL
            for here, there in LOOK:
                ws.cell(row=r, column=here).value = (
                    '=IFERROR(VLOOKUP($J%d,実施割当!$A$5:$L$14,%d,FALSE),"")' % (r, there))
            r += 1

    dv = DataValidation(type="list", formula1='"1,2,3,4,5,6,7,8,9,10"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("J6:J%d" % (r - 1))

    dv2 = DataValidation(type="list", formula1='"未定,候補,確定,実施済,中止"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add("Q6:Q%d" % (r - 1))

    ws.freeze_panes = "D6"
    ws.cell(row=r + 1, column=1,
            value="※ 便はすべて阪急バス 吹田市内線2系統。4便とも始発便なので車内0人から数えられます。").font = Font(size=9, color="555555")
    ws.cell(row=r + 2, column=1,
            value="※ 南口のりば2からは [2] と [3] の両方が「桃山台駅ゆき」で出ます。行先番号が 2 のバスに乗ること。").font = Font(size=9, color="C00000")
    ws.cell(row=r + 3, column=1,
            value="※ 桃山台では「[2] ＪＲ吹田駅（南口）ゆき」に乗り、終点の南口まで乗ります。「北口ゆき」と出ているのは [5] という別系統です。").font = Font(size=9, color="C00000")
    ws.cell(row=r + 5, column=1, value="■ 土日は使えません").font = Font(bold=True)
    ws.cell(row=r + 6, column=1,
            value="10/24・25、11/14・15 は土日。桃山台17時台が土休日は 北10/20/北40/50 で 17:18 が無く、夕方の2本目が組めません。").font = Font(size=9, color="555555")
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


def sheet_plan(wb):
    """全10日の割り当て。日付以外はすべて決まっているので、候補16日から埋めるだけ。

    本実験8日（昼4・夕方4）＋ 計数方法の実験2日。候補16日なので予備が6日残る。
    条件の組み合わせは docs/operation_plan.md §4 と同じ。
    """
    ws = wb.create_sheet("実施割当")
    ws["A1"] = "実施割当（全10日）"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("日付以外はすべて決まっています。候補16日の中から日付を入れてください。"
                "本実験8日＋計数方法の実験2日で、予備が6日残ります。")
    ws["A2"].font = Font(size=9, color="555555")

    header(ws, 4, ["回", "区分", "枠", "日付", "組", "群", "入力方式",
                   "1本目のポスター", "2本目のポスター", "参加者", "状態", "備考"])
    widths(ws, [5, 14, 16, 11, 7, 7, 12, 15, 15, 20, 9, 22])

    # 組1・組3 は昼＝ステッパー、組2・組4 は昼＝テンキー（入力方式と混雑を釣り合わせる）
    # ポスター A=ステッパー+あり B=ステッパー+なし C=テンキー+あり D=テンキー+なし
    main = []
    for team in (1, 2, 3, 4):
        gun = 1 if team % 2 == 1 else 2
        if gun == 1:
            noon = ("ステッパー", "A", "B")
            eve = ("テンキー", "D", "C")
        else:
            noon = ("テンキー", "D", "C")
            eve = ("ステッパー", "A", "B")
        who = "p%02d〜p%02d" % (team * 3 - 2, team * 3)
        main.append(("本実験", "昼（空いている）", "組%d" % team, "群%d" % gun) + noon + (who,))
        main.append(("本実験", "夕方（混んでいる）", "組%d" % team, "群%d" % gun) + eve + (who,))

    r = 5
    for i, row in enumerate(main, start=1):
        kubun, waku, team, gun, hoshiki, p1, p2, who = row
        vals = [i, kubun, waku, "", team, gun, hoshiki, p1, p2, who, "", ""]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BOX
            if c in (1, 3, 5, 6, 7, 8, 9, 11):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if "夕方" in waku:
                cell.fill = SUB_FILL
        r += 1

    for i in (9, 10):
        vals = [i, "計数方法の実験", "昼＋夕方（4便）", "", "―", "―", "―", "―", "―",
                "参加者なし・計数役3名", "", "足し引き2名＋数え直し1名。乗降計測アプリを使う"]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BOX
            cell.fill = PatternFill("solid", fgColor="EAF3EA")
            if c in (1, 3, 5, 6, 7, 8, 9, 11):
                cell.alignment = Alignment(horizontal="center", vertical="center")
        r += 1

    dv = DataValidation(type="list", formula1='"未定,候補,確定,実施済,中止"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("K5:K%d" % (r - 1))

    r += 1
    for t, col in [
        ("■ 読み方", "000000"),
        ("・同じ組の2日（昼と夕方）は、同じ3名が別の日に来ます。間が空いても構いません", "555555"),
        ("・入力方式は1日を通して変わりません。組1・組3は昼がステッパー、組2・組4は昼がテンキー", "555555"),
        ("・ポスター A＝ステッパー＋ボタンあり／B＝ステッパー＋ボタンなし／"
         "C＝テンキー＋ボタンあり／D＝テンキー＋ボタンなし", "555555"),
        ("・これでどの参加者も4乗車で「入力方式2通り × ボタン2通り」を1回ずつ経験します", "555555"),
        ("", "555555"),
        ("■ 日数", "000000"),
        ("・本実験8日 ＋ 計数方法の実験2日 ＝ 10日。候補16日なので予備が6日", "555555"),
        ("・混雑時のクラス回答は、全回答でクラスを記録しているので追加の実験日は要りません", "555555"),
        ("", "555555"),
        ("■ 1日に昼と夕方の両方を入れることもできます", "000000"),
        ("・昼は11:35〜14:32、夕方は15:25〜18:03で重なりません。別の組を入れれば1日で2枠こなせます", "555555"),
        ("・そうすると本実験は4日に縮みますが、計数役2名が1日4便（約6時間半）乗ることになります", "555555"),
        ("・基準値の質が落ちると研究の土台が崩れるので、1日1枠（8日）を勧めます", "C00000"),
    ]:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=(10 if t.startswith("■") else 9), bold=t.startswith("■"), color=col)
        r += 1

    ws.freeze_panes = "A5"
    return ws


def sheet_forms(wb):
    """版ごとに、どのGoogleフォームのURLを渡すか。当日これを見れば迷わない。

    URLは questionnaire_urls.txt（.gitignore 済み）から読む。
    回答用リンクが公開の場所にあると、知らない人の回答が混ざるため、
    このリポジトリには入れない。
    """
    ws = wb.create_sheet("アンケートURL")
    ws["A1"] = "事後アンケート（Googleフォーム）"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("参加者の版は配布カードに刷ってあります。その版の行のURLを渡してください。"
                "①前半を送信してもらってから①後半のURLを渡すこと。")
    ws["A2"].font = Font(size=9, color="555555")

    urls = {}
    path = os.path.join(ROOT, "questionnaire_urls.txt")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.split("|")
                if len(parts) >= 2:
                    urls[parts[0]] = parts[1]
    if not urls:
        ws["A4"] = "★ questionnaire_urls.txt がありません。フォームを作ってからURLを書いてください。"
        ws["A4"].font = Font(bold=True, color="C00000")
        return ws

    header(ws, 4, ["版", "① 前半（その日の解散直後・全員）",
                   "② 後半（①を送信してから渡す）"])
    widths(ws, [8, 56, 56])
    plan = [("v1", "後半A"), ("v2", "後半A"), ("v3", "後半B"), ("v4", "後半B")]
    for i, (ver, kouhan) in enumerate(plan, start=5):
        for c, v in enumerate([ver, urls.get("前半", ""), urls.get(kouhan, "")], start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = BOX
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c == 1:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
        ws.row_dimensions[i].height = 30

    ws["A10"] = "参加者の番号と版の対応"
    ws["A10"].font = Font(bold=True)
    ws["A11"] = "p01 p05 p09 → v1　／　p02 p06 p10 → v2　／　p03 p07 p11 → v3　／　p04 p08 p12 → v4"

    ws["A13"] = "渡す順番"
    ws["A13"].font = Font(bold=True)
    for j, t in enumerate([
        "1日目も2日目も同じ … ①前半 → 送信を確認 → ②後半",
        "②後半の最後に「★2日目の方だけ」の区切りがあり、1日目の人はそこで終わる",
        "★②後半を先に見せないこと。「わからない」ボタンの話が先に出ると、"
        "①前半のS3（今日どう数えたか）が、ボタンを意識した答えになってしまう",
    ], start=14):
        c = ws.cell(row=j, column=1, value="・" + t)
        c.font = Font(size=9, color=("C00000" if t.startswith("★") else "555555"))

    ws["A18"] = "フォーム側で確認しておくこと"
    ws["A18"].font = Font(bold=True)
    for j, t in enumerate([
        "回答を1回に制限しない（1人が1日目と2日目で2回答えるため）",
        "①前半と②後半は「今日は何日目ですか」で1日目/2日目を区別する",
        "各フォームの「回答」タブからスプレッドシートに出力しておくと集計が楽",
        "リンクをSNSやリポジトリなど公開の場所に貼らない（知らない人の回答が混ざる）",
    ], start=19):
        c = ws.cell(row=j, column=1, value="・" + t)
        c.font = Font(size=9, color="555555")
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
    sheet_plan(wb)
    sheet_forms(wb)
    wb.save(out)
    print("書き出しました: " + out)
    print("Excel で開いて自由に編集できます。日程が変わってもこのスクリプトを回し直す必要はありません")
    print("（回すと上書きされるので注意）。")
    print("\n★ 氏名・連絡先が入るので .gitignore 済みです。commit しないこと。")


if __name__ == "__main__":
    main()
