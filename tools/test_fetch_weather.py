# -*- coding: utf-8 -*-
"""tools/fetch_weather.py の確かめ（通信なし）

    python tools/test_fetch_weather.py

気象庁のページの読み方・10分の区切り・時刻の読み方・回答への付け方を、作った値で確かめる。
"""
import datetime as dt
import os
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fetch_weather as W

JST = W.JST
fail = 0


def check(ok, msg):
    global fail
    print(("OK  " if ok else "NG  ") + msg)
    fail += 0 if ok else 1


# 気象庁 10分値のページ（アメダス）の形をまねた表
PAGE = """<table><tr class="mtx"><th>時分</th></tr>
<tr class="mtx" style="text-align:right;"><td class="data_0_0">12:00</td><td class="data_0_0">0.0</td><td class="data_0_0">26.3</td>
<td>///</td><td>1.0</td><td>西南西</td><td>2.1</td><td>南西</td><td>///</td></tr>
<tr class="mtx"><td>12:10</td><td>1.5)</td><td>25.6</td><td>///</td><td>1.3</td><td>東南東</td><td>3.1</td><td>東</td><td>///</td></tr>
<tr class="mtx"><td>12:20</td><td>--</td><td>×</td><td>///</td><td>1.0</td><td>///</td><td>2.1</td><td>南</td><td>///</td></tr>
<tr class="mtx"><td>24:00</td><td>0.5]</td><td>20.0</td><td>///</td><td>0.5</td><td>北</td><td>1.0</td><td>北</td><td>///</td></tr>
</table>"""
day = dt.date(2026, 10, 13)
rows = W.parse_etrn(PAGE, day, "a")
check(len(rows) == 4, "10分値の行を4つ読む")
check(rows[1]["降水量_mm"] == 1.5 and rows[1]["気温_C"] == 25.6, "「1.5)」（準正常値）は 1.5 として読む")
check(rows[2]["降水量_mm"] == 0.0 and rows[2]["気温_C"] is None and rows[2]["風向"] == "", "「--」は 0、「×」「///」は欠測")
check(rows[3]["time"] == dt.datetime(2026, 10, 14, 0, 0, tzinfo=JST) and rows[3]["降水量_mm"] == 0.5, "「24:00」は次の日の 0:00、「0.5]」は 0.5")

# 10分の区切り（値の時刻は10分の終わり）
t = lambda h, m, s=0, us=0: dt.datetime(2026, 10, 13, h, m, s, us, tzinfo=JST)
check(W.bin10(t(12, 3)) == t(12, 10), "12:03 は 12:10 の値")
check(W.bin10(t(12, 10)) == t(12, 10), "12:10:00 は 12:10 の値")
check(W.bin10(t(12, 10, 1)) == t(12, 20), "12:10:01 は 12:20 の値")
check(W.bin10(t(12, 55)) == t(13, 0), "12:55 は 13:00 の値")
check(W.bin10(dt.datetime(2026, 10, 13, 3, 3, tzinfo=dt.timezone.utc)) == t(12, 10), "UTC の 03:03 は日本時間 12:10 の値")

# Supabase の書き出しの時刻
check(W.to_utc("2026-10-13 03:03:27.123+00") == dt.datetime(2026, 10, 13, 3, 3, 27, 123000, tzinfo=dt.timezone.utc), "「+00」で終わる時刻を読む")
check(W.to_utc("2026-10-13T03:03:27.1234567Z") is not None, "7桁の小数秒と Z を読む")
check(W.to_utc("") is None and W.to_utc(None) is None and W.to_utc("あ") is None, "空や読めない時刻は None")

# 回答に付ける
jma = {r["time"]: dict(r, 出どころ="確定") for r in rows}
jma[t(11, 50)] = {"time": t(11, 50), "降水量_mm": 0.0, "気温_C": 26.0, "風速_m/s": 1.0, "出どころ": "確定"}
om = {t(12, 0): {"天気": "晴れ", "雲量_%": 20, "湿度_%": 60, "降水量_mm": 0.0},
      t(13, 0): {"天気": "雨（弱）", "雲量_%": 90, "湿度_%": 80, "降水量_mm": 0.4}}
a = W.attach(t(12, 5), jma, om, "豊中")
check(a["豊中_10分の終わり"] == "2026-10-13 12:10" and a["豊中_降水量_10分_mm"] == 1.5, "12:05 の回答には 12:10 の10分値")
check(a["豊中_降水量_30分_mm"] == 1.5 and a["雨（実測）"] == "あり", "直前30分（11:50〜12:10）の雨を足す")
check(a["天気（Open-Meteo）"] == "晴れ" and a["降水量_1時間_mm（Open-Meteo）"] == 0.4, "天気はいちばん近い時刻、1時間の雨は回答を含む1時間")
b = W.attach(t(12, 45), jma, om, "豊中")
check(b["豊中_10分の終わり"] == "" and b["雨（実測）"] == "" and b["天気（Open-Meteo）"] == "雨（弱）", "10分値が無い時刻は空のまま（推し量らない）")

print("すべて OK" if not fail else "NG が %d 件" % fail)
sys.exit(1 if fail else 0)
