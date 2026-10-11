# -*- coding: utf-8 -*-
"""join_survey_app.py の確かめ（偽のデータ）

    python tools/test_join_survey_app.py [--out フォルダ]

偽の参加者 A（奇数組）と B（偶数組）が1回目・2回目に来たときの、アプリの記録（ui_sessions・ui_responses）と
アンケート4シートの CSV を作り、join_survey_app.py に通して、出てきた表と問題の一覧を確かめる。
問題の例を1つずつ入れてある（下の SCENARIO）。出力はリポジトリの外（--out。省くと一時フォルダ）に出す。

SCENARIO
  A 1回目 10/13 昼  練習2回・再送で同じ回答が2行・操作ログのない回答・B のログインで送った回答・
                    南口のアンケートが3行（桃山台駅で答えた行・正しい行・もう1度送った行）
  B 1回目 10/13 昼  「わからない」で送った回答・かえりで入れ方がまざっている・南口のアンケートを桃山台駅で答えた
                    問1 を全角と空白で「ｂ 」と書いた
  C       10/13 昼  アプリの記録がなく、--ids の名簿にもない ID のアンケート
  D 1回目 10/13 昼  バスが停まらず10分あいた間に、降りる前に桃山台のアンケートに答えた
  admin   10/13     確認用アカウント
  A       10/12     実験の前のテスト（--since 2026-10-13 で読まない）
  B 2回目 10/14 昼  ?round のない画面で1回答・かえりに切り替え忘れ・桃山台駅で1回目のフォームに答えた
  A 2回目 10/15 夕方 端末の時計が5分進んでいる・かえりだけ古い版・桃山台のアンケートがない・SUS が全部3・
                    ゆきの1回答だけ別の端末・かえりで人数の途中でページが閉じた
  E 1回目 10/16 昼  ゆきの途中でログアウトし、?round のない画面で入り直した（そのあとは回・区間のない記録）
  F 1回目 10/16 昼  ふつう
  F 2回目 10/17 昼  かえりの途中で、前の日の ?round=1 のタブが読み込み直された（残り2回答が1回目かえりの記録）
  H 1回目 10/16 夕方 桃山台で切り替えを忘れ、かえりのバスの2回答のあとで切り替えた
  J 1回目 10/18 昼  南口で練習したあと練習を切り忘れ、田中町を出たときの回答も練習のまま送った
  L 1回目 10/18 夕方 桃山台のアンケートの代わりに、桃山台駅で南口の QR を読んで答えた（南口でも答えた）
  M 1回目 10/19 昼  ?round のない画面で1回目を全部答えた（ボタンはいつも出る）。2回目 10/20 はふつう
  N 1回目 10/21 昼  朝（08:30）に練習だけ試した。往復はふつう
  N 2回目 10/22 昼  南口で、前の日の ?round=1 のタブのまま練習してから、2回目の QR を読んだ
  1回目_南口.csv は列の順を入れ替え、古い「問5」の列を足し、問2 の見出しの文を変えた
"""
import argparse
import json
import os
import re
import sys
import tempfile
import uuid

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import join_survey_app as J

UUID = {k: "uuid-" + k.lower() for k in ("A", "B", "D", "E", "F", "H", "J", "L", "M", "N", "ADMIN")}
GROUP = {"A": "odd", "B": "even", "D": "even", "E": "odd", "F": "even", "H": "even", "J": "even", "L": "even", "M": "odd",
         "N": "even", "ADMIN": "odd"}
SESS, RESP = [], []
SEQ = [0]


def T(day, hm, sec=0):
    return pd.Timestamp("2026-10-%02d %s:%02d" % (day, hm, sec), tz="Asia/Tokyo")


def iso(t):
    return t.tz_convert("UTC").strftime("%Y-%m-%dT%H:%M:%S.000Z")


def answer(login, rnd, leg, t, inp, unk, exact=12, cls=3, practice=False, outcome="submit", version="uiStudy-0.9.5",
           lat=None, dev_ahead_s=0, sender=None, login_v=None, pressed_unknown=False, no_session=False, resend=False,
           test=False, study="ui", leg_by=None, leg_at=None, device=None, reason=None, lng=None, url_study=None):
    """1回答ぶんの ui_sessions と ui_responses の行を足す。t は本当の時刻、端末の時計は dev_ahead_s 秒進んでいる"""
    SEQ[0] += 1
    sid, cid = str(uuid.uuid4()), str(uuid.uuid4())
    dev = t + pd.Timedelta(seconds=dev_ahead_s)
    sent = t + pd.Timedelta(seconds=20)
    who = UUID[sender or login]
    v = {"input": inp, "count_init": "none", "seq": SEQ[0], "offset": 0, "unknown_ui": unk,
         "url_fixed": {"input": None, "unknown": None}, "practice": practice}
    if study == "ui" and rnd:
        v.update({"study": "ui", "round": rnd, "leg": leg, "group": GROUP[login], "test_account": test,
                  "label": "%d回目%s" % (rnd, J.LEG_JA[leg])})
    else:
        v.update({"study": None, "round": None, "leg": None, "group": None, "label": None})
    if login_v:
        v["login_id"] = login_v
    if leg_by:                                       # uiStudy-0.9.6 から
        v.update({"leg_by": leg_by, "leg_at": iso(leg_at)})
    if url_study:
        v["url_study"] = url_study
    ev = [{"dt": 0, "type": "session_start"}, {"dt": 5, "type": "screen_view", "screen": "count", "input": inp,
                                                 "unknown_ui": unk}]
    if pressed_unknown:
        ev.append({"dt": 900, "type": "count_unknown", "value_before": None})
    ev.append({"dt": 1500, "type": "session_end", "outcome": outcome})
    if not no_session:
        SESS.append({"id": len(SESS) + 1, "created_at": iso(sent), "surveyor_id": who, "session_id": sid,
                     "started_at": iso(dev), "client_submitted_at": iso(sent + pd.Timedelta(seconds=dev_ahead_s)),
                     "outcome": outcome, "variant": v,
                     "context": {"congestion_class": cls, "exact_count": exact, "count_unknown": pressed_unknown,
                                 "count_touched": exact is not None, "app_version": version, "practice": practice,
                                 "response_cid": cid, "gps_points": 3, "device_id": device or "dev-" + login.lower(),
                                 **({"reason": reason} if reason else {})},
                     "summary": {"t_total": 6000, "t_class": 1500, "t_count": 4000, "t_first": 900, "t_stop": 0,
                                 "taps_total": 5, "taps_value": 3, "corrections": 0},
                     "events": ev})
    if outcome in ("submit", "skip") or reason == "pagehide":
        r = {"id": len(RESP) + 1, "created_at": iso(sent), "surveyor_id": who, "answered_at": iso(dev),
             "congestion_class": cls, "exact_count": exact, "practice": practice, "app_version": version,
             "lat": lat, "lng": lng if lng is not None else (135.51 if lat else None), "session_id": None if no_session else sid, "client_id": cid,
             "client_submitted_at": iso(sent + pd.Timedelta(seconds=dev_ahead_s))}
        RESP.append(r)
        if resend:                                   # 返事が届かず、同じ回答をもう1度送った
            r2 = dict(r, id=len(RESP) + 1, created_at=iso(sent + pd.Timedelta(seconds=90)))
            RESP.append(r2)


YUKI_T = ["11:52", "11:56", "12:00", "12:04"]
KAERI_T = ["12:22", "12:26", "12:30", "12:34"]
YUKI_LAT = [34.765, 34.772, 34.780, 34.788]
KAERI_LAT = [34.790, 34.782, 34.774, 34.766]


def build_app():
    # 実験の前の日のテスト（--since 2026-10-13 で読まない）
    answer("A", 1, "yuki", T(12, "15:00"), "numpad", False)
    # admin（確認用）
    answer("ADMIN", 1, "yuki", T(13, "11:30"), "numpad", False, test=True)
    answer("ADMIN", 1, "yuki", T(13, "11:31"), "numpad", False, test=True)
    # A 1回目 10/13：ゆき＝数字キー・ボタンなし、かえり＝＋−・ボタンなし
    answer("A", 1, "yuki", T(13, "11:45"), "numpad", False, practice=True)
    answer("A", 1, "yuki", T(13, "11:46"), "numpad", False, practice=True)
    for i, (hm, lat) in enumerate(zip(YUKI_T, YUKI_LAT)):
        answer("A", 1, "yuki", T(13, hm), "numpad", False, exact=[10, 13, 15, 19][i], lat=lat, resend=(i == 1))
    for i, (hm, lat) in enumerate(zip(KAERI_T, KAERI_LAT)):
        answer("A", 1, "kaeri", T(13, hm), "stepper", False, exact=20 - i, lat=lat,
               sender="B" if i == 2 else None, login_v="A" if i == 2 else None, leg_by="switch", leg_at=T(13, "12:13"))
    answer("A", 1, "kaeri", T(13, "12:32"), "stepper", False, exact=None, lat=34.770, no_session=True)
    # B 1回目 10/13：ゆき＝数字キー・ボタンあり、かえり＝＋−・ボタンあり（1つだけ数字キー）
    for i, (hm, lat) in enumerate(zip(YUKI_T, YUKI_LAT)):
        answer("B", 1, "yuki", T(13, hm, 10), "numpad", True, lat=lat,
               exact=None if i == 3 else 12, pressed_unknown=(i == 3))
    for i, (hm, lat) in enumerate(zip(KAERI_T, KAERI_LAT)):
        answer("B", 1, "kaeri", T(13, hm, 10), "numpad" if i == 1 else "stepper", True, lat=lat)
    # D 1回目 10/13：バスが停まらず10分あいた間に、桃山台のアンケートに答えた（切り替え忘れではない）
    for hm, lat in zip(["11:52", "11:56", "12:06"], [34.765, 34.772, 34.788]):
        answer("D", 1, "yuki", T(13, hm, 40), "numpad", True, lat=lat)
    for hm, lat in zip(KAERI_T, KAERI_LAT):
        answer("D", 1, "kaeri", T(13, hm, 40), "stepper", True, lat=lat)
    # B 2回目 10/14：?round なしで1回答。ゆき＝＋−・ボタンなし。かえりに切り替え忘れ（かえりの回答も yuki）
    answer("B", None, None, T(14, "11:40"), "stepper", True, study=None)
    for hm, lat in list(zip(YUKI_T, YUKI_LAT)) + list(zip(KAERI_T, KAERI_LAT)):
        answer("B", 2, "yuki", T(14, hm), "stepper", False, lat=lat)
    # A 2回目 10/15 夕方：端末の時計が5分進んでいる。ゆき＝＋−・ボタンあり、かえり＝数字キー・ボタンあり（古い版）
    for i, (hm, lat) in enumerate(zip(["15:52", "15:56", "16:00", "16:04"], YUKI_LAT)):
        answer("A", 2, "yuki", T(15, hm), "stepper", True, lat=lat, dev_ahead_s=300, device="dev-x" if i == 3 else None)
    for hm, lat in zip(["16:22", "16:26", "16:30", "16:34"], KAERI_LAT):
        answer("A", 2, "kaeri", T(15, hm), "numpad", True, lat=lat, dev_ahead_s=300, version="uiStudy-0.9.4")
    answer("A", 2, "kaeri", T(15, "16:28"), "numpad", True, exact=None, lat=34.778, dev_ahead_s=300,
           version="uiStudy-0.9.4", outcome="abandon", reason="pagehide")             # 人数の途中でページが閉じた
    V6 = dict(version="uiStudy-0.9.6")
    # E 1回目 10/16：ゆきの2回答のあと、ログアウトして ?round のない画面で入り直した（入れ方は回答ごとに替わり、ボタンが出る）
    for hm in YUKI_T[:2]:
        answer("E", 1, "yuki", T(16, hm), "numpad", False, **V6)
    for i, hm in enumerate(YUKI_T[2:] + KAERI_T):
        answer("E", None, None, T(16, hm), ["stepper", "numpad"][i % 2], True, study=None, **V6)
    # F 1回目 10/16（ふつう）・2回目 10/17：かえりの2回答のあと、前の日の ?round=1 のタブが読み込み直された
    for hm in YUKI_T:
        answer("F", 1, "yuki", T(16, hm, 5), "numpad", True, **V6)
    for hm in KAERI_T:
        answer("F", 1, "kaeri", T(16, hm, 5), "stepper", True, **V6)
    for hm in YUKI_T:
        answer("F", 2, "yuki", T(17, hm), "stepper", False, url_study="r2", **V6)
    for hm in KAERI_T[:2]:
        answer("F", 2, "kaeri", T(17, hm), "numpad", False, url_study="r2", leg_by="switch", leg_at=T(17, "12:15"), **V6)
    for hm in KAERI_T[2:]:
        answer("F", 1, "kaeri", T(17, hm), "stepper", True, url_study="r1", leg_by="pick", leg_at=T(17, "12:28"), **V6)
    # H 1回目 10/16 夕方：桃山台のアンケート（16:08）のあと、かえりのバスの2回答（16:16・16:19）を「ゆき」のまま送り、16:20 に切り替えた
    for hm in ["15:52", "15:56", "16:00", "16:04", "16:16", "16:19"]:
        answer("H", 1, "yuki", T(16, hm), "numpad", True, leg_by="pick", leg_at=T(16, "15:45"), **V6)
    for hm in ["16:22", "16:26", "16:30"]:
        answer("H", 1, "kaeri", T(16, hm), "stepper", True, leg_by="switch", leg_at=T(16, "16:20"), **V6)
    # J 1回目 10/18：南口のバス停で練習（11:40）。練習を切り忘れ、田中町を出たときの回答（11:52）も練習のまま送った
    answer("J", 1, "yuki", T(18, "11:40"), "numpad", True, practice=True, lat=34.76290, lng=135.52440, **V6)
    answer("J", 1, "yuki", T(18, "11:52"), "numpad", True, practice=True, lat=34.75984, lng=135.52007, **V6)
    for hm, lat in zip(YUKI_T[1:], YUKI_LAT[1:]):
        answer("J", 1, "yuki", T(18, hm), "numpad", True, lat=lat, **V6)
    # L 1回目 10/18 夕方：桃山台のアンケートはなく、桃山台駅（16:10）と南口（16:40）で南口のアンケートに答えた
    for hm in ["15:52", "15:56", "16:00", "16:04"]:
        answer("L", 1, "yuki", T(18, hm), "numpad", True, **V6)
    for hm in ["16:22", "16:26", "16:30", "16:34"]:
        answer("L", 1, "kaeri", T(18, hm), "stepper", True, **V6)
    # M 1回目 10/19：?round のない画面で全部答えた（ボタンはいつも出る）。2回目 10/20 はふつう（奇数組なのでボタンあり）
    for i, hm in enumerate(YUKI_T + KAERI_T):
        answer("M", None, None, T(19, hm), ["numpad", "stepper"][i % 2], True, study=None, **V6)
    for hm in YUKI_T:
        answer("M", 2, "yuki", T(20, hm), "stepper", True, **V6)
    for hm in KAERI_T:
        answer("M", 2, "kaeri", T(20, hm), "numpad", True, **V6)
    # N 1回目 10/21：朝 08:30 に練習だけ試した（往復の3時間前）。2回目 10/22：前の日のタブ（?round=1）で練習してから QR を読んだ
    answer("N", 1, "yuki", T(21, "08:30"), "numpad", True, practice=True, **V6)
    for d, rnd, yi, ki, unk in ((21, 1, "numpad", "stepper", True), (22, 2, "stepper", "numpad", False)):
        for hm in YUKI_T:
            answer("N", rnd, "yuki", T(d, hm), yi, unk, **V6)
        for hm in KAERI_T:
            answer("N", rnd, "kaeri", T(d, hm), ki, unk, **V6)
    answer("N", 1, "yuki", T(22, "11:40"), "numpad", True, practice=True, url_study="r1", **V6)


def sheet(name, rows, qmap, shuffle=False, extra=None, rename=None):
    """フォームの列見出しどおりの CSV の行を作る。rows は {"ts":…, "問N": 値}"""
    heads = ["タイムスタンプ"] + [q["title"] for q in qmap[name]]
    recs = []
    for r in rows:
        d = {"タイムスタンプ": r["ts"].strftime("%Y/%m/%d %H:%M:%S")}
        for q in qmap[name]:
            d[q["title"]] = str(r.get(q["n"], ""))
        recs.append(d)
    df = pd.DataFrame(recs, columns=heads)
    if rename:
        df = df.rename(columns=rename)
    if extra:
        for h, v in extra.items():
            df[h] = v
    if shuffle:
        cols = list(df.columns)
        df = df[[cols[0]] + list(reversed(cols[1:]))]
    return df


def south(id_, ts, base=None, over=None):
    """南口のアンケートの1行。問番号 → 答え。SUS（問7〜16）は既定で 4,2,5,1,4,2,5,1,4,2。
    1回目_南口 の自由記述は問18（2回目_南口 では問18 はボタンの問なので over で上書きする）"""
    r = {"ts": ts, 1: id_, 2: 3, 3: 2, 4: 1, 5: 4, 6: 4, 17: "行きで利用したアプリ", 18: "なし"}
    for i, x in enumerate(base or [4, 2, 5, 1, 4, 2, 5, 1, 4, 2]):
        r[7 + i] = x
    r.update(over or {})
    return r


def north(id_, ts):
    return {"ts": ts, 1: id_, 2: 2, 3: 1, 4: 1, 5: 5, 6: 4, 7: "なし"}


def build_forms(folder, qmap):
    os.makedirs(folder, exist_ok=True)
    m1 = [north("A", T(12, "15:05")),                                        # 実験の前の日のテスト
          north("A", T(13, "12:10", 30)), north("ｂ ", T(13, "12:11")), north("C", T(13, "12:12")),
          north("B", T(14, "12:10")),                                         # B は2回目なのに1回目のフォーム
          north("D", T(13, "12:01")),                                         # D は降りる前に答えた
          north("E", T(16, "12:10")), north("F", T(16, "12:10", 20)), north("H", T(16, "16:08")),
          north("M", T(19, "12:10"))]
    s1 = [south("A", T(13, "12:15")),                                        # 桃山台駅で南口の QR を読んだ
          south("A", T(13, "12:40")),                                        # 正しい行（これを使う）
          south("A", T(13, "12:41", 30), base=[3, 3, 4, 2, 4, 2, 5, 1, 4, 2]),  # もう1度送った（答えが違う）
          south("B", T(13, "12:15", 20), over={17: "帰りで利用したアプリ"}),   # 桃山台駅で答えた（あとの行はない）
          south("E", T(16, "12:40")), south("F", T(16, "12:40", 20)), south("H", T(16, "16:36")),
          south("L", T(18, "16:10")),                                        # 桃山台駅で南口の QR を読んだ
          south("L", T(18, "16:40")), south("M", T(19, "12:40"))]
    s2 = [south("A", T(15, "16:36"), base=[3] * 10,
                over={3: 2, 4: 3, 17: "帰りで利用したアプリ", 18: "2回目", 19: "2回目", 20: 4, 21: "はい", 22: "なし"}),
          south("B", T(14, "12:40"), over={18: "1回目", 19: "変わらない", 20: 2, 21: "いいえ", 22: "とくになし"}),
          south("F", T(17, "12:40"), over={18: "1回目", 19: "1回目", 20: 2, 21: "はい", 22: "なし"}),
          south("M", T(20, "12:40"), over={18: "2回目", 19: "2回目", 20: 4, 21: "はい", 22: "なし"})]
    sheet("1回目_桃山台", m1, qmap).to_csv(os.path.join(folder, "1回目_桃山台.csv"), index=False, encoding="utf-8")
    t2 = qmap["1回目_南口"][1]["title"]
    sheet("1回目_南口", s1, qmap, shuffle=True, extra={"問5 その区間で使った入れ方の答えやすさ": "3"},
          rename={t2: t2 + "（5段階）"}).to_csv(
        os.path.join(folder, "UI研究 アンケート回答 - 1回目_南口.csv"), index=False, encoding="utf-8")
    sheet("2回目_桃山台", [north("F", T(17, "12:10")), north("M", T(20, "12:10"))], qmap).to_csv(os.path.join(folder, "2回目_桃山台.csv"), index=False, encoding="utf-8")
    sheet("2回目_南口", s2, qmap).to_csv(os.path.join(folder, "2回目_南口.csv"), index=False, encoding="utf-8")


# ---------------------------------------------------------------- 確かめ
FAIL = []


def check(cond, msg):
    print(("  OK  " if cond else "  NG  ") + msg)
    if not cond:
        FAIL.append(msg)


def row(t, id_, rnd, leg):
    x = t[(t["ID"] == id_) & (t["回"] == rnd) & (t["区間"] == leg)]
    if len(x) != 1:
        FAIL.append("行がちょうど1つではない: %s %s %s（%d 行）" % (id_, rnd, leg, len(x)))
        return None
    return x.iloc[0]


def has(r, code):
    return r is not None and isinstance(r["問題"], str) and code in r["問題"].split("、")


def sus_by_hand(ans):
    odd = sum(ans[i] - 1 for i in (0, 2, 4, 6, 8))
    even = sum(5 - ans[i] for i in (1, 3, 5, 7, 9))
    return (odd + even) * 2.5


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", help="出力フォルダ（リポジトリの外）")
    a = ap.parse_args()
    out = a.out or tempfile.mkdtemp(prefix="join_survey_test_")
    if os.path.abspath(out).startswith(J.ROOT + os.sep) and "exports" not in out:
        sys.exit("テストの出力はリポジトリの外に出してください: %s" % out)
    os.makedirs(out, exist_ok=True)

    # 組の表が index.html と同じか
    s = open(os.path.join(J.ROOT, "index.html"), encoding="utf-8").read()
    odd = re.search(r'"([A-Z]+)"\.split\(""\)\.forEach\(c=>\{ GROUP_OF\[c\]="odd"', s).group(1)
    even = re.search(r'"([A-Z]+)"\.split\(""\)\.forEach\(c=>\{ GROUP_OF\[c\]="even"', s).group(1)
    check({c: "odd" for c in odd} | {c: "even" for c in even} == J.GROUP_OF, "組の表（GROUP_OF）が index.html と同じ")

    with open(J.latest_form_json(), encoding="utf-8") as f:
        qmap = J.question_map(json.load(f))
    counts = {k: len(v) for k, v in qmap.items()}
    check(counts == {"1回目_桃山台": 7, "1回目_南口": 18, "2回目_桃山台": 7, "2回目_南口": 22}, "問数 7・18・7・22：%s" % counts)
    check(all(q["key"] for qs in qmap.values() for q in qs), "どの問にも表の列名が決まっている")
    sus_n = [q["n"] for q in qmap["1回目_南口"] if q["sus"]]
    check(sus_n == list(range(7, 17)), "SUS は南口の問7〜16：%s" % sus_n)

    build_app()
    forms = os.path.join(out, "forms")
    build_forms(forms, qmap)
    with open(os.path.join(out, "ui_sessions.json"), "w", encoding="utf-8") as f:
        json.dump(SESS, f, ensure_ascii=False)
    rd = pd.DataFrame(RESP)
    rd["practice"] = rd["practice"].map(lambda x: "true" if x else "false")     # Supabase の CSV と同じ書き方
    rd.to_csv(os.path.join(out, "ui_responses.csv"), index=False, encoding="utf-8")
    pd.DataFrame([{"surveyor_id": v, "login_id": k.lower()} for k, v in UUID.items()]).to_csv(
        os.path.join(out, "id対応.csv"), index=False, encoding="utf-8")

    t, iss = J.main(["--forms", forms, "--sessions", os.path.join(out, "ui_sessions.json"),
                     "--responses", os.path.join(out, "ui_responses.csv"), "--idmap", os.path.join(out, "id対応.csv"),
                     "--out", os.path.join(out, "result"), "--ids", "A,B,D,E,F,H,J,L,M,N", "--app-version", "uiStudy-0.9.5,uiStudy-0.9.6",
                     "--since", "2026-10-13"])
    t = pd.read_csv(os.path.join(out, "result", "区間ごと.csv"), encoding="utf-8-sig", dtype=str, keep_default_na=False)
    t["回"] = t["回"].astype(int)
    t["問題"] = t["問題"].where(t["問題"] != "", None)
    iss = pd.read_csv(os.path.join(out, "result", "問題の一覧.csv"), encoding="utf-8-sig", dtype=str, keep_default_na=False)
    sv = pd.read_csv(os.path.join(out, "result", "アンケートの行.csv"), encoding="utf-8-sig", dtype=str, keep_default_na=False)
    print()
    print("区間ごとの行：")
    for _, r in t.iterrows():
        print("  %s %s %s %s 入れ方=%s ボタン=%s 回答=%s フォーム=%s SUS=%s 問題=%s 外す=%s" % (
            r["ID"], r["回"], r["区間"], r["日付"], r["入れ方"], r["わからないボタン"], r["本番の回答数"], r["フォーム"],
            r["SUS点"], r["問題"], r["主な分析から外す理由"]))
    print()

    check(len(t) == 31, "区間ごとの行は31行（A 4・B 4・D 2・C 1・admin 1・E 2・F 4・H 2・J 1・L 2・M 4・N 4）：%d" % len(t))
    d1y = row(t, "D", 1, "ゆき")
    check(has(d1y, "アンケートが記録より前") and not has(d1y, "区間の名前と時刻が合わない"),
          "D 1回目ゆき：降りる前に答えた桃山台のアンケートは「記録より前」（切り替え忘れとは出さない）")
    a1y, a1k = row(t, "A", 1, "ゆき"), row(t, "A", 1, "かえり")
    a2y, a2k = row(t, "A", 2, "ゆき"), row(t, "A", 2, "かえり")
    b1y, b1k = row(t, "B", 1, "ゆき"), row(t, "B", 1, "かえり")
    b2y, b2k = row(t, "B", 2, "ゆき"), row(t, "B", 2, "かえり")
    c1y, adm = row(t, "C", 1, "ゆき"), row(t, "ADMIN", 1, "ゆき")

    # 条件（記録）と計画
    expect = [(a1y, "数字キー", "なし"), (a1k, "＋−ボタン", "なし"), (a2y, "＋−ボタン", "あり"), (a2k, "数字キー", "あり"),
              (b1y, "数字キー", "あり"), (b2y, "＋−ボタン", "なし")]
    for r, inp, unk in expect:
        if r is not None:
            check(r["入れ方"] == inp and r["わからないボタン"] == unk and r["計画の入れ方"] == inp and r["計画のボタン"] == unk,
                  "%s %s回目 %s：入れ方 %s・ボタン %s（記録＝計画）" % (r["ID"], r["回"], r["区間"], inp, unk))
    check(b1k is not None and b1k["入れ方"] == "数字キー・＋−ボタン" and has(b1k, "入れ方がまざっている"), "B 1回目かえり：入れ方がまざっている")
    check(all(r is not None and r["組"] == g for r, g in ((a1y, "奇数組"), (b1y, "偶数組"))), "組：A 奇数組・B 偶数組")
    check(a1y is not None and a1y["時間帯"] == "昼" and a2y is not None and a2y["時間帯"] == "夕方", "時間帯：10/13 昼・10/15 夕方")

    # 数
    check(a1y is not None and a1y["本番の回答数"] == "4" and a1y["練習の回答数"] == "2", "A 1回目ゆき：本番4件（再送の2行目を数えない）・練習2件")
    check(has(a1y, "練習の回答"), "A 1回目ゆき：練習の回答の印")
    check(a1k is not None and a1k["本番の回答数"] == "5" and a1k["操作ログのない回答数"] == "1", "A 1回目かえり：操作ログのない回答を時刻で入れて5件")
    check(b1y is not None and b1y["わからないで送った数"] == "1" and b1y["わからないを押した回数"] == "1" and b1y["人数なしの数"] == "1",
          "B 1回目ゆき：「わからない」で送った1件・押した1回・人数なし1件")
    check(a1y is not None and a1y["5の倍数の割合"] == "0.5", "A 1回目ゆき：5の倍数の割合 0.5（10・13・15・19 人）：%s"
          % (a1y["5の倍数の割合"] if a1y is not None else None))

    # アンケートのつなぎ方
    check(a1k is not None and a1k["アンケートの時刻"] == "2026-10-13 12:40:00" and has(a1k, "同じフォームに2回"),
          "A 1回目かえり：3行のうち、時刻が合う最初の行（12:40）を使った")
    used = sv[(sv["ID"] == "A") & (sv["フォーム"] == "1回目_南口")]["使ったか"].tolist()
    check(used.count("使った") == 1 and used.count("使わなかった（同じ区間に別の行）") == 2, "アンケートの行：A の南口3行のうち1行だけ使った")
    ans = [4, 2, 5, 1, 4, 2, 5, 1, 4, 2]
    check(a1k is not None and float(a1k["SUS点"]) == sus_by_hand(ans) == 85.0, "SUS 点：手計算 85.0 と同じ（%s）"
          % (a1k["SUS点"] if a1k is not None else None))
    check(a2k is not None and float(a2k["SUS点"]) == 50.0 and has(a2k, "答えの矛盾"), "A 2回目かえり：SUS が全部3 → 50点・答えの矛盾の印")
    check(a1k is not None and a1k["問17_選んだ入れ方"] == "数字キー", "問17：1回目に「行き」→ 数字キー")
    check(a2k is not None and a2k["問17_選んだ入れ方"] == "数字キー", "問17：2回目に「帰り」→ 数字キー")
    check(a2k is not None and a2k["問18_正しいか"] == "正しい" and a2k["問19_ボタンで読むと"] == "ボタンあり回"
          and a2k["問20_ボタンあり回の向き"] == "4", "A（奇数組）問18〜20：2回目＝ボタンあり回、問20 はそのまま 4")
    check(b2k is not None and b2k["問18_正しいか"] == "正しい" and b2k["問19_ボタンで読むと"] == "変わらない"
          and b2k["問20_ボタンあり回の向き"] == "4", "B（偶数組）問18〜20：1回目＝ボタンあり回、問20 は 6−2＝4")
    check(b1y is not None and b1y["アンケートの行"] == "4" and b1y["フォーム"] == "1回目_桃山台", "問1「ｂ 」を B として読んだ")
    check(a1k is not None and a1k["問2_混み具合"] == "3" and a1k["問5_1枚目の答えやすさ"] == "4",
          "列の順を入れ替え・古い「問5」の列・見出しの文の違いがあっても、問番号と文で正しい列を読んだ")

    # 問題の一覧
    check(has(a2y, "アンケートがない"), "A 2回目ゆき：アンケートがない")
    check(has(a2y, "2台の端末") and a2y["端末の数"] == "2", "A 2回目ゆき：2台の端末")
    check(a2k is not None and a2k["途中で閉じた数"] == "1" and a2k["本番の回答数"] == "5" and a2k["戻った数"] == "0",
          "A 2回目かえり：途中で閉じた回答1件（戻った数には入れない）")
    check(a1k is not None and a1k["区間の決め方"] == "switch" and a1k["かえりに切り替えた時刻"] == "2026-10-13 12:13:00",
          "A 1回目かえり：かえりに切り替えた時刻 12:13")
    check(c1y is not None and has(c1y, "IDが名簿にない") and has(c1y, "記録のないアンケート"), "C：名簿にない・記録のないアンケート")
    check(has(adm, "確認用アカウント") and adm["主な分析から外す理由"] == "確認用アカウント", "admin：確認用アカウントとして外す")
    check(has(b1k, "間違ったフォームの疑い") and "間違ったフォームの疑い" in (b1k["主な分析から外す理由"] or ""),
          "B 1回目かえり：かえりの前に南口のアンケート → 間違ったフォームの疑い（外す）")
    check(has(b2y, "区間の名前と時刻が合わない") and has(b2y, "間違ったフォームの疑い") and not has(b2y, "アンケートが記録より前"),
          "B 2回目ゆき：切り替え忘れ・1回目のフォーム（「記録より前」とは出さない）")
    sw = iss[(iss["ID"] == "B") & (iss["種類"] == "区間の名前と時刻が合わない")]["内容"].tolist()
    check(len(sw) == 1 and "あとの 4 件" in sw[0] and "18 分" in sw[0], "切り替え忘れ：桃山台のアンケートの前後 18 分の空き、あとの4件：%s" % sw)
    check(len(iss[(iss["種類"] == "回が記録されていない") & (iss["ID"] == "B") & (iss["日付"] == "2026-10-14")]) == 1,
          "回のない記録：ID と日付が問題の一覧に出る")
    check(has(b2k, "かえりを切り替え忘れた") and not has(b2k, "記録のないアンケート") and not b2k["主な分析から外す理由"]
          and b2k["入れ方"] == "＋−ボタン" and b2k["SUS点"] == "85.0",
          "B 2回目かえり：かえりを切り替え忘れた（外さない。SUS は使える。入れ方はかえりのバスの記録＝＋−ボタン）")
    check(b2k is not None and b2k["問17_選んだ入れ方"] == "", "B 2回目かえり：行きも帰りも＋−ボタンなので、問17 は読み替えない")
    check(has(a2k, "版が違う") and not has(a2y, "版が違う"), "A 2回目かえり：古い版の印（ゆきには付かない）")
    check(has(a2k, "端末の時計のずれ") and not has(a2k, "アンケートが記録より前"), "A 2回目：端末の時計のずれを直し、「記録より前」と出さない")
    check(has(a1k, "ほかの人のログインで送った"), "A 1回目かえり：B のログインで送った回答に印（A の区間に入れた）")
    check(has(a1k, "操作ログのない回答"), "A 1回目かえり：操作ログのない回答の印")
    check(any(iss["種類"] == "回が記録されていない"), "B 10/14：回のない記録の印")
    check(any(iss["種類"] == "再送で重なった行"), "再送で重なった行の印")
    pre = iss[iss["種類"] == "実験の前の行"]["内容"].tolist()
    check(len(pre) == 1 and "アンケート 1 行・操作ログ 1 行・回答 1 行" in pre[0], "--since：10/12 のテストの行を読まなかった：%s" % pre)
    col = iss[iss["種類"] == "列の見出し"]["内容"].tolist()
    check(any("問5 の列が 2 本" in c for c in col) and any("問2 の見出しの文" in c for c in col), "列の見出しの確かめ（問5 が2本・問2 の文）")
    check(has(b2y, "回が記録されていない") is False, "回のない記録は区間の行に混ぜない")
    for r in (a1y, b1y):
        check(r is not None and not r["主な分析から外す理由"], "%s %s回目 %s：外す理由なし" % (r["ID"], r["回"], r["区間"]) if r is not None else "-")

    # 往復の途中で ?round のない画面で入り直した（E）
    e1y, e1k = row(t, "E", 1, "ゆき"), row(t, "E", 1, "かえり")
    check(has(e1y, "回が記録されていない") and "回が記録されていない" in (e1y["主な分析から外す理由"] or "") and e1y["本番の回答数"] == "2",
          "E 1回目ゆき：途中から回のない記録 → 区間の行にも「回が記録されていない」（外す）")
    ek = " ".join(iss[(iss["ID"] == "E") & (iss["種類"] == "記録のないアンケート")]["内容"])
    check(has(e1k, "記録のないアンケート") and "回・区間のない記録が 4 行" in ek and "切り替え忘れ" not in ek,
          "E 1回目かえり：南口のアンケートには、切り替え忘れではなく回のない記録と書く：%s" % ek)
    # 前の日の ?round=1 のタブがかえりの途中で読み込み直された（F）
    f1y, f1k, f2y, f2k = row(t, "F", 1, "ゆき"), row(t, "F", 1, "かえり"), row(t, "F", 2, "ゆき"), row(t, "F", 2, "かえり")
    check(f2k is not None and f2k["本番の回答数"] == "4" and f2k["フォーム"] == "2回目_南口" and has(f2k, "回の記録が途中で替わった")
          and has(f2k, "入れ方がまざっている"), "F 2回目かえり：1回目として記録された2回答もこの往復の行に入れ、印を付けて 2回目_南口 につないだ")
    check(f1k is not None and f1k["日付"] == "2026-10-16" and not has(f1k, "同じ回を2度") and not has(f2k, "間違ったフォームの疑い"),
          "F：「1回目かえり 10/17」の行を作らない（同じ回を2度・間違ったフォームと出さない）")
    check(f2k is not None and f2k["問17_選んだ入れ方"] == "＋−ボタン" and f2k["問18_正しいか"] == "正しい"
          and f2k["問19_ボタンで読むと"] == "ボタンあり回" and f2k["問20_ボタンあり回の向き"] == "4",
          "F 2回目かえり：問17 は2回目ゆきの記録（＋−ボタン）、問18〜20 は偶数組の読み替え")
    fr = " ".join(iss[(iss["ID"] == "F") & (iss["種類"] == "回の記録が途中で替わった")]["内容"])
    check("?round=1" in fr and "2 件" in fr, "F：開いた URL（?round=1）と件数を書く：%s" % fr)
    # 桃山台で切り替えを忘れ、かえりのバスの2回答のあとで気づいた（H）
    h1y, h1k = row(t, "H", 1, "ゆき"), row(t, "H", 1, "かえり")
    check(has(h1y, "区間の名前と時刻が合わない") and not has(h1y, "アンケートが記録より前") and not h1y["主な分析から外す理由"],
          "H 1回目ゆき：あとの2件はかえりのバスの回答（切り替え忘れ）。桃山台のアンケートを「記録より前」と出さない")
    hk = " ".join(iss[(iss["ID"] == "H") & (iss["区間"] == "かえり")]["内容"])
    check(has(h1k, "区間の名前と時刻が合わない") and "最初の 2 件" in hk and "16:20" in hk,
          "H 1回目かえり：最初の2件が「ゆき」に入っている印（切り替えの時刻 16:20）：%s" % hk)
    check(h1k is not None and h1k["問17_選んだ入れ方"] == "数字キー", "H：問17「行き」→ 数字キー（ゆきは数字キーだけ）")
    # 練習の切り忘れ（J）。A の練習（南口の位置なし）には付けない
    j1y = row(t, "J", 1, "ゆき")
    jp = " ".join(iss[(iss["ID"] == "J") & (iss["種類"] == "練習の切り忘れの疑い")]["内容"])
    check(has(j1y, "練習の切り忘れの疑い") and "11:52" in jp and "11:40" not in jp and not has(a1y, "練習の切り忘れの疑い"),
          "J 1回目ゆき：南口から離れた所の練習の回答（11:52）に印：%s" % jp)
    # 桃山台駅で南口の QR を読んだ（L）
    l1y, l1k = row(t, "L", 1, "ゆき"), row(t, "L", 1, "かえり")
    ly = " ".join(iss[(iss["ID"] == "L") & (iss["種類"] == "間違ったフォームの疑い")]["内容"])
    check(has(l1y, "間違ったフォームの疑い") and has(l1y, "アンケートがない") and "16:10" in ly
          and l1k is not None and l1k["アンケートの時刻"] == "2026-10-18 16:40:00",
          "L：桃山台駅で答えた南口の行（16:10）を、ゆきの区間に知らせる。かえりは 16:40 の行：%s" % ly)
    # 1回目を ?round のない画面で答えた（M）。ボタンは両方の回で出ていたので、問18〜20 は読み替えない
    m2k = row(t, "M", 2, "かえり")
    check(m2k is not None and has(m2k, "ボタンの回が決まらない") and m2k["問18_正しいか"] == "" and m2k["問19_ボタンで読むと"] == ""
          and m2k["問20_ボタンあり回の向き"] == "", "M 2回目かえり：ボタンの回が決まらない → 問18〜20 の読み替えは空")
    mk = " ".join(iss[(iss["ID"] == "M") & (iss["種類"] == "記録のないアンケート")]["内容"])
    check("回・区間のない記録" in mk, "M 1回目：アンケートの行に、回のない記録があると書く")
    # 練習だけのまとまりは往復にしない（N）
    n1y, n2y = row(t, "N", 1, "ゆき"), row(t, "N", 2, "ゆき")
    check(n1y is not None and n1y["練習の回答数"] == "1" and n1y["本番の回答数"] == "4" and not has(n1y, "同じ回を2度")
          and not has(n1y, "本番の回答がない"), "N 1回目ゆき：朝の練習は往復の行に入れる（別の往復・同じ回を2度にしない）")
    check(n2y is not None and n2y["練習の回答数"] == "1" and n2y["入れ方"] == "＋−ボタン" and not has(n2y, "回の記録が途中で替わった"),
          "N 2回目ゆき：前の日のタブでの練習があっても、往復の回は本番の回（2回目）")
    print()
    print("出力：%s" % os.path.join(out, "result"))
    if FAIL:
        print("NG が %d 件" % len(FAIL))
        sys.exit(1)
    print("すべて OK")


if __name__ == "__main__":
    main()
