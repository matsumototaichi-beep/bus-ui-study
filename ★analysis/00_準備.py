# -*- coding: utf-8 -*-
"""割り付けの表と、分析用の1枚表を作る

    python -I 00_準備.py 2026-11-12

出すもの（このスクリプトと同じフォルダ）
    割り付けの表.xlsx          … 誰が・いつ・どの区間で・どの条件だったか
    データ/ui_spss.csv           … 1行1回答の分析用データ（列名は半角英数）
    データ/列の説明.xlsx         … 列の意味

※ 位置情報（緯度経度）は入れない。分析で使わないため
"""
import csv
import json
import sys
import os
import numpy as np
import pandas as pd

# ====== 2026-11-12 に確定。結果を見てから変えないこと =======================
#
# W_N = 3
#   1区間は約20件。1件ではぶれすぎ、5件では区間の1/4になり「慣れ」と区別が
#   つかなくなる。3件なら区間の約15%で「切り替え直後の数回」として自然
#
# K_RESET = False（通算で数える）
#   (1) 久しぶりにその方式へ戻っても、前の経験は消えないため
#   (2) 数え直すと K が「その区間の何件目か」になり、W（最初の3件）と
#       ぴったり重なって多重共線性が起きるため
#
W_N = 3
K_RESET = False
# ==========================================================================

HERE = os.path.dirname(os.path.abspath(__file__))
EXPORTS = os.path.join(os.path.dirname(HERE), "★exports", "★アプリデータ")
DAY = sys.argv[1] if len(sys.argv) > 1 else "2026-11-12"
SRC = os.path.join(EXPORTS, DAY)
START = pd.Timestamp("2026-10-13").date()

ODD = set("ACEGIKMOQSUW")
INPUT_BY = {(1, "yuki"): "numpad", (1, "kaeri"): "stepper",
            (2, "yuki"): "stepper", (2, "kaeri"): "numpad"}
UNK_BY = {(1, "odd"): False, (2, "odd"): True, (1, "even"): True, (2, "even"): False}


def jl(x):
    try:
        v = json.loads(x) if isinstance(x, str) else {}
        return v if isinstance(v, dict) else {}
    except Exception:
        return {}


# ---------- 読み込み ----------
S = pd.read_csv(os.path.join(SRC, "ui_sessions_rows.csv"))
R = pd.read_csv(os.path.join(SRC, "ui_responses_rows.csv"))
S["created_at"] = pd.to_datetime(S["created_at"], format="mixed", utc=True)
V = S["variant"].apply(jl)
C = S["context"].apply(jl)
SM = S["summary"].apply(jl)
EV = S["events"].apply(lambda x: json.loads(x) if isinstance(x, str) else [])

for k in ["study", "practice", "test_account", "login_id", "round", "leg",
          "input", "unknown_ui", "url_study"]:
    S[k] = V.apply(lambda v, k=k: v.get(k))
for k in ["exact_count", "congestion_class", "app_version"]:
    S[k] = C.apply(lambda v, k=k: v.get(k))
for k in ["t_first", "t_stop", "t_count", "t_total",
          "taps_value", "taps_total", "corrections"]:
    S[k] = SM.apply(lambda v, k=k: v.get(k))

S["jst"] = S["created_at"] + pd.Timedelta(hours=9)
S["login_id"] = S["login_id"].astype(str).str.upper()


def t_make(ev):
    v = [e["dt"] for e in ev if e.get("type") in ("numkey", "stepper")]
    return (v[-1] - v[0]) if v else np.nan


def t_after(ev):
    v = [e["dt"] for e in ev if e.get("type") in ("numkey", "stepper")]
    sd = [e["dt"] for e in ev if e.get("type") == "tap" and e.get("target") == "send"]
    return (sd[-1] - v[-1]) if (v and sd) else np.nan


S["t_make"] = [t_make(e) for e in EV]
S["t_after"] = [t_after(e) for e in EV]
S["t_input"] = S["t_first"] + S["t_make"]

# ---------- 回答から停留所名を拾う ----------
stop = {}
for sid, ns in zip(R["session_id"], R["next_stop"]):
    nm = jl(ns).get("name")
    if nm:
        stop[sid] = nm
S["stop_name"] = S["session_id"].map(stop)

d = S[(S["study"] == "ui") & (S["practice"] != True) & (S["test_account"] != True)
      & (S["outcome"] == "submit") & (S["jst"].dt.date >= START)].copy()
d = d.sort_values(["login_id", "jst"]).reset_index(drop=True)

# ---------- 変数を作る ----------
d["day"] = d["jst"].dt.strftime("%Y-%m-%d")
d["slot"] = np.where(d["jst"].dt.hour < 15, "day", "eve")
d["rnd"] = d["round"].astype(float).astype("Int64")
d["numpad"] = (d["input"] == "numpad").astype(int)
d["unknown_btn"] = (d["unknown_ui"] == True).astype(int)
d["grp"] = np.where(d["login_id"].isin(ODD), "odd", "even")
d["yuki"] = (d["leg"] == "yuki").astype(int)
d["eve"] = (d["slot"] == "eve").astype(int)

# 便（同じ日・同じ時間帯・同じ往復なら同じ物理の便）と、場面（便×停留所）
d["run_id"] = d["day"] + "_" + d["slot"] + "_" + d["leg"].astype(str)
d["scene_id"] = d["run_id"] + "_" + d["stop_name"].fillna("NA").astype(str)

# 乗車ブロック（同じ日・回・往復の続き）
key = d[["login_id", "day", "rnd", "leg"]].astype(str)
d["blk"] = (key.ne(key.shift()).any(axis=1)).cumsum()

# E：通算の記録番号
d["seq_all"] = d.groupby("login_id").cumcount() + 1
# K：その入力方式での通算の記録番号
if K_RESET:
    d["seq_method"] = d.groupby(["login_id", "blk"]).cumcount() + 1
else:
    d["seq_method"] = d.groupby(["login_id", "input"]).cumcount() + 1

# W：方式が切り替わった直後の最初の W_N 件
bf = d.groupby("blk").agg(lid=("login_id", "first"), inp=("input", "first"))
bf["prev"] = bf.groupby("lid")["inp"].shift()
switched = set(bf.index[(bf["prev"].notna()) & (bf["prev"] != bf["inp"])])
d["in_blk"] = d.groupby("blk").cumcount() + 1
d["switch_new"] = ((d["blk"].isin(switched)) & (d["in_blk"] <= W_N)).astype(int)

for c, lg in [("t_input", "log_t_input"), ("t_stop", "log_t_stop"),
              ("t_count", "log_t_count")]:
    d[lg] = np.log(pd.to_numeric(d[c], errors="coerce").clip(lower=1))

# 同じ条件（方式×不明可）の何回目の乗車か
d["cond"] = d["input"].astype(str) + "_" + d["unknown_btn"].astype(str)
bc = d.groupby("blk").agg(lid=("login_id", "first"), cond=("cond", "first")).reset_index()
bc["cond_nth"] = bc.groupby(["lid", "cond"]).cumcount() + 1
d = d.merge(bc[["blk", "cond_nth"]], on="blk", how="left")

# 同じ便・同じ停留所に同じ人が2回以上答えた分は、最後の1件だけを採用する
#   （2026-11-12 決定。アンケートと同じ規則にそろえた）
#   停留所名が入っていない回答は「同じ停留所」とは言えないので対象外にする
d["use_row"] = 1
named = d["stop_name"].notna()
dup = named & d.duplicated(subset=["login_id", "scene_id"], keep="last")
d.loc[dup, "use_row"] = 0
N_DUP = int(dup.sum())

# 本目（その人の何本目の乗車か）と、座席（docs/座席位置.csv から）
d["ride_no"] = d.groupby("login_id")["blk"].transform(
    lambda s: s.rank(method="dense").astype(int))
SEATCSV = os.path.join(os.path.dirname(HERE), "docs", "座席位置.csv")
seat_fb, seat_lr, seat_raw = {}, {}, {}
with open(SEATCSV, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f):
        k = ((r["ID"] or "").strip().upper(), int(r["本目"]))
        seat_fb[k] = r.get("座席区分") or ""
        seat_lr[k] = r.get("左右") or ""
        seat_raw[k] = r.get("座席（本人の言葉）") or ""
kk = list(zip(d["login_id"], d["ride_no"]))
d["seat_fb"] = [seat_fb.get(k, "") for k in kk]
d["seat_lr"] = [seat_lr.get(k, "") for k in kk]
d["seat_raw"] = [seat_raw.get(k, "") for k in kk]
N_NOSEAT = int((d["seat_fb"] == "").sum())

# ---------- 分析用の1枚表 ----------
COLS = [
    ("pid", "login_id", "参加者ID（A〜N）"),
    ("day", "day", "実験日"),
    ("slot", "slot", "時間帯 day=昼 / eve=夕方"),
    ("eve", "eve", "夕方なら1、昼なら0"),
    ("rnd", "rnd", "回（1回目／2回目）"),
    ("leg", "leg", "往復 yuki=ゆき / kaeri=かえり"),
    ("yuki", "yuki", "ゆきなら1、かえりなら0"),
    ("grp", "grp", "参加者の組 odd / even"),
    ("input", "input", "入力方式 numpad=テンキー / stepper=＋−ボタン"),
    ("numpad", "numpad", "D：テンキーなら1、＋−ボタンなら0"),
    ("unknown_btn", "unknown_btn", "「わからない」ボタンがあれば1"),
    ("seq_all", "seq_all", "E：その人の通算の記録番号"),
    ("seq_method", "seq_method", "K：その入力方式での通算の記録番号"),
    ("switch_new", "switch_new", "W：方式が切り替わった直後なら1"),
    ("cond_nth", "cond_nth", "同じ条件（方式×不明可）の何回目の乗車か"),
    ("ride_no", "ride_no", "その人の何本目の乗車か"),
    ("seat_fb", "seat_fb", "座席の前後（前方／中央／後方）"),
    ("seat_lr", "seat_lr", "座席の左右（左／右／不明）"),
    ("seat_raw", "seat_raw", "座席（本人の言葉そのまま）"),
    ("use_row", "use_row", "分析に使うなら1。同じ便の同じ停留所に2回以上答えた分は最後の1件だけ1"),
    ("run_id", "run_id", "便（日×時間帯×往復）"),
    ("scene_id", "scene_id", "場面（便×停留所）"),
    ("stop_name", "stop_name", "次に停まる停留所の名前"),
    ("cls", "congestion_class", "混み具合クラス（1〜6）"),
    ("cnt", "exact_count", "入力された車内人数"),
    ("t_first", "t_first", "人数画面が出てから最初に値に触るまで（ms）"),
    ("t_make", "t_make", "最初の値タップから最後の値タップまで（ms）"),
    ("t_input", "t_input", "t_first＋t_make（ms）主指標"),
    ("t_after", "t_after", "最後の値タップから送信まで（ms）"),
    ("t_stop", "t_stop", "停留所欄にかけた時間（ms）"),
    ("t_count", "t_count", "人数画面が出てから送信まで（ms）"),
    ("t_total", "t_total", "クラスタップから確定まで（ms）"),
    ("log_t_input", "log_t_input", "log(t_input)。1ms未満は1msに置換"),
    ("log_t_stop", "log_t_stop", "log(t_stop)"),
    ("log_t_count", "log_t_count", "log(t_count)"),
    ("taps_value", "taps_value", "値を作るためのタップ数"),
    ("taps_total", "taps_total", "総タップ数"),
    ("corrections", "corrections", "アプリが数えた取り消し（⌫を含まない。使用注意）"),
    ("app_ver", "app_version", "アプリの版"),
]
out = pd.DataFrame({new: d[old] for new, old, _ in COLS})

os.makedirs(os.path.join(HERE, "データ"), exist_ok=True)
csv_path = os.path.join(HERE, "データ", "ui_spss.csv")
out.to_csv(csv_path, index=False, encoding="utf-8-sig")

dic = pd.DataFrame([{"列名": n, "意味": lab} for n, _, lab in COLS])
dic.loc[dic["列名"] == "switch_new", "意味"] = "W：方式が切り替わった直後の%d件なら1" % W_N
dic.to_excel(os.path.join(HERE, "データ", "列の説明.xlsx"), index=False)

# ---------- 割り付けの表 ----------
rows = []
for lid, g in d.groupby("login_id"):
    for bi, (_, b) in enumerate(g.groupby("blk"), 1):
        r0 = b.iloc[0]
        ok = (INPUT_BY.get((int(r0["rnd"]), r0["leg"])) == r0["input"]
              and UNK_BY.get((int(r0["rnd"]), r0["grp"])) == bool(r0["unknown_ui"]))
        rows.append({
            "ID": lid, "本目": bi, "日": r0["day"],
            "帯": "昼" if r0["slot"] == "day" else "夕方",
            "回": int(r0["rnd"]),
            "往復": "ゆき" if r0["leg"] == "yuki" else "かえり",
            "方式": "テンキー" if r0["input"] == "numpad" else "＋−ボタン",
            "不明可": "あり" if r0["unknown_ui"] else "なし",
            "URL": r0["url_study"], "記録数": len(b),
            "採用記録数": int(b["use_row"].sum()),
            "通算番号": "%d-%d" % (b["seq_all"].min(), b["seq_all"].max()),
            "同条件の何回目": int(r0["cond_nth"]),
            "便": r0["run_id"],
            "設計どおり": "○" if ok else "★ちがう",
        })
W = pd.DataFrame(rows)

x1 = pd.crosstab(W["回"].astype(str) + "回" + W["往復"], W["方式"])
x2 = pd.crosstab([W["ID"].isin(ODD).map({True: "odd", False: "even"}), W["回"]], W["不明可"])
x3 = pd.crosstab(W["帯"], W["方式"])
x4 = pd.crosstab(W["往復"], W["方式"])
first = W.sort_values(["ID", "本目"]).groupby("ID").first()[["日", "回", "往復", "方式"]]

cond = []
for lid, g in W.groupby("ID"):
    cond.append({"ID": lid,
                 "経験した条件数": len({(r["方式"], r["不明可"]) for _, r in g.iterrows()}),
                 "乗車本数": len(g), "記録数": int(g["記録数"].sum())})
cond = pd.DataFrame(cond)

xlsx = os.path.join(HERE, "割り付けの表.xlsx")
with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
    W.to_excel(w, sheet_name="割り付け", index=False)
    x1.to_excel(w, sheet_name="点検1_方式x回往復")
    x2.to_excel(w, sheet_name="点検2_不明可x組回")
    x3.to_excel(w, sheet_name="点検3_方式x時間帯")
    x4.to_excel(w, sheet_name="点検4_方式x往復")
    first.to_excel(w, sheet_name="点検5_最初の方式")
    cond.to_excel(w, sheet_name="点検6_条件数", index=False)

print("=" * 66)
print("入力  %s" % SRC)
print("設定  W_N=%d（切り替え直後の件数）  K_RESET=%s" % (W_N, K_RESET))
print("=" * 66)
print("回答 %d件 / 参加者 %d名 / 乗車 %d本 / 便 %d / 場面 %d"
      % (len(d), d["login_id"].nunique(), len(W), d["run_id"].nunique(), d["scene_id"].nunique()))
print("設計どおりでない乗車: %d本" % (W["設計どおり"] != "○").sum())
print("重複で不採用にした回答: %d件（同じ便・同じ停留所の2回目以降。最後の1件を採用）" % N_DUP)
print("分析に使う回答: %d件" % int(d["use_row"].sum()))
print("座席が付かなかった回答: %d件" % N_NOSEAT)
print()
print("出力")
print("  %s" % xlsx)
print("  %s" % csv_path)
print("  %s" % os.path.join(HERE, "データ", "列の説明.xlsx"))
print()
print("【点検1】入力方式 × 回×往復（各行に片方しか立たなければ完全な交絡）")
print(x1.to_string())
print()
print("【点検2】「わからない」ボタン × 組×回")
print(x2.to_string())
print()
print("最初に使った方式: テンキー %d名 / ＋−ボタン %d名"
      % ((first["方式"] == "テンキー").sum(), (first["方式"] == "＋−ボタン").sum()))
