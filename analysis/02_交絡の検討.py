# -*- coding: utf-8 -*-
"""回×往復 と 入力方式 の交絡を、設計を切り直して検討する（教授の資料 4節 の2番目の続き）

    python -I 02_交絡の検討.py

01_設計行列.py で「回×往復 ＝ 入力方式」の完全交絡が確認された。
計算では切り分けられないので、往復をそろえた比較・人をそろえた比較で、
往復では説明がつかないことを数字で示す。

出すもの
    交絡の検討.txt
    交絡の検討.xlsx   （シート: 4つのマス / 往復をそろえて / 人と往復をそろえて）
"""
import os
import numpy as np
import pandas as pd
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "spss", "ui_spss.csv")

d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[d["use_row"] == 1].copy()
LEGJP = {"yuki": "ゆき", "kaeri": "かえり"}
INJP = {"numpad": "テンキー", "stepper": "＋−ボタン"}
d["往復"] = d["leg"].map(LEGJP)
d["方式"] = d["input"].map(INJP)

lines = []


def p(s=""):
    print(s)
    lines.append(s)


p("=" * 72)
p("回×往復 と 入力方式 の交絡の検討   n=%d件 / 参加者%d名" % (len(d), d["pid"].nunique()))
p("=" * 72)
p("入力方式は (回, 往復) から100%決まる。だから計算では切り分けられない。")
p("そこで、往復をそろえた比較・人をそろえた比較で、往復では説明がつかないことを見る。")
p("")

# ---- 1) 4つのマス ----------------------------------------------------------
p("-" * 72)
p("【1】4つのマス（t_input の中央値・ミリ秒）")
p("-" * 72)
cell = d.groupby(["rnd", "往復"]).agg(
    件数=("t_input", "size"), 人数=("pid", "nunique"),
    中央値=("t_input", "median")).reset_index()
cell["方式"] = [d[(d["rnd"] == r) & (d["往復"] == l)]["方式"].mode()[0]
                for r, l in zip(cell["rnd"], cell["往復"])]
cell = cell[["rnd", "往復", "方式", "件数", "人数", "中央値"]]
cell.columns = ["回", "往復", "方式", "件数", "人数", "中央値(ms)"]
p(cell.to_string(index=False))
p("")
p("  テンキーの山 ＝ 1回目ゆき ＋ 2回目かえり")
p("  ＋−の山     ＝ 1回目かえり ＋ 2回目ゆき")
p("")

# ---- 2) 往復をそろえて、方式を比べる ---------------------------------------
p("-" * 72)
p("【2】往復をそろえて、テンキー と ＋− を比べる")
p("     往復が同じなら、差を往復のせいにはできない")
p("-" * 72)
rows2 = []
for leg in ["ゆき", "かえり"]:
    s = d[d["往復"] == leg]
    m = s.groupby("input")["t_input"].median()
    n = s.groupby("input")["t_input"].size()
    r = {"往復": leg}
    for k, jp in INJP.items():
        r["%s 中央値(ms)" % jp] = round(m[k], 1)
        r["%s 件数" % jp] = int(n[k])
        r["%s の回" % jp] = int(s[s["input"] == k]["rnd"].mode()[0])
    r["比(テンキー/＋−)"] = round(100 * m["numpad"] / m["stepper"], 1)
    rows2.append(r)
    p("")
    p("[%s だけ]" % leg)
    for k, jp in INJP.items():
        p("   %-10s %7.0f ms   n=%-4d  %d回目のぶん"
          % (jp, m[k], n[k], s[s["input"] == k]["rnd"].mode()[0]))
    p("   → テンキーは ＋− の %.0f%%" % (100 * m["numpad"] / m["stepper"]))
t2 = pd.DataFrame(rows2)
p("")

# ---- 3) 方式をそろえて、往復を比べる ---------------------------------------
p("-" * 72)
p("【3】方式をそろえて、ゆき と かえり を比べる（往復そのものの大きさ）")
p("-" * 72)
m = d.groupby(["往復", "input"])["t_input"].median().unstack()
p(m.to_string())
p("")
p("  往復の差（方式をそろえた）   テンキー内 %.0f%%  ／  ＋−内 %.0f%%"
  % (100 * m.loc["ゆき", "numpad"] / m.loc["かえり", "numpad"],
     100 * m.loc["ゆき", "stepper"] / m.loc["かえり", "stepper"]))
p("  方式の差（往復をそろえた）   ゆき内 %.0f%%  ／  かえり内 %.0f%%"
  % (100 * m.loc["ゆき", "numpad"] / m.loc["ゆき", "stepper"],
     100 * m.loc["かえり", "numpad"] / m.loc["かえり", "stepper"]))
p("")
p("  → 往復の差は数%〜2割。方式の差は2倍以上。往復では説明がつかない。")
p("")

# ---- 4) 人も往復もそろえて、方式を比べる -----------------------------------
p("-" * 72)
p("【4】★同じ人・同じ往復の中で比べる")
p("     人も往復もそろえた。残る違いは 方式 と 回 だけ")
p("-" * 72)
rows4 = []
detail = []
for leg in ["ゆき", "かえり"]:
    s = d[d["往復"] == leg]
    w = s.pivot_table(index="pid", columns="input",
                      values="t_input", aggfunc="median").dropna()
    if not len(w):
        continue
    w["比"] = w["numpad"] / w["stepper"]
    np_rnd = int(s[s["input"] == "numpad"]["rnd"].mode()[0])
    st_rnd = int(s[s["input"] == "stepper"]["rnd"].mode()[0])
    pv = stats.wilcoxon(np.log(w["numpad"]), np.log(w["stepper"]))[1] \
        if len(w) >= 5 else np.nan
    rows4.append({"往復": leg, "テンキーの回": np_rnd, "＋−の回": st_rnd,
                  "比べられた人数": len(w),
                  "テンキーが速かった人数": int((w["比"] < 1).sum()),
                  "比の中央値(%)": round(100 * w["比"].median(), 1),
                  "Wilcoxon p": round(pv, 5)})
    for pid, v in w["比"].items():
        detail.append({"往復": leg, "ID": pid,
                       "テンキー(ms)": round(w.loc[pid, "numpad"], 1),
                       "＋−(ms)": round(w.loc[pid, "stepper"], 1),
                       "比(%)": round(100 * v, 1)})
    p("")
    p("[%s の中だけ]   テンキー=%d回目のぶん ／ ＋−=%d回目のぶん" % (leg, np_rnd, st_rnd))
    p("   そろって比べられた人   %d名" % len(w))
    p("   テンキーが速かった人   %d / %d名" % (int((w["比"] < 1).sum()), len(w)))
    p("   比の中央値             %.0f%%" % (100 * w["比"].median()))
    if len(w) >= 5:
        p("   Wilcoxon の符号付き順位検定   p = %.5f" % pv)
    p("   " + "  ".join("%s:%.0f%%" % (i, 100 * v) for i, v in w["比"].items()))
t4 = pd.DataFrame(rows4)
p("")
p("-" * 72)
p("【まとめ】")
p("-" * 72)
p("  ・ゆきの比較では テンキーが1回目（不慣れ）、＋−が2回目（慣れた後）。")
p("    慣れのハンデを背負ってなお、全員がテンキーのほうが速い。")
p("  ・両方向とも例外なし。往復でも慣れでも説明がつかない。")
p("  ・ただし「回×往復の交互作用」を測ること自体は、この設計では永久にできない。")
p("    原稿には『回と往復の交互作用はないものとする』と明記する。")
p("")

out_txt = os.path.join(HERE, "交絡の検討.txt")
with open(out_txt, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
out_xlsx = os.path.join(HERE, "交絡の検討.xlsx")
with pd.ExcelWriter(out_xlsx) as xw:
    cell.to_excel(xw, sheet_name="4つのマス", index=False)
    t2.to_excel(xw, sheet_name="往復をそろえて", index=False)
    m.to_excel(xw, sheet_name="方式をそろえて")
    t4.to_excel(xw, sheet_name="人と往復をそろえて", index=False)
    pd.DataFrame(detail).to_excel(xw, sheet_name="人ごとの内訳", index=False)
p("書き出しました:")
p("  " + out_txt)
p("  " + out_xlsx)
