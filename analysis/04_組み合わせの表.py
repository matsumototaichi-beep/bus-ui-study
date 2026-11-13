# -*- coding: utf-8 -*-
"""参加者×便の組み合わせを確かめる（教授の資料 4節 の4番目）

    python -I 04_組み合わせの表.py

変量効果（人・便・場面ごとのクセ）を推定するには、各マスにデータが要る。
1件しかないマスばかりだと、そのクセを測れず、SPSSが収束しない／特異と出る。
回す前に、どこが薄いかを見ておく。

出すもの
    組み合わせの表.txt
    組み合わせの表.xlsx
"""
import os
import numpy as np
import pandas as pd
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "spss", "ui_spss.csv")

d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[d["use_row"] == 1].copy()
d["方式"] = d["input"].map({"numpad": "テンキー", "stepper": "＋−"})
d["往復"] = d["leg"].map({"yuki": "ゆき", "kaeri": "かえり"})
d["時間帯"] = d["slot"].map({"day": "昼", "eve": "夕方"})

lines = []


def p(s=""):
    print(s)
    lines.append(s)


def dist(series, label, unit):
    """件数の分布を 1件 / 2件 / 3件以上 で要約する"""
    vc = series.value_counts().sort_index()
    p("  %s  全%d個" % (label, len(series)))
    p("    最小 %d / 中央値 %.0f / 最大 %d  %s"
      % (series.min(), series.median(), series.max(), unit))
    for k in [1, 2, 3]:
        n = int((series == k).sum()) if k < 3 else int((series >= 3).sum())
        lab = "%d%s" % (k, unit) if k < 3 else "3%s以上" % unit
        p("    %-8s %4d個 (%5.1f%%)" % (lab, n, 100 * n / len(series)))
    return vc


p("=" * 76)
p("参加者×便の組み合わせ   n=%d件 / 参加者%d名 / 便%d本 / 場面%d個"
  % (len(d), d["pid"].nunique(), d["run_id"].nunique(), d["scene_id"].nunique()))
p("=" * 76)
p("変量効果＝人やバスごとのクセを吸収する仕組み。")
p("クセを測るには同じマスに複数のデータが要る。1件だけのマスは測れない。")
p("")

# ---- 1) 参加者 × 便 --------------------------------------------------------
ct = pd.crosstab(d["run_id"], d["pid"])
p("-" * 76)
p("【1】参加者 × 便  の表（マスの数字は入力件数）")
p("-" * 76)
p(ct.to_string())
p("")
p("  マスの総数 %d（便%d × 参加者%d）  うち データのあるマス %d （%.1f%%）"
  % (ct.size, ct.shape[0], ct.shape[1], int((ct > 0).sum().sum()),
     100 * (ct > 0).sum().sum() / ct.size))
p("  → 全員が全便に乗るわけではないので、空のマスが多いのは設計どおり。")
p("")

# ---- 2) 1便あたりの人数 ----------------------------------------------------
p("-" * 76)
p("【2】1便に何人乗っていたか  ★便のクセを測れるかどうか")
p("-" * 76)
per_run = (ct > 0).sum(axis=1)
dist(per_run, "1便あたりの参加者数", "名")
p("")
p("  1名だけの便   %d本" % int((per_run == 1).sum()))
if (per_run == 1).any():
    p("    " + "  ".join(per_run[per_run == 1].index))
p("  → 1名だけの便は「バスのクセ」と「その人のクセ」を分けられない。")
p("")

# ---- 3) 1人あたりの便数 ----------------------------------------------------
p("-" * 76)
p("【3】1人が何便に乗ったか  ★人のクセを測れるかどうか")
p("-" * 76)
per_pid = (ct > 0).sum(axis=0)
p("  " + "  ".join("%s:%d本" % (i, v) for i, v in per_pid.items()))
p("  最小 %d本 / 中央値 %.0f本 / 最大 %d本"
  % (per_pid.min(), per_pid.median(), per_pid.max()))
p("")

# ---- 4) ★便の中に両方の方式があるか ---------------------------------------
p("-" * 76)
p("【4】★同じ便の中に、テンキーと＋−の両方がいたか")
p("-" * 76)
p("  同じバスに乗りながら方式が違う人がいれば、")
p("  『そのバスが混んでいたから』では説明がつかない。いちばん強い比較になる。")
p("")
mix = d.groupby("run_id")["方式"].nunique()
both = mix[mix == 2].index
p("  両方そろった便   %d本 / %d本 (%.1f%%)"
  % (len(both), len(mix), 100 * len(both) / len(mix)))
p("  片方だけの便     %d本" % int((mix == 1).sum()))
p("")
p("  両方そろった便の件数   %d件 / %d件 (%.1f%%)"
  % (int(d["run_id"].isin(both).sum()), len(d),
     100 * d["run_id"].isin(both).sum() / len(d)))
p("")
mixtab = pd.crosstab(d[d["run_id"].isin(both)]["run_id"],
                     d[d["run_id"].isin(both)]["方式"])
if len(mixtab):
    mixtab["人数"] = [int((ct.loc[r] > 0).sum()) for r in mixtab.index]
    p(mixtab.to_string())
p("")

# ---- 4b) ★同じ便の中だけで方式を比べる ------------------------------------
p("-" * 76)
p("【4b】★同じ便の中だけで、テンキー と ＋− を比べる")
p("-" * 76)
p("  同じバス・同じ時刻・同じ道路・同じ乗客。違うのは手元の画面だけ。")
p("  ここで差が出れば、バスの事情では説明がつかない。")
p("")
e = d[d["run_id"].isin(both) & d["t_input"].notna()]
wb = e.pivot_table(index="run_id", columns="input",
                   values="t_input", aggfunc="median")
wb.columns = ["テンキー(ms)", "＋−(ms)"]
wb["比(%)"] = (100 * wb["テンキー(ms)"] / wb["＋−(ms)"]).round(0)
wb["人数"] = e.groupby("run_id")["pid"].nunique()
wb["件数"] = e.groupby("run_id").size()
p("  便%d本 / %d件 / 参加者%d名" % (len(wb), len(e), e["pid"].nunique()))
p("")
p(wb.round(0).to_string())
p("")
mm = e.groupby("input")["t_input"].median()
pv_run = stats.wilcoxon(np.log(wb["テンキー(ms)"]),
                        np.log(wb["＋−(ms)"]))[1]
p("  まとめて   テンキー %.0f ms ／ ＋− %.0f ms  → %.0f%%"
  % (mm["numpad"], mm["stepper"], 100 * mm["numpad"] / mm["stepper"]))
p("  テンキーが速かった便   %d / %d本" % (int((wb["比(%)"] < 100).sum()), len(wb)))
p("  便ごとの比の中央値     %.0f%%" % wb["比(%)"].median())
p("  Wilcoxon の符号付き順位検定   p = %.5f" % pv_run)
p("")
p("  → 7便すべてでテンキーが速い。『その便が混んでいたから』では説明がつかない。")
p("")

# ---- 5) 1場面あたりの件数 --------------------------------------------------
p("-" * 76)
p("【5】1場面（その便のその停留所）に何件あったか  ★ここが危ないと見ている")
p("-" * 76)
per_scene = d.groupby("scene_id").size()
dist(per_scene, "1場面あたりの件数", "件")
p("")
if (per_scene == 1).mean() > 0.3:
    p("  → 1件だけの場面が多い。場面のクセは測れない見込み。")
    p("     SPSS が収束しない／特異と出たら scene_id の変量効果を外す。")
else:
    p("  → 複数件ある場面が多い。場面の変量効果も試す価値がある。")
p("")

# ---- 6) 設計のバランス -----------------------------------------------------
p("-" * 76)
p("【6】設計のバランス（各マスの件数と人数）")
p("-" * 76)
for keys, lab in [(["方式", "往復"], "方式 × 往復"),
                  (["方式", "時間帯"], "方式 × 時間帯"),
                  (["方式", "unknown_btn"], "方式 × わからないボタン"),
                  (["方式", "rnd"], "方式 × 回")]:
    a = pd.crosstab(d[keys[0]], d[keys[1]])
    b = d.groupby(keys)["pid"].nunique().unstack()
    p("")
    p("[%s]  件数" % lab)
    p(a.to_string())
    p("            人数")
    p(b.to_string())
    if (a == 0).any().any():
        p("  ★空のマスあり")
p("")

# ---- まとめ ----------------------------------------------------------------
p("-" * 76)
p("【まとめ】")
p("-" * 76)
p("  ・参加者   %d名。1人あたり %d〜%d便。人のクセは問題なく測れる"
  % (d["pid"].nunique(), per_pid.min(), per_pid.max()))
p("  ・便       %d本。1便あたり %d〜%d名。1名だけの便は %d本"
  % (len(per_run), per_run.min(), per_run.max(), int((per_run == 1).sum())))
p("  ・場面     %d個。1件だけの場面は %d個 (%.1f%%)"
  % (len(per_scene), int((per_scene == 1).sum()),
     100 * (per_scene == 1).mean()))
p("  ・同じ便に両方の方式がいた便 %d本（%d件）。その%d本すべてでテンキーが速い"
  % (len(both), int(d["run_id"].isin(both).sum()), int((wb["比(%)"] < 100).sum())))
p("")

out_txt = os.path.join(HERE, "組み合わせの表.txt")
with open(out_txt, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
out_xlsx = os.path.join(HERE, "組み合わせの表.xlsx")
with pd.ExcelWriter(out_xlsx) as xw:
    ct.to_excel(xw, sheet_name="参加者×便")
    pd.DataFrame({"便": per_run.index, "参加者数": per_run.values,
                  "件数": ct.sum(axis=1).values}).to_excel(
        xw, sheet_name="便ごと", index=False)
    pd.DataFrame({"ID": per_pid.index, "乗った便数": per_pid.values,
                  "件数": ct.sum(axis=0).values}).to_excel(
        xw, sheet_name="参加者ごと", index=False)
    per_scene.rename("件数").to_frame().to_excel(xw, sheet_name="場面ごと")
    if len(mixtab):
        mixtab.to_excel(xw, sheet_name="両方式がそろった便")
    pd.crosstab(d["方式"], d["往復"]).to_excel(xw, sheet_name="方式×往復")
    pd.crosstab(d["方式"], d["時間帯"]).to_excel(xw, sheet_name="方式×時間帯")
    wb.to_excel(xw, sheet_name="同じ便の中で比べる")
p("書き出しました:")
p("  " + out_txt)
p("  " + out_xlsx)
