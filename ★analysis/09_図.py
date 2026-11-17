# -*- coding: utf-8 -*-
"""卒業論文に載せる図を作る

    python -I 09_図.py

出すもの（卒業論文_作成/figures/ に PDF と PNG）
    fig5_1_参加者ごとの入力時間   14名それぞれの テンキー と ＋−
    fig5_2_同じ便の中での比較     両方式が同乗した7便
    fig5_3_慣れのカーブ           その方式で何件目か × 入力時間
    fig5_4_混雑と数えきれなさ     主観の混み具合 × アンケートの回答
    fig5_5_残差の点検             正規性・等分散・ブートストラップ

図の題目（キャプション）は LaTeX の \\caption で付けるので、画像には入れない。
執筆要領 6.3「図の題目は図の下」に従うため。
"""
import os
import glob
import itertools
import warnings

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy import stats
import statsmodels.formula.api as smf

warnings.filterwarnings("ignore")

# 明朝系が要領の指定だが、図はゴシックのほうが読みやすいので游ゴシック系を使う
plt.rcParams["font.family"] = ["Yu Gothic", "Meiryo", "MS Gothic", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["font.size"] = 11
plt.rcParams["axes.grid"] = True
plt.rcParams["grid.alpha"] = 0.25
plt.rcParams["grid.linestyle"] = "-"
plt.rcParams["axes.spines.top"] = False
plt.rcParams["axes.spines.right"] = False

C_NUM = "#1f4e79"     # テンキー（濃い青）
C_STP = "#c0504d"     # ＋−ボタン（赤茶）
C_GRAY = "#808080"

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "データ", "ui_spss.csv")
FORMS = os.path.join(os.path.dirname(HERE), "★exports", "★アンケートデータ")
OUT = os.path.abspath(os.path.join(
    HERE, "..", "..", "..", "卒業論文_作成", "figures"))
os.makedirs(OUT, exist_ok=True)


def save(fig, name):
    for ext in ("pdf", "png"):
        fig.savefig(os.path.join(OUT, "%s.%s" % (name, ext)),
                    bbox_inches="tight", dpi=200)
    plt.close(fig)
    print("  %s.pdf / .png" % name)


d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[(d["use_row"] == 1) & d["t_input"].notna()].copy()
d["方式"] = d["input"].map({"numpad": "テンキー", "stepper": "＋−ボタン"})

print("図を作ります（%s）" % OUT)

# ---- 図5.1 参加者ごとの入力時間 --------------------------------------------
w = d.pivot_table(index="pid", columns="input",
                  values="t_input", aggfunc="median").dropna()
w = w.sort_values("stepper")
fig, ax = plt.subplots(figsize=(7.0, 4.6))
for i, (pid, r) in enumerate(w.iterrows()):
    ax.plot([0, 1], [r["stepper"], r["numpad"]], "-", color=C_GRAY,
            lw=1.0, alpha=0.75, zorder=1)
    ax.annotate(pid, (1.03, r["numpad"]), va="center", fontsize=9,
                color="#333333")
ax.scatter([0] * len(w), w["stepper"], s=46, color=C_STP, zorder=3,
           label="＋−ボタン")
ax.scatter([1] * len(w), w["numpad"], s=46, color=C_NUM, zorder=3,
           label="テンキー")
ax.set_xlim(-0.35, 1.3)
ax.set_xticks([0, 1])
ax.set_xticklabels(["＋−ボタン", "テンキー"])
ax.set_ylabel("人数を入れるのにかかった時間の中央値 [ms]")
ax.set_yscale("log")
ax.set_yticks([1000, 2000, 3000, 5000, 8000])
ax.get_yaxis().set_major_formatter(matplotlib.ticker.ScalarFormatter())
ax.grid(axis="x", alpha=0)
ax.legend(loc="upper left", frameon=False)
save(fig, "fig5_1_参加者ごとの入力時間")

# ---- 図5.2 同じ便の中での比較 ----------------------------------------------
mix = d.groupby("run_id")["input"].nunique()
both = mix[mix == 2].index
e = d[d["run_id"].isin(both)]
g = e.pivot_table(index="run_id", columns="input",
                  values="t_input", aggfunc="median").sort_index()
lab = [r.replace("2026-", "").replace("_day_", " 昼 ").replace("_eve_", " 夕 ")
         .replace("yuki", "ゆき").replace("kaeri", "かえり") for r in g.index]
y = np.arange(len(g))
fig, ax = plt.subplots(figsize=(7.0, 3.8))
h = 0.36
ax.barh(y + h / 2, g["stepper"], height=h, color=C_STP, label="＋−ボタン")
ax.barh(y - h / 2, g["numpad"], height=h, color=C_NUM, label="テンキー")
for i, r in enumerate(g.index):
    ax.text(g.loc[r, "stepper"] + 80, i + h / 2,
            "%.1f 秒" % (g.loc[r, "stepper"] / 1000), va="center", fontsize=9)
    ax.text(g.loc[r, "numpad"] + 80, i - h / 2,
            "%.1f 秒" % (g.loc[r, "numpad"] / 1000), va="center", fontsize=9)
ax.set_yticks(y)
ax.set_yticklabels(lab, fontsize=10)
ax.invert_yaxis()
ax.set_xlabel("人数を入れるのにかかった時間の中央値 [ms]")
ax.set_xlim(0, g.max().max() * 1.22)
ax.grid(axis="y", alpha=0)
ax.legend(loc="lower right", frameon=False)
save(fig, "fig5_2_同じ便の中での比較")

# ---- 図5.3 慣れのカーブ ----------------------------------------------------
fig, ax = plt.subplots(figsize=(7.0, 4.2))
BIN = 5
d["区切り"] = ((d["seq_method"] - 1) // BIN) * BIN + 1
for k, lab_, c in [("numpad", "テンキー", C_NUM), ("stepper", "＋−ボタン", C_STP)]:
    s = d[d["input"] == k]
    gg = s.groupby("区切り")["t_input"].agg(["median", "size"])
    gg = gg[gg["size"] >= 8]
    ax.plot(gg.index + BIN / 2, gg["median"], "-o", color=c, lw=2,
            ms=4.5, label=lab_)
ax.set_xlabel("その方式で何件目か")
ax.set_ylabel("人数を入れるのにかかった時間の中央値 [ms]")
ax.set_ylim(0, None)
ax.legend(frameon=False)
save(fig, "fig5_3_慣れのカーブ")

# ---- 図5.4 混雑と数えきれなさ ----------------------------------------------
ODD = set("ACEGIKMOQSUW")
rows = []
for f in sorted(glob.glob(os.path.join(FORMS, "UI研究*.csv"))):
    b = os.path.basename(f)
    rnd = 1 if "1回目" in b else 2
    leg = "ゆき" if "桃山台" in b else "かえり"
    t = pd.read_csv(f, encoding="utf-8-sig")
    idc = [c for c in t.columns if "ID" in c]
    qs = {n: [c for c in t.columns if c.startswith("問%d" % n)] for n in (2, 3, 4)}
    if not idc or not all(qs.values()):
        continue
    for _, r in t.iterrows():
        pid = str(r[idc[0]]).strip().upper()[:1]
        if pid in "ABCDEFGHIJKLMN":
            rows.append({"混み具合": r[qs[2][0]], "数えきれない": r[qs[3][0]],
                         "適当に送った": r[qs[4][0]]})
a = pd.DataFrame(rows).dropna()
fig, ax = plt.subplots(figsize=(6.4, 4.2))
for col, c, m, lab_ in [("数えきれない", "#2e6f9e", "o", "数えきれないことがあった"),
                        ("適当に送った", "#b35c00", "s", "だいたいの人数を入れて送った")]:
    gg = a.groupby("混み具合")[col].agg(["mean", "size"])
    ax.plot(gg.index, gg["mean"], "-", color=c, lw=2, marker=m, ms=6, label=lab_)
n = a.groupby("混み具合").size()
for x, v in n.items():
    ax.annotate("n=%d" % v, (x, 0.72), ha="center", fontsize=9, color="#666666")
ax.set_xlabel("そのバスの混み具合（1 = 空いていた 〜 5 = 混んでいた）")
ax.set_ylabel("回答の平均\n（1 = なかった 〜 5 = 何度もあった）")
ax.set_xticks([1, 2, 3, 4, 5])
ax.set_ylim(0.5, 5.2)
ax.legend(frameon=False, loc="upper left")
save(fig, "fig5_4_混雑と数えきれなさ")

# ---- 図5.5 残差の点検 ------------------------------------------------------
SEED, B = 20261113, 4000
d["E"] = d["seq_all"] - d["seq_all"].mean()
d["K"] = d["seq_method"] - d["seq_method"].mean()
d["y"] = np.log(d["t_input"].clip(lower=1))
d["c"] = (d["cnt"] - d["cnt"].mean()).fillna(0)
m = smf.mixedlm("y ~ numpad + E + K + numpad:K + switch_new + unknown_btn + c",
                d, groups=d["pid"], re_formula="~numpad").fit(method="lbfgs")
d["res"] = m.resid
d["fit"] = d["y"] - d["res"]
rng = np.random.default_rng(SEED)
pids = d["pid"].unique()
boot = []
for _ in range(B):
    sub = pd.concat([d[d["pid"] == q] for q in rng.choice(pids, len(pids), True)])
    mm = sub.groupby("input")["t_input"].median()
    if {"numpad", "stepper"} <= set(mm.index):
        boot.append(100 * mm["numpad"] / mm["stepper"])
boot = np.array(boot)

fig, ax = plt.subplots(2, 2, figsize=(9.2, 6.6))
ax[0, 0].hist(d["res"], bins=50, color=C_NUM, alpha=0.85)
ax[0, 0].set_xlabel("残差（対数）")
ax[0, 0].set_ylabel("件数")
ax[0, 0].set_title("(a) 残差の分布", fontsize=11, loc="left")
stats.probplot(d["res"], dist="norm", plot=ax[0, 1])
ax[0, 1].get_lines()[0].set(color=C_NUM, ms=2.5, alpha=0.5)
ax[0, 1].get_lines()[1].set(color=C_STP, lw=1.5)
ax[0, 1].set_title("(b) 正規Q–Qプロット", fontsize=11, loc="left")
ax[0, 1].set_xlabel("理論分位点")
ax[0, 1].set_ylabel("標本分位点")
ax[1, 0].scatter(d["fit"], d["res"], s=5, alpha=0.3, color=C_NUM)
ax[1, 0].axhline(0, color="k", lw=0.8)
ax[1, 0].set_xlabel("予測値（対数）")
ax[1, 0].set_ylabel("残差")
ax[1, 0].set_title("(c) 予測値と残差", fontsize=11, loc="left")
ax[1, 1].hist(boot, bins=50, color="#4f8a8b", alpha=0.9)
lo, hi = np.percentile(boot, [2.5, 97.5])
ax[1, 1].axvline(lo, color=C_STP, lw=1.2, ls="--")
ax[1, 1].axvline(hi, color=C_STP, lw=1.2, ls="--")
ax[1, 1].set_xlabel("テンキーの時間は＋−ボタンの何%か")
ax[1, 1].set_ylabel("回数")
ax[1, 1].set_title("(d) 参加者のブートストラップ（%d回）" % len(boot),
                   fontsize=11, loc="left")
fig.tight_layout()
save(fig, "fig5_5_残差の点検")

print("\n図に使った数字")
print("  5.1  14名中 テンキーが速い %d名 / 比の中央値 %.0f%%"
      % (int((w["numpad"] < w["stepper"]).sum()),
         100 * (w["numpad"] / w["stepper"]).median()))
print("  5.2  便%d本すべてでテンキーが速い: %s"
      % (len(g), bool((g["numpad"] < g["stepper"]).all())))
print("  5.4  混み具合×数えきれない ρ=%+.2f"
      % stats.spearmanr(a["混み具合"], a["数えきれない"]).statistic)
print("  5.5  ブートストラップ 95%%区間 [%.0f%%, %.0f%%]" % (lo, hi))
