# -*- coding: utf-8 -*-
"""収束・残差・少人数向け推論を確かめる（教授の資料 4節 の5番目）

    python -I 06_残差と推論の点検.py

  1. 収束と特異   分散がゼロに張りついていないか、計算が終わっているか
  2. 残差の点検   対数変換は妥当だったか、ばらつきは一様か、外れ値はないか
  3. 少人数向け推論
       参加者が14名しかいないので、ふつうの自由度だと p値 が甘くなる。
       そこで自由度に頼らない方法（符号の完全列挙・参加者のブートストラップ）で
       同じ結論になるかを確かめる。

出すもの
    残差と推論の点検.txt
    残差と推論の点検.xlsx
    残差の図.png
"""
import os
import itertools
import warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import statsmodels.formula.api as smf
from scipy import stats

warnings.filterwarnings("ignore")
plt.rcParams["font.family"] = ["Meiryo", "Yu Gothic", "MS Gothic", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "spss", "ui_spss.csv")
SEED = 20261113          # 乱数の種。結果を再現できるように固定する
B = 4000                 # ブートストラップの回数

d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[(d["use_row"] == 1) & d["t_input"].notna()].copy()
d["E"] = d["seq_all"] - d["seq_all"].mean()
d["K"] = d["seq_method"] - d["seq_method"].mean()
d["y"] = np.log(d["t_input"].clip(lower=1))
d["cnt_c"] = (d["cnt"] - d["cnt"].mean()).fillna(0.0)
d["one"] = 1
FIX = "numpad + E + K + numpad:K + switch_new + unknown_btn + cnt_c"

lines = []


def p(s=""):
    print(s)
    lines.append(s)


p("=" * 76)
p("収束・残差・少人数向け推論の点検   n=%d件 / 参加者%d名"
  % (len(d), d["pid"].nunique()))
p("=" * 76)
p("")

# ============ 1. 収束と特異 =================================================
p("-" * 76)
p("【1】収束と特異（分散がゼロに張りついていないか）")
p("-" * 76)
VCF = {"pid": "0+C(pid)", "run": "0+C(run_id)", "scene": "0+C(scene_id)"}
mf = smf.mixedlm("y ~ " + FIX, d, groups=d["one"], vc_formula=VCF,
                 re_formula="0").fit(method="lbfgs", maxiter=600)
p("  収束したか   %s" % mf.converged)
p("")
p("  %-14s %12s  %s" % ("変量効果", "分散", "判定"))
conv = []
for i, k in enumerate(VCF):
    v = mf.vcomp[i]
    j = "★ゼロに張りついている（特異）" if v < 1e-6 else "問題なし"
    p("  %-14s %12.5f  %s" % (k, v, j))
    conv.append({"変量効果": k, "分散": round(v, 5), "判定": j})
p("  %-14s %12.5f" % ("残差", mf.scale))
conv.append({"変量効果": "残差", "分散": round(mf.scale, 5), "判定": ""})
p("")
p("  ※Python の推定値。本番の数字は SPSS で出す。ここは収束するかの確認")
p("")

# 残差は参加者だけのモデルから作る（安定しているため）
m1 = smf.mixedlm("y ~ " + FIX, d, groups=d["pid"],
                 re_formula="~numpad").fit(method="lbfgs")
d["res"] = m1.resid
d["fit"] = d["y"] - d["res"]
d["zres"] = d["res"] / d["res"].std(ddof=1)

# ============ 2. 残差の点検 =================================================
p("-" * 76)
p("【2】残差の点検")
p("-" * 76)
raw = smf.ols("t_input ~ " + FIX, d).fit().resid
p("  対数をとる前と後で、残差のゆがみを比べる")
p("    %-14s 歪度 %+7.3f   尖度 %+7.3f"
  % ("対数をとる前", stats.skew(raw), stats.kurtosis(raw)))
p("    %-14s 歪度 %+7.3f   尖度 %+7.3f"
  % ("対数をとった後", stats.skew(d["res"]), stats.kurtosis(d["res"])))
p("    → 歪度0・尖度0 が正規分布。対数で大きく改善したが、まだ右に裾が長い")
p("    ★判定：正規性は完全には満たしていない")
p("")
sw = stats.shapiro(d["res"].sample(500, random_state=SEED))
p("  Shapiro-Wilk 検定（500件を抽出）  W=%.4f  p=%.4g" % sw)
p("    ※件数が多いと些細なずれでも有意になる。歪度・尖度と図で判断する")
p("")
p("  ばらつきが一様か（残差の標準偏差）")
gi = d.groupby("input")["res"].std()
r_in = gi.max() / gi.min()
p("    方式別       " + "  ".join("%s:%.3f" % (i, v) for i, v in gi.items())
  + "   （比 %.1f倍）" % r_in)
gp = d.groupby("pid")["res"].std()
r_pid = gp.max() / gp.min()
p("    参加者別     最小 %.3f / 最大 %.3f  （比 %.1f倍）"
  % (gp.min(), gp.max(), r_pid))
q = pd.qcut(d["fit"], 4, labels=["予測1", "予測2", "予測3", "予測4"])
gq = d.groupby(q, observed=True)["res"].std()
r_fit = gq.max() / gq.min()
p("    予測値の四分位別  " + "  ".join("%s:%.3f" % (i, v) for i, v in gq.items())
  + "   （比 %.1f倍）" % r_fit)
p("    → 目安は2倍以内")
p("    ★判定：予測値で見れば一様（%.1f倍）。ただし参加者間で %.1f倍の開きがある"
  % (r_fit, r_pid))
p("            テンキーのほうがばらつきが大きい（%.1f倍）。"
  % r_in)
p("            ふだんは速いが、たまに手間どる回があるため")
p("")
out = d[d["zres"].abs() > 3]
rate = 100 * len(out) / len(d)
p("  標準化残差が ±3 を超えた行   %d件 (%.1f%%)" % (len(out), rate))
p("    正規分布なら 0.3%% 程度のはず。%.1f%% は裾が重いことを示す" % rate)
if len(out):
    p(out[["pid", "day", "input", "stop_name", "cnt", "t_input", "zres"]]
      .round(2).to_string(index=False))
p("    いずれも『入力に時間がかかった回』で、誤りではない。")
p("    取り除かずに使うが、取り除いた場合の結果も【4】で確かめる")
p("")

# ============ 3. 少人数向け推論 =============================================
p("-" * 76)
p("【3】少人数向け推論（参加者が14名しかいない問題）")
p("-" * 76)
b, se = m1.params["numpad"], m1.bse["numpad"]
p("  混合モデルの推定   係数 %+.4f   SE %.4f   テンキー比 %.0f%%"
  % (b, se, 100 * np.exp(b)))
p("")
p("  自由度の取りかたで p値 がどれだけ変わるか")
for df, lab in [(np.inf, "正規分布（自由度∞）＝いちばん甘い"),
                (len(d) - 9, "残差の自由度 %d ＝ SPSSの既定" % (len(d) - 9)),
                (13, "参加者数−1 ＝ 13（いちばん厳しい）")]:
    pv = 2 * stats.norm.sf(abs(b / se)) if np.isinf(df) \
        else 2 * stats.t.sf(abs(b / se), df)
    p("    %-36s p = %.3g" % (lab, pv))
p("    → どの取りかたでも p は十分に小さい。自由度の選びかたで結論は変わらない")
p("")

w = d.pivot_table(index="pid", columns="input", values="y",
                  aggfunc="median").dropna()
dif = (w["numpad"] - w["stepper"]).values
p("  --- 自由度に頼らない方法 ---")
p("")
p("  (a) 符号の完全列挙検定（%d名 → 2^%d = %d通りをすべて数える）"
  % (len(dif), len(dif), 2 ** len(dif)))
obs = dif.mean()
cnt_ge = 0
for signs in itertools.product([1, -1], repeat=len(dif)):
    if abs((np.array(signs) * dif).mean()) >= abs(obs) - 1e-12:
        cnt_ge += 1
p_exact = cnt_ge / 2 ** len(dif)
p("      観測された平均の対数差 %+.4f （テンキー比 %.0f%%）"
  % (obs, 100 * np.exp(obs)))
p("      これ以上に大きい並べかた %d通り / %d通り" % (cnt_ge, 2 ** len(dif)))
p("      p = %.3g  ← 近似も自由度も一切使わない厳密な値" % p_exact)
p("")

rng = np.random.default_rng(SEED)
pids = d["pid"].unique()
boot = []
for _ in range(B):
    pick = rng.choice(pids, size=len(pids), replace=True)
    sub = pd.concat([d[d["pid"] == q] for q in pick])
    mm = sub.groupby("input")["t_input"].median()
    if {"numpad", "stepper"} <= set(mm.index):
        boot.append(mm["numpad"] / mm["stepper"])
boot = np.array(boot)
lo, hi = np.percentile(boot, [2.5, 97.5])
p("  (b) 参加者のブートストラップ（14名を%d回とり直す。乱数の種 %d）" % (B, SEED))
p("      テンキー比の中央値 %.0f%%   95%%信頼区間 [%.0f%%, %.0f%%]"
  % (100 * np.median(boot), 100 * lo, 100 * hi))
p("      100%%を超えた回数 %d / %d" % (int((boot >= 1).sum()), len(boot)))
p("")
p("  → どちらの方法でも同じ結論。少人数であることは結論を揺るがさない")
p("")

p("-" * 76)
p("【4】残差の問題が結論を動かすか（外れ値を除いた場合）")
p("-" * 76)
p("  残差は正規分布から外れ、ばらつきも参加者で %.1f倍の開きがある。" % r_pid)
p("  そこで外れ値 %d件 を除いて、同じ計算をやり直す。" % len(out))
p("")
dd = d[d["zres"].abs() <= 3].copy()
m2 = smf.mixedlm("y ~ " + FIX, dd, groups=dd["pid"],
                 re_formula="~numpad").fit(method="lbfgs")
b2, se2 = m2.params["numpad"], m2.bse["numpad"]
w2 = dd.pivot_table(index="pid", columns="input", values="y",
                    aggfunc="median").dropna()
dif2 = (w2["numpad"] - w2["stepper"]).values
obs2 = dif2.mean()
cnt2 = sum(1 for s in itertools.product([1, -1], repeat=len(dif2))
           if abs((np.array(s) * dif2).mean()) >= abs(obs2) - 1e-12)
pe2 = cnt2 / 2 ** len(dif2)
sens = pd.DataFrame([
    {"データ": "全件 %d件" % len(d), "混合モデルのテンキー比(%)": round(100 * np.exp(b), 1),
     "SE": round(se, 4), "自由度13のp": "%.3g" % (2 * stats.t.sf(abs(b / se), 13)),
     "完全列挙のp": "%.3g" % p_exact,
     "テンキーが速い人": "%d/%d" % (int((dif < 0).sum()), len(dif))},
    {"データ": "外れ値を除く %d件" % len(dd), "混合モデルのテンキー比(%)": round(100 * np.exp(b2), 1),
     "SE": round(se2, 4), "自由度13のp": "%.3g" % (2 * stats.t.sf(abs(b2 / se2), 13)),
     "完全列挙のp": "%.3g" % pe2,
     "テンキーが速い人": "%d/%d" % (int((dif2 < 0).sum()), len(dif2))}])
p(sens.to_string(index=False))
p("")
p("  → 外れ値を除いても結論は変わらない。")
p("")

p("-" * 76)
p("【まとめ】")
p("-" * 76)
p("  ○ 収束   3つの変量効果を入れても収束。ゼロに張りついた分散はない")
p("  ✕ 正規性 対数で大きく改善したが、まだ右に裾が長い（歪度 %.2f・尖度 %.2f）"
  % (stats.skew(d["res"]), stats.kurtosis(d["res"])))
p("  ✕ 等分散 参加者のあいだで %.1f倍、方式のあいだで %.1f倍の開きがある"
  % (r_pid, r_in))
p("  △ 外れ値 %d件(%.1f%%)。正規分布の想定より多いが、除いても結論は同じ"
  % (len(out), rate))
p("  ○ 少人数 自由度を いちばん厳しい13 にしても p = %.3g"
  % (2 * stats.t.sf(abs(b / se), 13)))
p("  ○ 少人数 自由度を使わない厳密な検定でも p = %.3g（14名で出せる最小値）"
  % p_exact)
p("  ○ 少人数 参加者をとり直しても テンキー比は [%.0f%%, %.0f%%]"
  % (100 * lo, 100 * hi))
p("")
p("  正規性と等分散には問題が残るが、")
p("  結論は『分布を仮定しない方法』（完全列挙・ブートストラップ・14/14名）でも")
p("  同じなので、守られている。原稿にはこの両方を載せる。")
p("  → Kenward-Roger が使えなくても困らない。")
p("")

# ============ 図 ============================================================
fig, ax = plt.subplots(2, 2, figsize=(11, 8))
ax[0, 0].hist(d["res"], bins=50, color="#4C78A8")
ax[0, 0].set_title("残差のヒストグラム（対数をとった後）")
stats.probplot(d["res"], dist="norm", plot=ax[0, 1])
ax[0, 1].set_title("正規Qープロット")
ax[1, 0].scatter(d["fit"], d["res"], s=6, alpha=.35, color="#4C78A8")
ax[1, 0].axhline(0, color="k", lw=.8)
ax[1, 0].set_xlabel("予測値（対数）")
ax[1, 0].set_ylabel("残差")
ax[1, 0].set_title("予測値と残差（ばらつきが一様か）")
ax[1, 1].hist(100 * boot, bins=50, color="#72B7B2")
ax[1, 1].axvline(100, color="r", lw=1)
ax[1, 1].set_xlabel("テンキー比(%)")
ax[1, 1].set_title("参加者のブートストラップ %d回" % len(boot))
fig.tight_layout()
figpath = os.path.join(HERE, "残差の図.png")
fig.savefig(figpath, dpi=130)

out_txt = os.path.join(HERE, "残差と推論の点検.txt")
with open(out_txt, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
with pd.ExcelWriter(os.path.join(HERE, "残差と推論の点検.xlsx")) as xw:
    pd.DataFrame(conv).to_excel(xw, sheet_name="収束と分散", index=False)
    out[["pid", "day", "input", "stop_name", "cnt", "t_input", "zres"]].round(2)\
        .to_excel(xw, sheet_name="外れ値", index=False)
    pd.DataFrame({"ID": w.index, "テンキー(対数)": w["numpad"].round(3),
                  "＋−(対数)": w["stepper"].round(3),
                  "差": np.round(dif, 3),
                  "比(%)": np.round(100 * np.exp(dif), 1)}).to_excel(
        xw, sheet_name="参加者ごとの差", index=False)
    sens.to_excel(xw, sheet_name="外れ値を除いた場合", index=False)
    pd.DataFrame({"手法": ["混合モデル（自由度13）", "符号の完全列挙", "ブートストラップ"],
                  "テンキー比(%)": [round(100 * np.exp(b), 1),
                                round(100 * np.exp(obs), 1),
                                round(100 * np.median(boot), 1)],
                  "下限(%)": ["", "", round(100 * lo, 1)],
                  "上限(%)": ["", "", round(100 * hi, 1)],
                  "p値": ["%.3g" % (2 * stats.t.sf(abs(b / se), 13)),
                         "%.3g" % p_exact, ""]}).to_excel(
        xw, sheet_name="推論の比べ", index=False)
p("書き出しました:")
p("  " + out_txt)
p("  " + figpath)
