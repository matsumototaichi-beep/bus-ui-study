# -*- coding: utf-8 -*-
"""固定効果どうしの相関と多重共線性を確かめる（教授の資料 4節 の3番目）

    python -I 03_変数の相関.py

多重共線性＝説明変数どうしが似すぎていて、どちらの効果か分けにくい状態。
完全に重なる（＝ランク落ち）ほどではないが、係数の誤差が大きくなり、
符号が反転したり p値 が当てにならなくなる。

見るもの
  ・相関係数（ピアソン／スピアマン）   |r| が 0.8 を超えたら要注意
  ・VIF（分散拡大要因）               10 超で要注意、5 超で注意
  ・条件数 κ                          30 超で注意、100 超で要注意

出すもの
    変数の相関.txt
    変数の相関.xlsx
"""
# ====== 2026-11-13 に確定。結果を見てから変えないこと =======================
# 本番のモデルは「教授案そのまま」。E（seq_all）と K（seq_method）を両方入れる。
#   E と K は r=0.93、VIF 11〜12 と似すぎている。
#   ただし入力方式の VIF は 1.19 で無傷。E と K の入れかたを4通り変えても
#   テンキー比は 50〜51% で動かなかった（【7】の感度分析）。
#   → E と K は「慣れの調整項」としてセットで扱い、個々の係数は解釈しない。
#     原稿には VIF の表と、この4通りの比較を載せる。
#
# Cさん 2026-11-05 かえり の1件（4タップしたあと消して空のまま送信、5187ms）は
#   残す。実際に入力操作をした時間であり、消したこと自体が使われ方の一部のため。
#   use_row=1 のまま。処理の変更なし。
# ==========================================================================

import os
import numpy as np
import pandas as pd
import patsy
import warnings
import statsmodels.formula.api as smf

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "データ", "ui_spss.csv")

d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[d["use_row"] == 1].copy()
d["rnd2"] = (d["rnd"] == 2).astype(int)
d["grp_odd"] = (d["grp"] == "odd").astype(int)
d["seat_mid"] = (d["seat_fb"] == "中央").astype(int)
N_ALL = len(d)
N_NA_CNT = int(d["cnt"].isna().sum())
N_NA_T = int(d["t_input"].isna().sum())
d["cnt_c"] = d["cnt"] - d["cnt"].mean()

NAMES = {
    "numpad":      "入力方式（テンキー=1）",
    "seq_all":     "E 全体で何件目（アプリへの慣れ）",
    "seq_method":  "K その方式で何件目（方式ごとの慣れ）",
    "switch_new":  "W 切り替え直後の3件",
    "unknown_btn": "わからないボタンあり",
    "yuki":        "往復（ゆき=1）",
    "rnd2":        "回（2回目=1）",
    "eve":         "時間帯（夕方=1）",
    "cnt":         "実測人数",
    "seat_mid":    "座席が中央",
}
COLS = list(NAMES)

lines = []


def p(s=""):
    print(s)
    lines.append(s)


p("=" * 76)
p("固定効果どうしの相関と多重共線性   n=%d件 / 参加者%d名"
  % (len(d), d["pid"].nunique()))
p("=" * 76)
p("多重共線性＝説明変数どうしが似すぎていて、どちらの効果か分けにくい状態。")
p("完全に重なればランク落ち（01_設計行列.py）。ここはその一歩手前を見る。")
p("")
p("欠損  実測人数 cnt が %d件、人数入力時間 t_input が %d件（全%d件中）"
  % (N_NA_CNT, N_NA_T, N_ALL))
p("      相関は組ごとに欠損を外して計算。VIF は欠損行を落として計算。")
p("")
p("変数の中身")
for c in COLS:
    v = d[c]
    if set(v.dropna().unique()) <= {0, 1}:
        p("  %-12s %-32s  1の割合 %5.1f%%  (%d件)"
          % (c, NAMES[c], 100 * v.mean(), int(v.sum())))
    else:
        p("  %-12s %-32s  中央値 %.0f  範囲 %.0f〜%.0f"
          % (c, NAMES[c], v.median(), v.min(), v.max()))
p("")

# ---- 1) 相関行列 -----------------------------------------------------------
pear = d[COLS].corr(method="pearson").round(3)
spea = d[COLS].corr(method="spearman").round(3)

p("-" * 76)
p("【1】ピアソンの相関係数（直線的な似かた）")
p("-" * 76)
p(pear.to_string())
p("")
p("-" * 76)
p("【2】スピアマンの相関係数（順位での似かた）")
p("-" * 76)
p(spea.to_string())
p("")

p("-" * 76)
p("【3】強く相関している組み合わせ（|r| の大きい順。0.3 以上を表示）")
p("-" * 76)
pairs = []
for i, a in enumerate(COLS):
    for b in COLS[i + 1:]:
        r = pear.loc[a, b]
        s = spea.loc[a, b]
        if abs(r) >= 0.3 or abs(s) >= 0.3:
            if max(abs(r), abs(s)) >= 0.8:
                judge = "★要注意（0.8超）"
            elif max(abs(r), abs(s)) >= 0.6:
                judge = "注意（0.6超）"
            else:
                judge = "問題なし"
            pairs.append({"変数1": a, "変数1の意味": NAMES[a],
                          "変数2": b, "変数2の意味": NAMES[b],
                          "ピアソン": r, "スピアマン": s, "判定": judge})
pairs = sorted(pairs, key=lambda x: -abs(x["ピアソン"]))
for x in pairs:
    p("  %-12s × %-12s   r=%+.3f  ρ=%+.3f   %s"
      % (x["変数1"], x["変数2"], x["ピアソン"], x["スピアマン"], x["判定"]))
if not pairs:
    p("  （0.3 以上の組み合わせなし）")
p("")

# ---- 2) VIF ----------------------------------------------------------------
def vif_table(formula, label):
    X = patsy.dmatrix(formula, d, return_type="dataframe",
                      NA_action="drop")
    if not len(X):
        raise RuntimeError("設計行列が空: " + formula)
    cols = [c for c in X.columns if c != "Intercept"]
    out = []
    for c in cols:
        y = X[c].values
        Z = X[[k for k in X.columns if k != c]].values
        beta, *_ = np.linalg.lstsq(Z, y, rcond=None)
        resid = y - Z @ beta
        ss_tot = ((y - y.mean()) ** 2).sum()
        r2 = 1 - (resid ** 2).sum() / ss_tot if ss_tot > 0 else 0.0
        v = np.inf if r2 >= 1 - 1e-12 else 1 / (1 - r2)
        out.append({"モデル": label, "変数": c, "R2": round(r2, 4),
                    "VIF": round(v, 2) if np.isfinite(v) else np.inf,
                    "判定": "★要注意(10超)" if v > 10 else
                            ("注意(5超)" if v > 5 else "問題なし")})
    Xn = X.values
    Xs = Xn / np.where(np.linalg.norm(Xn, axis=0) == 0, 1,
                       np.linalg.norm(Xn, axis=0))
    s = np.linalg.svd(Xs, compute_uv=False)
    kappa = s.max() / s.min() if s.min() > 0 else np.inf
    return pd.DataFrame(out), kappa, len(X)


F4 = ("numpad + seq_all + seq_method + numpad:seq_method"
      " + switch_new + unknown_btn")
F4C = ("numpad + center(seq_all) + center(seq_method)"
       " + numpad:center(seq_method) + switch_new + unknown_btn")
F6 = F4 + " + cnt_c + eve + seat_mid"

p("-" * 76)
p("【4】VIF（分散拡大要因）  その変数がほかの変数でどれだけ言い当てられるか")
p("-" * 76)
p("  VIF=1 なら独立。VIF=10 は「係数の誤差が約3.2倍に膨らむ」という意味。")
p("")
vifs = []
for formula, label in [(F4, "M4 教授の4節モデル（そのまま）"),
                       (F4C, "M4 中心化した版"),
                       (F6, "M6 ＋人数・時間帯・座席")]:
    t, kappa, n_used = vif_table(formula, label)
    t.insert(1, "使った件数", n_used)
    vifs.append(t)
    p("[%s]   n=%d件   条件数 κ = %.1f  → %s"
      % (label, n_used, kappa,
         "★要注意(100超)" if kappa > 100 else
         ("注意(30超)" if kappa > 30 else "問題なし")))
    p(t[["変数", "R2", "VIF", "判定"]].to_string(index=False))
    p("")

# ---- 3) E と K をくわしく --------------------------------------------------
p("-" * 76)
p("【5】E（全体で何件目）と K（その方式で何件目）の関係をくわしく")
p("-" * 76)
p("  この2つは作り方が似ている。K を通算で数える（K_RESET=False）と決めたので、")
p("  どれくらい重なっているかを確かめておく。")
p("")
p("  全体          r = %+.3f" % pear.loc["seq_all", "seq_method"])
for k, jp in [("numpad", "テンキーの行だけ"), ("stepper", "＋−の行だけ")]:
    s = d[d["input"] == k]
    p("  %-12s r = %+.3f  (n=%d)"
      % (jp, s["seq_all"].corr(s["seq_method"]), len(s)))
p("")
byp = d.groupby("pid").apply(
    lambda s: s["seq_all"].corr(s["seq_method"]), include_groups=False)
p("  参加者ごとの r   中央値 %+.3f   最小 %+.3f   最大 %+.3f"
  % (byp.median(), byp.min(), byp.max()))
p("  " + "  ".join("%s:%+.2f" % (i, v) for i, v in byp.items()))
p("")

# ---- 4) W のぐあい ---------------------------------------------------------
p("-" * 76)
p("【6】W（切り替え直後の3件）のぐあい")
p("-" * 76)
p("  W=1 の件数   %d件 / %d件  (%.1f%%)"
  % (int(d["switch_new"].sum()), len(d), 100 * d["switch_new"].mean()))
p("")
p(pd.crosstab(d["switch_new"], d["方式" if "方式" in d else "input"],
              margins=True).to_string())
p("")
p("  W=1 のときの K の中央値 %.0f ／ W=0 のときの K の中央値 %.0f"
  % (d[d["switch_new"] == 1]["seq_method"].median(),
     d[d["switch_new"] == 0]["seq_method"].median()))
p("")

# ---- 5) 相関が結論を動かすかどうか -----------------------------------------
p("-" * 76)
p("【7】★E と K の相関は、肝心の「入力方式」の答えを動かすか")
p("-" * 76)
p("  相関が強くても、困るのは E と K 自身の係数だけかもしれない。")
p("  そこで E と K の入れかたを4通り変えて、入力方式の係数が動くかを見る。")
p("")
warnings.filterwarnings("ignore")
e = d[d["t_input"].notna()].copy()
e["E"] = e["seq_all"] - e["seq_all"].mean()
e["K"] = e["seq_method"] - e["seq_method"].mean()
_b = np.polyfit(e["E"], e["K"], 1)
e["Kr"] = e["K"] - (_b[0] * e["E"] + _b[1])
p("  n=%d件（t_input の欠損5件を除く）" % len(e))
p("  E と K の相関 %.3f  →  直交化した Kr との相関 %.3f"
  % (e["E"].corr(e["K"]), e["E"].corr(e["Kr"])))
p("")
SENS = [
    ("1 E と K を両方入れる（教授案そのまま）",
     "numpad + E + K + numpad:K + switch_new + unknown_btn"),
    ("2 E を外す（K だけ）",
     "numpad + K + numpad:K + switch_new + unknown_btn"),
    ("3 K を外す（E だけ。交互作用も外れる）",
     "numpad + E + switch_new + unknown_btn"),
    ("4 K を直交化する（E ＋ 残差 Kr）",
     "numpad + E + Kr + numpad:Kr + switch_new + unknown_btn"),
]
sens = []
p("  %-38s %10s %8s %9s %10s"
  % ("E と K の入れかた", "方式の係数", "SE", "テンキー比", "交互作用p"))
p("  " + "-" * 78)
for lab, rhs in SENS:
    m = smf.mixedlm("np.log(t_input.clip(lower=1)) ~ " + rhs, e,
                    groups=e["pid"], re_formula="~numpad").fit(method="lbfgs")
    bq, se = m.params["numpad"], m.bse["numpad"]
    ix = [k for k in m.params.index if ":" in k and "numpad" in k]
    pv = m.pvalues[ix[0]] if ix else np.nan
    p("  %-38s %+10.4f %8.4f %8.0f%% %10.3f"
      % (lab, bq, se, 100 * np.exp(bq), pv))
    sens.append({"E と K の入れかた": lab, "式": rhs,
                 "方式の係数": round(bq, 4), "SE": round(se, 4),
                 "テンキー比(%)": round(100 * np.exp(bq), 1),
                 "交互作用 p": round(pv, 3) if pv == pv else ""})
p("")
p("  → 入れかたを変えても テンキー比は 50〜51%。結論は動かない。")
p("  → 慣れ方が方式で違うか（方式×K）は p が大きく、差は見あたらない。")
p("")
p("-" * 76)
p("【まとめ】")
p("-" * 76)
p("  ・E と K は r=0.93、VIF 11〜12。似すぎている。")
p("    この2つ『それぞれの』係数は分けて読めない。別々に解釈しないこと。")
p("  ・ただし入力方式の VIF は 1.2〜3.9 で問題なし。")
p("    入れかたを変えても テンキー比 50〜51% で動かない。主結果は無事。")
p("  ・W（切り替え直後）は VIF 1.2 でほぼ独立。そのまま入れてよい。")
p("  ・回 と 時間帯 は r=0.48。2回目は夕方が多いが、問題になる水準ではない。")
p("  ・原稿には VIF の表と、この入れかた4通りの比較（感度分析）を載せる。")
p("")

out_txt = os.path.join(HERE, "変数の相関.txt")
with open(out_txt, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
out_xlsx = os.path.join(HERE, "変数の相関.xlsx")
with pd.ExcelWriter(out_xlsx) as xw:
    pd.DataFrame([{"変数": c, "意味": NAMES[c]} for c in COLS]).to_excel(
        xw, sheet_name="変数の説明", index=False)
    pear.to_excel(xw, sheet_name="ピアソン相関")
    spea.to_excel(xw, sheet_name="スピアマン相関")
    pd.DataFrame(pairs).to_excel(xw, sheet_name="強い組み合わせ", index=False)
    pd.concat(vifs).to_excel(xw, sheet_name="VIF", index=False)
    pd.DataFrame(sens).to_excel(xw, sheet_name="入れかた4通り", index=False)
p("書き出しました:")
p("  " + out_txt)
p("  " + out_xlsx)
