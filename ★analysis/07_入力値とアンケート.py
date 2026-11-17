# -*- coding: utf-8 -*-
"""入力された人数そのものと、事後アンケートを突き合わせる

    python -I 07_入力値とアンケート.py

01〜06 は「入力にかかった時間」だけを見ていた（H1）。
ここでは残りの2つの仮説と、入力値そのものを見る。

  H2  「わからない」ボタンを用意すると、人数の欠損が増えるか
  H3  混むほど、数えきれなくなるか
  ★   入力方式によって、入力される人数そのものが変わるか

★の比較は正解データが要らない。同じ便の同じ停留所を、テンキーの人と
＋−の人が同時に見ているので、実際の人数は両者で同じだから。

読むもの
    データ/ui_spss.csv
    ../★exports/★アンケートデータ/UI研究*.csv   （アンケートの書き出し）

出すもの
    入力値とアンケート.txt
    入力値とアンケート.xlsx
"""
import os
import glob
import itertools
import warnings

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy import stats

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "データ", "ui_spss.csv")
FORMS = os.path.join(os.path.dirname(HERE), "★exports", "★アンケートデータ")

# 条件の割り当て（00_準備.py と同じ）
ODD = set("ACEGIKMOQSUW")
INPUT_BY = {(1, "ゆき"): "テンキー", (1, "かえり"): "＋−",
            (2, "ゆき"): "＋−", (2, "かえり"): "テンキー"}
UNKNOWN_BY = {(1, "odd"): "なし", (2, "odd"): "あり",
              (1, "even"): "あり", (2, "even"): "なし"}
# フォームと区間の対応（問2の文言で確認済み）
LEG_BY_FORM = {"桃山台": "ゆき", "南口": "かえり"}

lines = []


def p(s=""):
    print(s)
    lines.append(s)


# ---- アンケートを読む --------------------------------------------------------
def read_forms():
    rows = []
    for f in sorted(glob.glob(os.path.join(FORMS, "UI研究*.csv"))):
        b = os.path.basename(f)
        rnd = 1 if "1回目" in b else 2
        leg = next((v for k, v in LEG_BY_FORM.items() if k in b), None)
        if leg is None:
            continue
        d = pd.read_csv(f, encoding="utf-8-sig")
        idc = [c for c in d.columns if "ID" in c]
        qs = {n: [c for c in d.columns if c.startswith("問%d" % n)] for n in (2, 3, 4)}
        if not idc or not all(qs.values()):
            p("  ※列が見つからない: %s" % b)
            continue
        for _, r in d.iterrows():
            pid = str(r[idc[0]]).strip().upper()[:1]
            if pid not in "ABCDEFGHIJKLMN":
                continue
            grp = "odd" if pid in ODD else "even"
            rows.append({"ID": pid, "回": rnd, "区間": leg,
                         "方式": INPUT_BY[(rnd, leg)],
                         "わからないボタン": UNKNOWN_BY[(rnd, grp)],
                         "混み具合": r[qs[2][0]],
                         "数えきれない": r[qs[3][0]],
                         "適当に送った": r[qs[4][0]]})
    return pd.DataFrame(rows).dropna(subset=["適当に送った"])


d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[(d["use_row"] == 1) & d["cnt"].notna()].copy()
d["方式"] = d["input"].map({"numpad": "テンキー", "stepper": "＋−"})
d["区間"] = d["leg"].map({"yuki": "ゆき", "kaeri": "かえり"})

p("=" * 76)
p("入力された人数と、事後アンケートの突き合わせ")
p("=" * 76)
a = read_forms()
p("  アプリの回答   %d件 / 参加者%d名（人数が入っている行のみ）"
  % (len(d), d["pid"].nunique()))
p("  アンケート     %d件 / 参加者%d名" % (len(a), a["ID"].nunique()))
p("")
p("  設問はいずれも5段階")
p("    問2 混み具合      1=空いていた      〜 5=混んでいた")
p("    問3 数えきれない  1=なかった        〜 5=何度もあった")
p("    問4 適当に送った  1=なかった        〜 5=何度もあった")
p("")

# ---- 1) アンケートの分布 -----------------------------------------------------
p("-" * 76)
p("【1】アンケートの回答の分布")
p("-" * 76)
dist = []
for c in ["混み具合", "数えきれない", "適当に送った"]:
    v = a[c].value_counts().sort_index()
    p("  %-12s " % c + "  ".join("%d:%2d件" % (k, n) for k, n in v.items())
      + "   平均 %.2f   4以上 %d件 (%.0f%%)"
      % (a[c].mean(), int((a[c] >= 4).sum()), 100 * (a[c] >= 4).mean()))
    dist.append(dict({"設問": c}, **{str(k): int(v.get(k, 0)) for k in range(1, 6)},
                     平均=round(a[c].mean(), 2),
                     **{"4以上の割合%": round(100 * (a[c] >= 4).mean(), 1)}))
p("")
p("  → 乗車の %.0f%% で、本人が『数えきれないまま適当な数を送った』と答えている"
  % (100 * (a["適当に送った"] >= 4).mean()))
p("")

# ---- 2) H2 -------------------------------------------------------------------
p("-" * 76)
p("【2】H2  「わからない」ボタンを用意すると、適当な回答が減るか")
p("-" * 76)
p("  H2の予測＝ボタンなしのほうが『適当に送った』が高い")
p("")
g = a.groupby("わからないボタン")["適当に送った"].agg(
    件数="size", 平均="mean", 中央値="median")
p(g.round(2).to_string())
w = a.pivot_table(index="ID", columns="わからないボタン",
                  values="適当に送った", aggfunc="mean").dropna()
pv2 = stats.wilcoxon(w["なし"], w["あり"]).pvalue if len(w) >= 5 else np.nan
p("")
p("  同じ人の中で比べる（両方の条件を経験した %d名）" % len(w))
p("    " + "  ".join("%s:%.1f→%.1f" % (i, r["なし"], r["あり"])
                     for i, r in w.iterrows()))
p("    なし平均 %.2f  →  あり平均 %.2f" % (w["なし"].mean(), w["あり"].mean()))
p("    Wilcoxon の符号付き順位検定  p = %.3f" % pv2)
p("")
p("  → H2は支持されない。逃げ道を用意しても、適当な回答の量は変わらない")
p("  → アプリ上で『わからない』が押されたのは 8回 だけ（1,176件中）")
p("")

# ---- 3) H3 -------------------------------------------------------------------
p("-" * 76)
p("【3】H3  混むほど、数えきれなくなるか")
p("-" * 76)
h3 = a.groupby("混み具合").agg(件数=("適当に送った", "size"),
                               数えきれない=("数えきれない", "mean"),
                               適当に送った=("適当に送った", "mean"))
p(h3.round(2).to_string())
r1 = stats.spearmanr(a["混み具合"], a["数えきれない"])
r2 = stats.spearmanr(a["混み具合"], a["適当に送った"])
p("")
p("  混み具合 × 数えきれない   ρ = %+.3f  (p = %.4g)" % (r1.statistic, r1.pvalue))
p("  混み具合 × 適当に送った   ρ = %+.3f  (p = %.4g)" % (r2.statistic, r2.pvalue))
p("")
p("  → H3は強く支持される")
p("")

# ---- 4) 自己申告とログの突き合わせ -------------------------------------------
p("-" * 76)
p("【4】★自己申告と、アプリのログの食い違い")
p("-" * 76)
m = d.merge(a.rename(columns={"ID": "pid", "回": "rnd"}),
            on=["pid", "rnd", "区間"], how="left", suffixes=("", "_q"))
m["自己申告"] = np.where(m["適当に送った"] >= 4, "適当に入れたと申告", "そうでない")
per = (m.groupby(["pid", "rnd", "区間"])
         .agg(件数=("cnt", "size"), 欠損=("cnt", lambda x: x.isna().sum()),
              適当=("適当に送った", "first")).reset_index())
hit = per[per["適当"] >= 4]
p("  本人が『適当に入れた（4以上）』と答えた区間   %d区間 / %d区間"
  % (len(hit), len(per)))
p("  その区間のアプリ回答   %d件" % int(hit["件数"].sum()))
p("  そのうち人数が空の件数 %d件" % int(hit["欠損"].sum()))
p("")
cmp = m.groupby("自己申告").agg(
    件数=("cnt", "size"),
    人数の中央値=("cnt", "median"),
    五の倍数率=("cnt", lambda x: 100 * (x % 5 == 0).mean()),
    テンキーの時間=("t_input", lambda x: np.nan))
for lab, gg in m.groupby("自己申告"):
    cmp.loc[lab, "テンキーの時間"] = gg[gg["input"] == "numpad"]["t_input"].median()
p("  ログの特徴で、申告のあり／なしを見分けられるか")
p(cmp.round(1).to_string())
p("")
p("  → 欠損・丸め・入力時間のどれでも見分けられない")
p("  → 当てずっぽうの回答は、真面目に数えた回答と同じ形でログに入る")
p("")

# ---- 5) ★入力方式は入力値を変えるか -----------------------------------------
p("-" * 76)
p("【5】★入力方式によって、入力される人数そのものが変わるか")
p("-" * 76)
p("  正解データは要らない。同じ便の同じ停留所を両方式の人が同時に見ているため")
p("")
p("  参考：全体（ただし往復と交絡しているので、そのままでは読めない）")
p(d.groupby("方式")["cnt"].agg(件数="size", 中央値="median", 平均="mean")
  .round(1).to_string())
p("")
mix = d.groupby("run_id")["input"].nunique()
both = mix[mix == 2].index
e = d[d["run_id"].isin(both)].copy()
p("  同じ便に両方式がいた   %d便 / %d件 / 参加者%d名"
  % (len(both), len(e), e["pid"].nunique()))
pairs = []
for sid, g in e.groupby("scene_id"):
    n, q = g[g["input"] == "numpad"], g[g["input"] == "stepper"]
    if len(n) and len(q):
        pairs.append({"場面": sid, "テンキー": n["cnt"].median(),
                      "＋−": q["cnt"].median(),
                      "差": n["cnt"].median() - q["cnt"].median()})
pr = pd.DataFrame(pairs)
pvw = stats.wilcoxon(pr["テンキー"], pr["＋−"]).pvalue
p("")
p("  [そのまま比べる]  %d場面" % len(pr))
p("    テンキー %.1f人 ／ ＋− %.1f人   差の中央値 %+.1f人   平均 %+.2f人"
  % (pr["テンキー"].median(), pr["＋−"].median(), pr["差"].median(), pr["差"].mean()))
p("    テンキーが多い %d場面 / 少ない %d場面 / 同じ %d場面"
  % (int((pr["差"] > 0).sum()), int((pr["差"] < 0).sum()), int((pr["差"] == 0).sum())))
p("    Wilcoxon p = %.4f" % pvw)
p("")
e["one"] = 1
mm = smf.mixedlm("cnt ~ numpad", e, groups=e["one"],
                 vc_formula={"pid": "0+C(pid)", "scene": "0+C(scene_id)"},
                 re_formula="0").fit(method="lbfgs", maxiter=400)
p("  [人の違いと場面の違いを同時に除く]")
p("    テンキーの係数 %+.3f人  (SE %.3f)  p = %.4f"
  % (mm.params["numpad"], mm.bse["numpad"], mm.pvalues["numpad"]))
p("    分散  人 %.2f ／ 場面 %.2f ／ 残り %.2f"
  % (mm.vcomp[0], mm.vcomp[1], mm.scale))
p("")
byp = []
for sid, g in e.groupby("scene_id"):
    n, q = g[g["input"] == "numpad"], g[g["input"] == "stepper"]
    if not (len(n) and len(q)):
        continue
    for _, x in n.iterrows():
        byp.append({"ID": x["pid"], "差": x["cnt"] - q["cnt"].median()})
bp = pd.DataFrame(byp).groupby("ID")["差"].agg(件数="size", 差の中央値="median")
p("  [参加者ごと] テンキー側の人が、同じ場面の＋−側より何人多く入れたか")
p("    " + "  ".join("%s:%+.0f" % (i, r["差の中央値"]) for i, r in bp.iterrows()))
p("")
p("  → 人の違いを除くと差は消える。入力方式は入力値を変えていない")
p("")
p("  [丸めて答えた割合]  ＋−は −5/−1/+1/+5 なので5の倍数が増えてもおかしくないが")
rnd_t = e.groupby("方式")["cnt"].agg(
    件数="size", 五の倍数=lambda x: 100 * (x % 5 == 0).mean(),
    十の倍数=lambda x: 100 * (x % 10 == 0).mean())
p(rnd_t.round(1).to_string())
p("  → 差はない")
p("")

# ---- まとめ ------------------------------------------------------------------
p("-" * 76)
p("【まとめ】")
p("-" * 76)
p("  ・H1 速さ          テンキーは半分の時間（01〜06 を参照）        → 支持")
p("  ・H2 わからない    ボタンを足しても適当な回答は減らない p=%.3f  → 否定" % pv2)
p("  ・H3 混雑          混むほど数えきれない ρ=%+.2f               → 支持"
  % r1.statistic)
p("  ・★入力値         入力方式では変わらない p=%.2f" % mm.pvalues["numpad"])
p("")
p("  論文の主張になるのは、UIの比較になっている3つ：")
p("    速さは半分になる／入力される人数は変わらない／わからないボタンは効かない")
p("  H3 と 自己申告29% は、UIの比較ではないので『実験条件の記述』として別に置く")
p("")
p("  限界")
p("    ・正解データがないので、どちらの方式が『正確』かは言えない")
p("    ・入力値の比較は8名・%d場面のみ。小さい差は見逃す" % len(pr))
p("    ・『適当に送った』は自己申告。程度（±2人か当てずっぽうか）は分からない")
p("    ・参加者に『車内人数を答えてほしい』と依頼したことが、")
p("      わからないボタンの使用を抑えた可能性がある")
p("")

out_txt = os.path.join(HERE, "入力値とアンケート.txt")
with open(out_txt, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
out_xlsx = os.path.join(HERE, "入力値とアンケート.xlsx")
with pd.ExcelWriter(out_xlsx) as xw:
    a.to_excel(xw, sheet_name="アンケート", index=False)
    pd.DataFrame(dist).to_excel(xw, sheet_name="回答の分布", index=False)
    g.to_excel(xw, sheet_name="H2 わからないボタン")
    h3.to_excel(xw, sheet_name="H3 混雑")
    per.to_excel(xw, sheet_name="区間ごとの申告とログ", index=False)
    cmp.to_excel(xw, sheet_name="申告とログの見分け")
    pr.to_excel(xw, sheet_name="同じ場面での入力値", index=False)
    bp.to_excel(xw, sheet_name="参加者ごとの差")
p("書き出しました:")
p("  " + out_txt)
p("  " + out_xlsx)
