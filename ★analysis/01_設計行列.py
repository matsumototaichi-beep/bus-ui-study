# -*- coding: utf-8 -*-
"""設計行列のランクを確かめる（教授の資料 4節 の2番目）

    python -I 01_設計行列.py

データ/ui_spss.csv を読み、固定効果の候補を入れた設計行列が
ランク落ち（ある列がほかの列の足し算引き算で作れてしまう状態）していないかを見る。
ランク落ちしていれば、その変数の組み合わせは区別できない。

出すもの
    設計行列の点検.txt
    設計行列の点検.xlsx
"""
import os
import numpy as np
import pandas as pd
import patsy

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "データ", "ui_spss.csv")

d = pd.read_csv(SRC, encoding="utf-8-sig")
d = d[d["use_row"] == 1].copy()
d["rnd"] = d["rnd"].astype(int)
d["grp_odd"] = (d["grp"] == "odd").astype(int)
d["rnd2"] = (d["rnd"] == 2).astype(int)
d["cnt_c"] = d["cnt"] - d["cnt"].mean()
d["seat_mid"] = (d["seat_fb"] == "中央").astype(int)

MODELS = [
    ("M0  入力方式だけ",
     "numpad"),
    ("M1  主効果だけ（入力方式・往復・回・わからない）",
     "numpad + yuki + rnd2 + unknown_btn"),
    ("M2  M1 ＋ 回×往復の交互作用  ★教授の指摘",
     "numpad + yuki + rnd2 + unknown_btn + rnd2:yuki"),
    ("M3  M1 ＋ 組×回の交互作用",
     "numpad + yuki + rnd2 + unknown_btn + rnd2:grp_odd"),
    ("M4  教授の4節モデルの固定効果",
     "numpad + seq_all + seq_method + numpad:seq_method + switch_new + unknown_btn"),
    ("M5  M4 ＋ 往復・回",
     "numpad + seq_all + seq_method + numpad:seq_method + switch_new + unknown_btn"
     " + yuki + rnd2"),
    ("M6  M4 ＋ 実測人数・時間帯・座席",
     "numpad + seq_all + seq_method + numpad:seq_method + switch_new + unknown_btn"
     " + cnt_c + eve + seat_mid"),
    ("M7  全部入り（M5 ＋ M6 ＋ 回×往復）",
     "numpad + seq_all + seq_method + numpad:seq_method + switch_new + unknown_btn"
     " + yuki + rnd2 + rnd2:yuki + cnt_c + eve + seat_mid"),
]


def deps(X, names, tol=1e-9):
    """ランク落ちしているとき、どの列が重なっているかを返す"""
    Xc = np.asarray(X, dtype=float)
    _, s, Vt = np.linalg.svd(Xc, full_matrices=False)
    null = Vt[s < max(Xc.shape) * np.finfo(float).eps * s.max() * 1e3]
    out = []
    for v in null:
        w = v / np.abs(v).max()
        terms = [(names[i], w[i]) for i in range(len(names)) if abs(w[i]) > 1e-6]
        out.append(terms)
    return out


lines = []
rows = []


def p(s=""):
    print(s)
    lines.append(s)


p("=" * 78)
p("設計行列のランクの点検   n=%d件 / 参加者%d名" % (len(d), d["pid"].nunique()))
p("=" * 78)
p("ランク落ち＝ある列がほかの列の足し算引き算で作れてしまう状態。")
p("そのときは係数が1つに決まらず、その変数の効果を分けて測れない。")
p("")

for name, formula in MODELS:
    X = patsy.dmatrix(formula, d, return_type="dataframe")
    cols = list(X.columns)
    r = int(np.linalg.matrix_rank(np.asarray(X, dtype=float)))
    short = len(cols) - r
    ok = "OK（ランク落ちなし）" if short == 0 else "★ランク落ち %d" % short
    p("-" * 78)
    p("【%s】" % name)
    p("  式     %s" % formula)
    p("  列数 %d   ランク %d   → %s" % (len(cols), r, ok))
    rows.append({"モデル": name, "式": formula, "列数": len(cols),
                 "ランク": r, "落ち": short, "判定": ok})
    if short:
        for terms in deps(X, cols):
            s = "  重なり: " + "  ".join("%+.2f×%s" % (w, n) for n, w in terms) + "  = 0"
            p(s)
p("-" * 78)
p("")

# ---- 完全交絡の直接確認 ----
p("=" * 78)
p("【直接の確認】ある組み合わせから、別の変数が100%決まってしまうか")
p("=" * 78)


def crosscheck(left, right, lab):
    t = pd.crosstab([d[c] for c in left], d[right])
    determ = (t > 0).sum(axis=1).max() == 1
    p("")
    p("[%s]" % lab)
    p(t.to_string())
    p("  → %s" % ("★ 完全に決まる（交絡）" if determ else "決まらない（分けられる）"))
    return determ


crosscheck(["rnd", "yuki"], "numpad", "回×往復 → 入力方式")
crosscheck(["grp", "rnd"], "unknown_btn", "組×回 → わからないボタン")
crosscheck(["eve"], "numpad", "時間帯 → 入力方式")
crosscheck(["seat_fb"], "numpad", "座席の前後 → 入力方式")

out_txt = os.path.join(HERE, "設計行列の点検.txt")
with open(out_txt, "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
pd.DataFrame(rows).to_excel(os.path.join(HERE, "設計行列の点検.xlsx"), index=False)
p("")
p("書き出しました:")
p("  " + out_txt)
p("  " + os.path.join(HERE, "設計行列の点検.xlsx"))
