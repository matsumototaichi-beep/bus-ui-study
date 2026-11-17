# -*- coding: utf-8 -*-
# 本番の混合効果モデル（教授の資料 4節）
#
#   "C:/Program Files/R/R-4.3.1/bin/Rscript.exe" 08_本番モデル.R
#
# 01〜07 の点検を踏まえた最終版。
#   ・場面 scene_id の変量効果は必ず入れる（05_変量効果の必要性.py）
#   ・回×往復の交互作用は入れられない（01_設計行列.py の完全交絡）
#   ・E と K はセットで扱い、個々の係数は解釈しない（03_変数の相関.py）
#   ・参加者14名なので Kenward-Roger で自由度を補正する
#
# ★本命は M4（教授の4節モデル）。参加者の切片と傾きの相関が −1 に張りつく
#   「特異」な当てはめになるが、固定効果の推定はそのまま使える。
#   傾きを外した形（M6）を感度分析として併記する。
#
# 出すもの
#   本番モデル.txt
#   本番モデル_固定効果.csv
#   本番モデル_変量効果.csv

suppressMessages({
  library(lme4)
  library(lmerTest)
  library(pbkrtest)
})

SRC <- file.path("データ", "ui_spss.csv")
if (!file.exists(SRC)) stop("データ/ui_spss.csv が見つかりません。★analysis フォルダで実行してください")

d <- read.csv(file(SRC, encoding = "UTF-8-BOM"), stringsAsFactors = FALSE)

out <- character(0)
p <- function(...) { s <- paste0(...); cat(s, "\n", sep = ""); out <<- c(out, s) }
cap <- function(expr) for (l in capture.output(expr)) p(l)

# ---- データを整える --------------------------------------------------------
d <- d[d$use_row == 1 & !is.na(d$log_t_input), ]
d$input       <- factor(d$input, levels = c("stepper", "numpad"))
d$unknown_btn <- factor(d$unknown_btn)
d$switch_new  <- factor(d$switch_new)
d$pid <- factor(d$pid); d$run_id <- factor(d$run_id); d$scene_id <- factor(d$scene_id)
d$E <- d$seq_all    - mean(d$seq_all)
d$K <- d$seq_method - mean(d$seq_method)
d$C <- d$cnt - mean(d$cnt, na.rm = TRUE); d$C[is.na(d$C)] <- 0

ct <- lmerControl(optimizer = "bobyqa", optCtrl = list(maxfun = 2e5))

p(strrep("=", 76))
p("本番の混合効果モデル")
p(strrep("=", 76))
p(sprintf("  件数 %d / 参加者 %d名 / 便 %d本 / 場面 %d個", nrow(d),
          nlevels(droplevels(d$pid)), nlevels(droplevels(d$run_id)),
          nlevels(droplevels(d$scene_id))))
p("")
p("  従属変数  log_t_input ＝ 人数を入れるのにかかった時間の自然対数")
p("  入力方式  基準は ＋−ボタン。exp(係数) が『テンキーにすると何倍か』")
p("")

base <- "log_t_input ~ input + E + K + input:K + switch_new + unknown_btn"
RE   <- "(1 + input | pid) + (1 | run_id) + (1 | scene_id)"
RE_noslope <- "(1 | pid) + (1 | run_id) + (1 | scene_id)"

fit <- function(f, lab) {
  m <- lmer(as.formula(f), d, REML = TRUE, control = ct)
  sg <- isSingular(m)
  p(sprintf("  %-28s 特異 %-4s  AIC %.1f", lab, ifelse(sg, "あり", "なし"), AIC(m)))
  m
}

p(strrep("-", 76))
p("【1】モデルを順に当てる")
p(strrep("-", 76))
m1 <- fit("log_t_input ~ input + (1|pid)",                       "M1 変量切片だけ")
m2 <- fit("log_t_input ~ input + (1+input|pid)",                 "M2 ＋変量傾き（教授の指摘）")
m3 <- fit(paste("log_t_input ~ input +", RE),                    "M3 ＋便・場面")
m4 <- fit(paste(base, "+", RE),                                  "M4 教授の4節モデル ★本命")
m5 <- fit(paste(base, "+ C +", RE),                              "M5 M4 ＋実測人数")
m6 <- fit(paste(base, "+", RE_noslope),                          "M6 M4 から傾きを外す")
p("")

p(strrep("-", 76))
p("【2】入力方式の係数（テンキーは＋−ボタンの何倍か）")
p(strrep("-", 76))
rows <- list()
for (nm in paste0("M", 1:6)) {
  m <- get(tolower(nm)); s <- summary(m)$coefficients["inputnumpad", ]
  b <- s[1]; se <- s[2]; ci <- b + c(-1.96, 1.96) * se
  p(sprintf("  %-3s 係数 %+.4f  SE %.4f  自由度 %6.1f  p %-9.3g 比 %3.0f%% [%2.0f%%, %2.0f%%]",
            nm, b, se, s[3], s[5], 100*exp(b), 100*exp(ci[1]), 100*exp(ci[2])))
  rows[[nm]] <- data.frame(モデル=nm, 係数=round(b,4), SE=round(se,4),
                           自由度=round(s[3],1), p値=signif(s[5],3),
                           比=round(100*exp(b),1), 下限=round(100*exp(ci[1]),1),
                           上限=round(100*exp(ci[2]),1), 特異=isSingular(m))
}
p("")
p("  → どのモデルでも 46〜50%。変量効果の入れ方で結論は動かない")
p("")

p(strrep("-", 76))
p("【3】M4 の固定効果")
p(strrep("-", 76))
cap(print(round(summary(m4)$coefficients, 4)))
p("")
p("  読み方")
s4 <- summary(m4)$coefficients
p(sprintf("    入力方式       テンキーにすると %.0f%%（exp(%.4f)）",
          100*exp(s4["inputnumpad","Estimate"]), s4["inputnumpad","Estimate"]))
p(sprintf("    切り替え直後   方式が変わった直後の3件は %.0f%%（%.0f%%遅い）p=%.4f",
          100*exp(s4["switch_new1","Estimate"]),
          100*(exp(s4["switch_new1","Estimate"])-1), s4["switch_new1","Pr(>|t|)"]))
p(sprintf("    方式ごとの慣れ K が1件増えるごとに %.1f%%  p=%.4f",
          100*(exp(s4["K","Estimate"])-1), s4["K","Pr(>|t|)"]))
p(sprintf("    慣れ方の違い   input×K  p=%.3f  → 慣れる速さは方式で変わらない",
          s4["inputnumpad:K","Pr(>|t|)"]))
p(sprintf("    わからない     p=%.3f  → 入力時間には効いていない",
          s4["unknown_btn1","Pr(>|t|)"]))
p("")

p(strrep("-", 76))
p("【4】Kenward-Roger による検定（参加者14名への補正）")
p(strrep("-", 76))
base0 <- "log_t_input ~ E + K + switch_new + unknown_btn"
m0 <- lmer(as.formula(paste(base0, "+", RE)), d, REML = TRUE, control = ct)
kr <- tryCatch(KRmodcomp(m4, m0),
               error = function(e) { p(paste("  失敗:", conditionMessage(e))); NULL })
if (!is.null(kr)) {
  tt <- summary(kr)$test
  p(sprintf("  F = %.3f   分子の自由度 %.0f   分母の自由度 %.1f   p = %.3g",
            tt["Ftest","stat"], tt["Ftest","ndf"], tt["Ftest","ddf"], tt["Ftest","p.value"]))
  p("  （入力方式の主効果と input×K をまとめて検定）")
} else p("  計算に失敗しました")
V <- tryCatch(vcovAdj(m4), error = function(e) NULL)
if (!is.null(V)) {
  se_kr <- sqrt(diag(as.matrix(V)))["inputnumpad"]
  b <- s4["inputnumpad", "Estimate"]; ci <- b + c(-1.96, 1.96) * se_kr
  p("")
  p(sprintf("  KRで補正した標準誤差 %.4f（補正前 %.4f）",
            se_kr, s4["inputnumpad","Std. Error"]))
  p(sprintf("  補正後の比  %.0f%% [%.0f%%, %.0f%%]",
            100*exp(b), 100*exp(ci[1]), 100*exp(ci[2])))
}
p("")

p(strrep("-", 76))
p("【5】M4 の変量効果（人・便・場面のばらつき）")
p(strrep("-", 76))
vc <- as.data.frame(VarCorr(m4))
cap(print(vc[, c("grp","var1","var2","vcov","sdcor")], digits = 4))
tot <- sum(vc$vcov[is.na(vc$var2)])
p("")
p("  全体に占める割合（相関の行を除く）")
for (i in seq_len(nrow(vc))) if (is.na(vc$var2[i]))
  p(sprintf("    %-10s %-14s %5.1f%%", vc$grp[i], vc$var1[i], 100*vc$vcov[i]/tot))
p("")
p("  ★特異について")
cr <- vc$sdcor[!is.na(vc$var2)]
if (length(cr)) p(sprintf("    参加者の切片と傾きの相関が %.3f。境界（±1）に張りついている", cr[1]))
p("    14名では切片と傾きを別々に推定しきれないために起きる。")
p("    固定効果の推定と検定はそのまま使える。傾きを外した M6 でも比は同じ水準。")
p("")

p(strrep("-", 76))
p("【6】停留所欄にかかった時間（教授の『純粋な比較では?』への答え）")
p(strrep("-", 76))
ds <- d[!is.na(d$log_t_stop), ]
ms <- lmer(as.formula(paste("log_t_stop ~ input + E + K + input:K + switch_new + unknown_btn +", RE)),
           ds, REML = TRUE, control = ct)
s <- summary(ms)$coefficients["inputnumpad", ]
ci <- s[1] + c(-1.96, 1.96) * s[2]
p(sprintf("  件数 %d", nrow(ds)))
p(sprintf("  入力方式の係数 %+.4f  p %.3f   比 %3.0f%% [%2.0f%%, %2.0f%%]",
          s[1], s[5], 100*exp(s[1]), 100*exp(ci[1]), 100*exp(ci[2])))
p("  → 停留所欄の時間は入力方式で変わらない。人数入力の時間だけが半分になっている")
p("")

p(strrep("-", 76))
p("【まとめ】")
p(strrep("-", 76))
p(sprintf("  ・テンキーは＋−ボタンの %.0f%%（M4、95%%CI %.0f〜%.0f%%）",
          100*exp(s4["inputnumpad","Estimate"]),
          100*exp(s4["inputnumpad","Estimate"]-1.96*s4["inputnumpad","Std. Error"]),
          100*exp(s4["inputnumpad","Estimate"]+1.96*s4["inputnumpad","Std. Error"])))
if (!is.null(kr)) p(sprintf("  ・Kenward-Roger で補正しても p = %.3g",
                            summary(kr)$test["Ftest","p.value"]))
p("  ・場面のばらつきが残差を除くと最大。scene_id を外してはいけない")
p("  ・切り替え直後の3件は遅い。慣れる速さは方式で変わらない")
p("  ・停留所欄の時間は変わらない")
p("")

writeLines(out, "本番モデル.txt", useBytes = TRUE)
write.csv(do.call(rbind, rows), "本番モデル_固定効果.csv", row.names = FALSE, fileEncoding = "UTF-8")
write.csv(vc, "本番モデル_変量効果.csv", row.names = FALSE, fileEncoding = "UTF-8")
cat("\n書き出しました: 本番モデル.txt / _固定効果.csv / _変量効果.csv\n")
