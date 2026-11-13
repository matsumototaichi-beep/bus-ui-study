* ===========================================================
* バス車内人数入力アプリ UI研究 － 混合効果モデル
* 教授の資料「混合効果モデルの考え方」4節のモデル案に対応
*
* 先に ui_spss.csv を読み込んでおくこと（手順.md の 1 を参照）
* ===========================================================

* --- 1) まず変量切片だけ。いちばん単純な形で通るか確かめる -------------.
MIXED log_t_input BY input
  /FIXED = input | SSTYPE(3)
  /RANDOM = INTERCEPT | SUBJECT(pid) COVTYPE(VC)
  /METHOD = REML
  /PRINT = SOLUTION TESTCOV.

* --- 2) 変量傾きを足す（教授の指摘。p値が100倍ほど大きくなるはず）-------.
MIXED log_t_input BY input
  /FIXED = input | SSTYPE(3)
  /RANDOM = INTERCEPT input | SUBJECT(pid) COVTYPE(UN)
  /METHOD = REML
  /PRINT = SOLUTION TESTCOV.

* --- 3) 慣れと切り替えを足す ------------------------------------------.
MIXED log_t_input BY input unknown_btn switch_new WITH seq_all seq_method
  /FIXED = input seq_all seq_method input*seq_method switch_new unknown_btn
           | SSTYPE(3)
  /RANDOM = INTERCEPT input | SUBJECT(pid) COVTYPE(UN)
  /METHOD = REML
  /PRINT = SOLUTION TESTCOV.

* --- 4) 便と場面の変量効果も足す（教授の案の全部）----------------------.
*     ★場面 scene_id は必ず入れること（05_変量効果の必要性.py）.
*     同じ便の同じ停留所にいた2人の残差は相関している（ICC 15%, p<0.001）.
*     外すとテンキーの係数の誤差が小さく出すぎる（p値が甘くなる）.
*     どうしても収束しないときだけ外し、原稿にその旨を書くこと.
MIXED log_t_input BY input unknown_btn switch_new WITH seq_all seq_method
  /FIXED = input seq_all seq_method input*seq_method switch_new unknown_btn
           | SSTYPE(3)
  /RANDOM = INTERCEPT input | SUBJECT(pid) COVTYPE(UN)
  /RANDOM = INTERCEPT | SUBJECT(run_id) COVTYPE(VC)
  /RANDOM = INTERCEPT | SUBJECT(scene_id) COVTYPE(VC)
  /METHOD = REML
  /PRINT = SOLUTION TESTCOV.

* --- 5) 少人数向けの自由度の補正（Kenward-Roger）-----------------------.
*     ★ここが教授の指摘した部分.
*     ★下の DFMETHOD の書き方は SPSS の版によって違う.
*     ★エラーが出たら、ヘルプで MIXED の CRITERIA を調べて直すこと.
MIXED log_t_input BY input unknown_btn switch_new WITH seq_all seq_method
  /FIXED = input seq_all seq_method input*seq_method switch_new unknown_btn
           | SSTYPE(3)
  /RANDOM = INTERCEPT input | SUBJECT(pid) COVTYPE(UN)
  /RANDOM = INTERCEPT | SUBJECT(run_id) COVTYPE(VC)
  /METHOD = REML
  /CRITERIA = DFMETHOD(KENWARDROGER)
  /PRINT = SOLUTION TESTCOV.

* --- 6) 停留所欄の時間にも同じモデルを当てる --------------------------.
MIXED log_t_stop BY input unknown_btn switch_new WITH seq_all seq_method
  /FIXED = input seq_all seq_method input*seq_method switch_new unknown_btn
           | SSTYPE(3)
  /RANDOM = INTERCEPT input | SUBJECT(pid) COVTYPE(UN)
  /RANDOM = INTERCEPT | SUBJECT(run_id) COVTYPE(VC)
  /METHOD = REML
  /PRINT = SOLUTION TESTCOV.
