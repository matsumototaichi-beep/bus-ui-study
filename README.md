# bus-ui-study：バス混雑度回答アプリの UI／ユーザビリティ研究（卒業研究）

NAIST共同研究用に開発した「バス混雑度 回答アプリ」を**土台（コピー）**にして、
**どの回答UIが速く・少ない操作で答えられるか**を調べるための研究用アプリと分析基盤です。

## このプロジェクトの位置づけ
| | 用途 | 場所 | 触ってよいか |
|---|---|---|---|
| **元アプリ（NAIST提出用）** | BLE人数推定の正解データ収集 | `研究\★アプリケーション\naist-bus-congestion` | **触らない**（提出物） |
| **本プロジェクト** | UI／ユーザビリティの卒業研究 | `卒業研究\★バス人数計測アプリUI研究\bus-ui-study` | ここで自由に改造する |

- 元アプリの **`ee61321`（phaseA-1.3）** 時点をコピーして作成。
- 今の計画は **`HANDOFF.md` §8**（決定の表）と **`docs/実験の進め方.docx`**。

## ⚠️ 安全設計（元アプリのデータを汚さないため）
コピー時に以下を**分離済み**です。**絶対に元に戻さないでください。**

| 項目 | 元アプリ | 本プロジェクト |
|---|---|---|
| 回答テーブル | `responses` | **`ui_responses`** |
| 一覧ビュー | `report_feed` | **`ui_report_feed`** |
| ローカル退避キー | `bus_outbox_v1` | **`uistudy_outbox_v1`** |
| APP_VERSION | `phaseA-1.3` | **`uiStudy-0.9.7`**（2026-10-11時点。乗降アプリは `counter-2.0.3`） |

> **2026-10-06 に Supabase を元アプリとは別のプロジェクト（`wwavbfdojschxohvvpzb`）に切り替えた**（前のプロジェクトは削除されていた）。
> テーブルは `docs/テーブル作成.sql`（`ui_responses` / `ui_sessions` / `ui_report_feed`）で作る。
> 2026-09-18 に旧プロジェクトで作成し、テスト送信で `ui_responses` に保存されることを確認していた。

## どこに何があるか
| 場所 | 中身 |
|---|---|
| `index.html` | 人数アプリ（UI研究・計数研究のゆき） |
| `counter.html` | 乗降アプリ（計数研究のかえり） |
| `stops.js` | バス停辞書736件・阪急バス（`window.HANKYU_STOPS`） |
| `viewer.html` | 軌跡ビューア（元アプリから流用） |
| `HANDOFF.md` | 引き継ぎ。§8 が決定の表、§9 が次にやること |
| `UI_STUDY_SPEC.md` | 研究の問い・ログ仕様・分析計画 |
| `DESIGN_PHILOSOPHY.md` | UIの設計思想・変更台帳（変えたことの記録） |
| `LITERATURE.md` | 文献メモ（本文をどこまで読んだか） |
| `WORKFLOW.md` | 教授との共有・文献調査のルール |
| `docs/実験の進め方.docx` | 当日の運用 |
| `docs/questionnaire/` | UI研究のアンケート4つ（1回目_桃山台・1回目_南口・2回目_桃山台・2回目_南口）。**正本は Google フォーム**（2026-10-10 に松本が手で直した）。`フォームの中身_2026-10-11.md`・`.json` がその控え、`アンケート_*.docx` が紙の予備。`設問一覧.xlsx`・`GoogleForms生成.gs` は作ったときの版で、今のフォームと違う（**`createAll` は2度と実行しない**） |
| `docs/分析の準備.md` | アンケートとアプリの記録を書き出して、1つの表につなぐ手順 |
| `docs/login_prep.md` / `docs/login_paths.md` | ログインの準備／管理者ログイン |
| `docs/テーブル作成.sql` | Supabase のテーブル |
| `tools/` | 下の表 |
| `docs/log/` | 日ごとの記録（`2026-09-12.md` は教授との打ち合わせ） |
| `docs/old/` | 今は使わない古い文書（一覧は `docs/old/README.md`。バス会社の承諾条件は `docs/old/HANDOFF_hankyu.md` §1） |
| `exports/` | 下見（2026-09-18）のデータ。個人に関わるので git に入れない（`.gitignore`） |
| `../★乗降人数計算/` | 計数研究（1から数える vs 乗降で数える）の案内・配布用紙・QR・アンケート・アンケートURL.xlsx・ショートカット |
| `../★UI研究/` | 参加者向け案内_Ui研究.docx（松本が直接編集）と .pdf・配布用紙（1回目・2回目）・QRポスター.pdf・ログインカード.pdf・accounts.csv・アンケートURL.xlsx・人数アプリのショートカット |

| `tools/` | 作るもの |
|---|---|
| `make_operation_docx.py` | `docs/実験の進め方.docx` |
| `make_posters_pdf.py` | `../★UI研究/QRポスター.pdf` |
| `make_cards_pdf.py` | `../★UI研究/ログインカード.pdf` |
| `create_accounts.py` | 参加者アカウント・`../★UI研究/accounts.csv` |
| `read_forms.py` | UI研究のアンケートの今の中身を読み（読むだけ）、控え `docs/questionnaire/フォームの中身_*.json`・`.md` と比べる・作り直す |
| `make_questionnaire.py` | 紙の予備 `docs/questionnaire/アンケート_*.docx`（今の中身の控えから作る） |
| `make_forms.py` | 作ったときの版（2026-10-09）の設問一覧.xlsx・GoogleForms生成.gs。今のフォームとは違う |
| `join_survey_app.py` | アンケートの答えとアプリの記録をつないだ表と、問題の一覧（`docs/分析の準備.md`）。確かめは `test_join_survey_app.py` |
| `fetch_weather.py` | 回答した時刻の天気（気象庁アメダス豊中の10分値・Open-Meteo）を、あとで取って回答に付ける。書き出しは `exports/`。確かめは `test_fetch_weather.py` |
| `make_count_qr.py` | `../★乗降人数計算/QRポスター_計数研究.pdf`・`乗降アプリQR.png` |
| `make_count_forms.py` | `../★乗降人数計算/計数研究アンケート生成.gs`・`計数研究アンケート.docx` |
| `make_handouts.py` | 配布用紙（`../★UI研究/` に1回目・2回目、`../★乗降人数計算/` に計数研究）。UI研究の用紙には桃山台駅と南口のアンケートのQRが載る。アンケートURLは各フォルダの `アンケートURL.xlsx` から読む |

`stops.js` の出典：国土数値情報 バス停留所データ P11／国土交通省・大阪府・令和4年度〈2022年〉・PDL1.0＝出典明記で編集加工可。
2026-09-17 に奈良県+京都府版〈平成22年・非商用〉から差し替え。実施事業者が阪急バスに変わったため。
2026-09-19 に、バス停1つの名前の変更と1つの追加を入れた（`DESIGN_PHILOSOPHY.md` §5）。

## これからやること
→ `HANDOFF.md` §9

## 公開URL
**2026-09-18 に公開**：https://matsumototaichi-beep.github.io/bus-ui-study/（GitHub Pages・HTTPS）

## Git運用
元アプリ（`naist-bus-congestion`）とは**別の新規リポジトリ**（`matsumototaichi-beep/bus-ui-study`、public）
として2026-09-18に公開。元アプリのリモートには**絶対にpushしない**（別リポジトリなので誤って混ざる心配はない）。

```bash
git log --oneline -5   # 作業開始時：前回までの変更を確認
git add -A && git commit -m "変更内容"   # 作業終了時：必ずコミット
git push origin master # 公開先(GitHub Pages)に反映。実行前に必ず確認を取ること
```
