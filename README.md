# bus-ui-study：バス混雑度回答アプリの UI／ユーザビリティ研究（卒業研究）

NAIST共同研究用に開発した「バス混雑度 回答アプリ」を**土台（コピー）**にして、
**どの回答UIが速く・正確に・少ない操作で答えられるか**を調べるための研究用アプリと分析基盤です。

## このプロジェクトの位置づけ
| | 用途 | 場所 | 触ってよいか |
|---|---|---|---|
| **元アプリ（NAIST提出用）** | BLE人数推定の正解データ収集 | `研究\★アプリケーション\naist-bus-congestion` | **触らない**（提出物） |
| **本プロジェクト** | UI／ユーザビリティの卒業研究 | `研究\bus-ui-study` | ここで自由に改造する |

- 元アプリの **`ee61321`（phaseA-1.3）** 時点をコピーして作成。
- 研究計画・計測仕様は **`UI_STUDY_SPEC.md`** を参照。

## ⚠️ 安全設計（元アプリのデータを汚さないため）
コピー時に以下を**分離済み**です。**絶対に元に戻さないでください。**

| 項目 | 元アプリ | 本プロジェクト |
|---|---|---|
| 回答テーブル | `responses` | **`ui_responses`** |
| 一覧ビュー | `report_feed` | **`ui_report_feed`** |
| ローカル退避キー | `bus_outbox_v1` | **`uistudy_outbox_v1`** |
| APP_VERSION | `phaseA-1.3` | **`uiStudy-0.8.0`**（2026-09-21時点） |

> Supabaseの接続先（URL/キー）は元アプリと同じままです。**テーブルが別なので混ざりません**が、
> 完全分離したい場合は「別のSupabaseプロジェクトを作る」のが最も安全です（未決定）。
> **2026-09-18 に `ui_responses` / `ui_sessions` / `ui_report_feed` を作成済み**。
> テスト送信で `ui_responses` に保存されることを確認済み。

## 構成
- `index.html` … アプリ本体（**UI条件の切替・操作ログとも実装済み**）
- `stops.js` … バス停辞書736件・阪急バス（大阪府内、`window.HANKYU_STOPS`）
  （出典：国土数値情報 バス停留所データ P11／国土交通省・大阪府・令和4年度〈2022年〉・PDL1.0＝出典明記で編集加工可。
  2026-09-17 に奈良県+京都府版〈平成22年・非商用〉から差し替え。実施事業者が阪急バスに変わったため）
- `ui_responses.sql` / `ui_report_feed.sql` … Supabaseスキーマ（テーブル名を分離済み）
- `viewer.html` … 軌跡ビューア（元アプリから流用）
- `UI_STUDY_SPEC.md` … **研究計画・計測仕様（RQ／実験デザイン／ログスキーマ）**
- `EXPERIMENT_PLAN.md` … **実施計画の大枠**（奈良交通前提：人数・日数・時間帯・費用・基準値の作り方・想定問答）
- `QUESTIONNAIRE.md` … 事後アンケート
- `DESIGN_PHILOSOPHY.md` … UIの設計思想・改善シーケンス・変更台帳
- `LITERATURE.md` … 文献メモ（**本文をどこまで読んだか**を記録）
- `WORKFLOW.md` … 教授との共有・文献調査のルール
- `docs/log/` … 打ち合わせ・作業ごとの要点ログ

## これからやること（順序）
1. ~~Supabaseにテーブル作成~~ → **2026-09-18 完了**（`ui_responses` / `ui_sessions` / `ui_report_feed`）
2. **ログ計装**：操作イベント（タップ・所要時間・訂正回数など）の記録
3. **UI条件（バリアント）の割当**：セッションごとに自動で条件を切替・記録
4. **エクスポートと分析**：JSON出力 → Python(pandas)で条件間比較

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
