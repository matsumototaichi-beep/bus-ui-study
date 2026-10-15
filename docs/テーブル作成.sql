-- バス人数アプリ テーブル作成（3本まとめ）
-- Supabase の SQL Editor に貼って Run

-- ========== ui_sessions ==========
-- ============================================================
-- UI研究：操作ログ（1回答＝1セッション）
-- Supabase SQL Editor に貼り付けて実行してください。
-- ※NAIST提出用アプリの responses とは別テーブル（データは混ざりません）
-- ============================================================

create table if not exists public.ui_sessions (
  id                  bigint generated always as identity primary key,
  created_at          timestamptz not null default now(),   -- サーバ受信時刻
  surveyor_id         uuid        not null default auth.uid(),
  session_id          text,                                  -- クライアント側のセッションID
  started_at          timestamptz,                           -- セッション開始(端末時刻)
  client_submitted_at timestamptz,                           -- 実送信時の端末時刻
  outcome             text,                                  -- submit | skip | abandon
  variant             jsonb,                                 -- 割り当てたUI条件 {input, count_init, seq, offset, unknown_ui, url_fixed{input,unknown}, practice, study, round, leg, group, test_account, label,
                                                             --   login_id, url_study, leg_by, leg_at}
                                                             --   input=要因A(stepper/numpad) / unknown_ui=要因B(「わからない」ボタンの有無)
                                                             --   count_init は2026-09-19に要因から外れ、以降は常に "none"（それ以前は "median" が混在）
                                                             --   ★2026-10-10（uiStudy-0.9.6）から入る項目：
                                                             --     login_id  = 混み具合を押した時点でログインしていたID（例 "A"）。答えた人はこれで見る。
                                                             --                 surveyor_id は「送った瞬間にログインしていた人」なので、未送信の回答を別の人のログインで送ると食い違う
                                                             --     url_study = 開いたURLの指定。"r1"(?round=1) / "r2"(?round=2) / "count"(?study=count) / null(指定なし)。
                                                             --                 round と合わなければ、回はその日に覚えた値（別のQR・別のタブ）から来ている
                                                             --     leg_by    = いまの区間にした決め方。"pick"(区間を選ぶ画面で選んだ) / "switch"(かえりに切り替えた) / "back"(ゆきに戻した)
                                                             --     leg_at    = その時刻（UTC）。かえりに切り替え忘れた回答を、桃山台駅のアンケートの時刻と突き合わせて見つけるため
                                                             --     （0.9.5 までの行には、この4つは無い）
                                                             --     ★0.9.9 / counter-2.0.6 から、leg_by・leg_at は study="count" の行にも入る（人数アプリ・乗降アプリの両方）。
                                                             --     値が入るのは「かえりに切り替えた」あとだけで、ゆきの行は null（区間を選ぶ操作そのものが無いため）
  context             jsonb,                                 -- congestion_class / exact_count / count_touched / count_unknown /
                                                             --   next_stop / gps_points / screen_w / screen_h / lang / app_version /
                                                             --   posture / practice / response_cid / device_id / reason
                                                             --   （exp_code は2026-09-18、board_stop/dest_stop は2026-09-20に廃止）
                                                             --   ★2026-10-10（uiStudy-0.9.6 / counter-2.0.3）から：
                                                             --     device_id = 端末ごとのでたらめな番号（両アプリで共通。人を特定する情報は含まない）。
                                                             --                 同じIDが2台から送っている・カードの取り違えを見つけるため。
                                                             --                 端末に保存できない端末（Cookie をすべて止めた Safari など）では、開き直すたびに変わる
                                                             --     reason    = "pagehide"：人数を入れている途中でページが閉じた（outcome は abandon）。
                                                             --                 この行は同じ session_id でほぼ毎回2行届く。回答の行（人数は空）は、届かないことが多い
                                                             --                 reason が無い abandon は「戻る」を押した
  summary             jsonb,                                 -- t_total/t_class/t_count/t_first/t_stop/taps_total/taps_value/corrections
  events              jsonb                                  -- 操作イベント列 [{dt,type,...}]
);
-- ★2026-10-07: 計数研究かえり（乗降アプリ counter.html）の行も入る。1バス停＝1行。スキーマは同じ
--   variant {study:"count", leg:"kaeri", method:"flow", app:"counter", ...}
--   context {run_id, stop{name,seq}, board, alight, resend, ...}。resend:true があれば同じ run_id・stop の最後の行を使う
--   ★2026-10-10（counter-2.0.3）から variant.login_id（送信を押した時点のID）と context.device_id も入る

create index if not exists ui_sessions_created_idx  on public.ui_sessions (created_at desc);
create index if not exists ui_sessions_surveyor_idx on public.ui_sessions (surveyor_id);

alter table public.ui_sessions enable row level security;

drop policy if exists "ui insert own" on public.ui_sessions;
create policy "ui insert own" on public.ui_sessions
  for insert to authenticated with check (surveyor_id = auth.uid());

drop policy if exists "ui select own" on public.ui_sessions;
create policy "ui select own" on public.ui_sessions
  for select to authenticated using (surveyor_id = auth.uid());

-- 分析は SQL Editor（service_role）で全件参照:
--   select json_agg(t) from (select * from ui_sessions order by id) t;
-- 同じ session_id の行が2行入ることがある（送れたのに返事が届かず、あとで送り直したとき。人数を入れている途中で
-- ページが閉じたときは、ほぼ毎回）。session_id ごとに1行にまとめる。
--   ui_responses も同じで、client_id ごとに1行にまとめる

-- ========== ui_responses ==========
-- ============================================================
-- バス混雑度 HITL データ収集  Supabase スキーマ (Phase A)
-- Supabase ダッシュボード > SQL Editor に貼り付けて実行してください。
-- PROJECT_SPEC.md §6 responses に対応。
-- ============================================================

-- 生の回答（不変）
create table if not exists public.ui_responses (
  id               bigint generated always as identity primary key,
  created_at       timestamptz not null default now(),        -- サーバ受信時刻
  surveyor_id      uuid        not null default auth.uid(),    -- 認証ユーザーのUUID
  answered_at      timestamptz,                               -- 端末での回答時刻(TZ付き)
  congestion_class smallint    not null,                       -- 1〜6
  exact_count      integer,                                    -- 実測人数(任意, null可)
  trajectory       jsonb,                                      -- 回答中〜直後の短い軌跡 [{t,lat,lng,acc},...]
  next_stop        jsonb,                                      -- ★次に停まるバス停 {name,lat,lng}(任意)
  board_stop       jsonb,                                      -- 【廃止】乗車バス停。2026-09-20以降は書き込まれない
  dest_stop        jsonb,                                      -- 【廃止】行先バス停。同上
  lat              double precision,                           -- 代表位置(緯度) answered_at最寄りの軌跡点
  lng              double precision,                           -- 代表位置(経度)
  gps_accuracy     real,                                       -- 代表位置の精度(m)
  gps_fixed_at     timestamptz,                                -- 代表位置の測位時刻
  client_submitted_at timestamptz,                             -- 実際に送信した端末時刻(時計ズレ検出用)
  session_id       text,                                       -- ★UI研究: ui_sessions.session_id と結合するキー
  client_id        text,                                       -- 端末側の回答ID(cid)
  practice         boolean     not null default false,         -- ★UI研究: 練習回答。分析では除外する
  lang             text,
  app_version      text,
  constraint congestion_class_range check (congestion_class between 1 and 6),
  constraint exact_count_nonneg     check (exact_count is null or exact_count >= 0)
);

-- 既存テーブルに後から列を足す場合（冪等）:
-- ★2026-09-20：next_stop を追加。乗車/降車バス停（board_stop/dest_stop）の入力は廃止した。
--   理由：あの2項目はNAIST側の目的（BLEセンサーの位置情報紐づけ。バスターミナルなど多数のバスが
--   重なる地点でスマホのGPSだけでは便を識別できないため）のための項目で、BLEセンサーを設置していない
--   阪急バスでは意味を持たない。代わりに「次に停まるバス停」を聞き、
--   自己申告とGPS位置の一致度から**位置情報がどれだけ当てになるか**を測る。
-- 旧2列は 2026-09-18 の先行実験データが入っているので **drop しない**。
alter table public.ui_responses add column if not exists next_stop  jsonb;
alter table public.ui_responses add column if not exists board_stop jsonb;
alter table public.ui_responses add column if not exists dest_stop  jsonb;
alter table public.ui_responses add column if not exists lat double precision;
alter table public.ui_responses add column if not exists lng double precision;
alter table public.ui_responses add column if not exists gps_accuracy real;
alter table public.ui_responses add column if not exists gps_fixed_at timestamptz;
alter table public.ui_responses add column if not exists client_submitted_at timestamptz;
alter table public.ui_responses add column if not exists session_id text;
alter table public.ui_responses add column if not exists client_id  text;
alter table public.ui_responses add column if not exists practice boolean not null default false;
create index if not exists ui_responses_session_idx on public.ui_responses (session_id);

-- 分析で本番データだけを見るとき：
--   select * from public.ui_responses where practice = false;

create index if not exists responses_created_at_idx on public.ui_responses (created_at desc);
create index if not exists responses_surveyor_idx   on public.ui_responses (surveyor_id);

-- 行レベルセキュリティ
alter table public.ui_responses enable row level security;

-- ログイン済みユーザーは「自分の回答のみ」INSERT可能
drop policy if exists "auth insert own" on public.ui_responses;
create policy "auth insert own"
  on public.ui_responses
  for insert
  to authenticated
  with check (surveyor_id = auth.uid());

-- 自分の回答は閲覧可（本人の記録数=インセンティブ集計や画面表示用）
drop policy if exists "auth select own" on public.ui_responses;
create policy "auth select own"
  on public.ui_responses
  for select
  to authenticated
  using (surveyor_id = auth.uid());

-- 集計・全件閲覧は SQL Editor / service_role キーで行う（匿名SELECTは不可）。

-- ============================================================
-- 動作確認（任意, SQL Editorはservice_role実行なのでRLSを通らない点に注意）:
--   select count(*) from public.ui_responses;
-- ============================================================

-- ========== ui_report_feed ==========
-- ============================================================
-- データ一覧（管理者用）ビュー
-- responses から「安全な列だけ」を、管理者のみに見せる。
-- 軌跡(trajectory)・surveyor_id・メール等の個人情報は含めない。
-- ★アプリ側 ADMIN_EMAILS と同じメールをここにも列挙すること。
-- Supabase SQL Editor に貼って実行してください。
-- ============================================================

create or replace view public.ui_report_feed as
  select
    id,
    created_at,
    answered_at,
    congestion_class,
    exact_count,
    app_version
  from public.ui_responses
  -- 管理者のみ行を返す（それ以外のログインユーザーは0件）。UI制限だけでなくサーバー側でも制限。
  where auth.email() in ('taichi1104.deters@gmail.com', 'admin@id.local');

-- 匿名は不可、ログイン済みにSELECT権限（実際に見えるのは上のWHEREを満たす管理者のみ）
revoke select on public.ui_report_feed from anon;
grant  select on public.ui_report_feed to authenticated;

-- 反映されない場合は PostgREST にスキーマ再読込を通知
notify pgrst, 'reload schema';

