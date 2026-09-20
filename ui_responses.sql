-- ============================================================
-- バス混雑度 HITL データ収集  Supabase スキーマ (Phase A)
-- Supabase ダッシュボード > SQL Editor に貼り付けて実行してください。
-- PROJECT_SPEC.md §6 responses に対応。
-- ============================================================

-- 生の回答（不変）
create table if not exists public.ui_responses (
  id               bigint generated always as identity primary key,
  created_at       timestamptz not null default now(),        -- サーバ受信時刻
  surveyor_id      uuid        not null default auth.uid(),    -- 認証ユーザー(GitHub)のUUID
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
