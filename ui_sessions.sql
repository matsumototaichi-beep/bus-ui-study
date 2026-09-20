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
  variant             jsonb,                                 -- 割り当てたUI条件 {input, count_init, unknown_ui, seq, offset, url_fixed, practice}
                                                             --   input=要因A(stepper/numpad) / unknown_ui=要因B(「わからない」ボタンの有無)
                                                             --   count_init は2026-09-19に要因から外れ、以降は常に "none"（それ以前は "median" が混在）
  context             jsonb,                                 -- congestion_class / exact_count / count_touched / count_unknown /
                                                             --   next_stop / gps_points / screen_w / screen_h / lang / app_version /
                                                             --   posture / practice / response_cid
                                                             --   （exp_code は2026-09-18、board_stop/dest_stop は2026-09-20に廃止）
  summary             jsonb,                                 -- t_total/t_class/t_count/t_first/t_stop/taps_total/taps_value/corrections
  events              jsonb                                  -- 操作イベント列 [{dt,type,...}]
);

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
