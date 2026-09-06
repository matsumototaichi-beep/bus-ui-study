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
