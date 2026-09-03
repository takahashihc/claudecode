-- 自分の学び記録テーブル（Supabase プロジェクト: takahashihc's Project）
-- 4列: 日付 / テーマ / 学んだこと（タイトル） / 気づき（詳細）

create table if not exists public.learning_log (
  id         uuid primary key default gen_random_uuid(),
  date       date not null default current_date,        -- 日付
  theme      text not null,                              -- テーマ
  title      text not null,                              -- 学んだこと（タイトル）
  insight    text,                                       -- 気づき（詳細）
  created_at timestamptz not null default now()
);

comment on table  public.learning_log            is '自分の学びの記録';
comment on column public.learning_log.date       is '日付（学んだ日。未指定なら当日）';
comment on column public.learning_log.theme      is 'テーマ（例: 経営戦略、組織づくり）';
comment on column public.learning_log.title      is '学んだこと（タイトル）';
comment on column public.learning_log.insight    is '気づき（詳細）';
comment on column public.learning_log.created_at is '登録日時（自動）';

create index if not exists learning_log_date_idx on public.learning_log (date desc);

alter table public.learning_log enable row level security;

drop policy if exists "allow all for authenticated users" on public.learning_log;
create policy "allow all for authenticated users"
  on public.learning_log for all to authenticated
  using (true) with check (true);
