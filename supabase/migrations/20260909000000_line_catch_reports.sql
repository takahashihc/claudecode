-- LINE「競り市案内」グループから届く漁獲情報（写真・テキスト）の受信ログ
create table if not exists public.line_catch_reports (
  id             bigint generated always as identity primary key,
  message_id     text not null unique,            -- LINE message id（join等は webhookEventId）
  event_type     text not null,                   -- message / join / leave / memberJoined ...
  message_type   text,                            -- image / text / file / video ...
  source_type    text,                            -- group / room / user
  group_id       text,
  user_id        text,
  user_name      text,
  text           text,
  storage_path   text,                            -- Storage bucket 内のパス
  content_type   text,
  signed_url     text,                            -- 取得用の署名付きURL（7日有効）
  line_timestamp timestamptz not null,            -- LINE 側の送信時刻
  created_at     timestamptz not null default now()
);

comment on table public.line_catch_reports is 'LINE Messaging API Webhook で受信した漁獲情報（大田・競り市案内グループ）';

create index if not exists line_catch_reports_line_timestamp_idx
  on public.line_catch_reports (line_timestamp desc);
create index if not exists line_catch_reports_group_idx
  on public.line_catch_reports (group_id, line_timestamp desc);

-- Edge Function は service_role で書き込むため、匿名・認証ユーザー向けポリシーは作らない
alter table public.line_catch_reports enable row level security;

-- 受信した写真の保存先（非公開バケット）
insert into storage.buckets (id, name, public, file_size_limit)
values ('line-catch-reports', 'line-catch-reports', false, 20971520)
on conflict (id) do nothing;
