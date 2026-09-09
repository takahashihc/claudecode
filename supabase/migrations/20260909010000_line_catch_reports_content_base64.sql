-- 写真本体の base64 コピー（Claude が SQL 経由で読み取るため）。7日で自動的に破棄する。
alter table public.line_catch_reports
  add column if not exists content_base64 text,
  add column if not exists content_bytes integer;

comment on column public.line_catch_reports.content_base64 is '受信ファイルの base64。line_timestamp から7日経過で null に戻す';
