-- Google ニュース収集 Edge Function (collect-google-news) を毎朝 7:00 JST に実行する
-- pg_cron は UTC で動くため 22:00 UTC = 翌 07:00 JST

-- news_items: ヒットした検索クエリを記録する列と検索用インデックス
alter table public.news_items add column if not exists query text;
comment on column public.news_items.query is 'ヒットした Google ニュース検索クエリ';
create index if not exists news_items_theme_published_idx
  on public.news_items (theme, published_at desc);

-- 旧ジョブ（1テーマのみ・08:00 JST）は停止
select cron.unschedule('collect-food-packaging-news-daily')
 where exists (select 1 from cron.job where jobname = 'collect-food-packaging-news-daily');

-- 新ジョブ: 5テーマを毎朝 7:00 JST に収集
-- Authorization の anon キーは Vault に保存したものを参照する
select cron.schedule(
  'collect-google-news-daily',
  '0 22 * * *',
  $$
  select net.http_post(
    url := 'https://hjthckukwqskycthzqkk.supabase.co/functions/v1/collect-google-news',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'Authorization', 'Bearer ' || (select decrypted_secret from vault.decrypted_secrets where name = 'anon_key')
    ),
    body := '{}'::jsonb,
    timeout_milliseconds := 150000
  );
  $$
);
