# Supabase 学び記録テーブル

プロジェクト: takahashihc's Project（`hjthckukwqskycthzqkk`, ap-northeast-1）

## テーブル `learning_log`

| 列名 | 型 | 内容 |
|---|---|---|
| date | date | 日付（未指定なら当日） |
| theme | text | テーマ |
| title | text | 学んだこと（タイトル） |
| insight | text | 気づき（詳細） |

`id`（uuid）と `created_at`（登録日時）は自動で入ります。
RLS 有効。ログイン済み（authenticated）ユーザーのみ読み書き可能。

## 登録例（SQL）

```sql
insert into public.learning_log (date, theme, title, insight)
values ('2026-09-03', '経営戦略', '学んだことのタイトル', '気づきの詳細');
```

## 一覧（新しい順）

```sql
select date, theme, title, insight
from public.learning_log
order by date desc, created_at desc;
```

## Google ニュース自動収集（毎朝 7:00）

- Edge Function: `supabase/functions/collect-google-news/index.ts`
- スケジュール: pg_cron ジョブ `collect-google-news-daily`（`0 22 * * *` UTC = 毎朝 7:00 JST）
- 保存先: `public.news_items`（`theme` + `link` で重複排除）
- 取得範囲: 直近 2 日分（Google ニュース RSS の `when:2d`）

### 収集テーマと検索クエリ

| theme | 主なクエリ |
|---|---|
| 原料価格（原油・ナフサ・樹脂） | 原油 価格 / ナフサ 価格 / PS樹脂 / PET樹脂 / ポリエチレン 値上げ など |
| 食品包装業界 | 食品 包装資材 / 食品包装 容器 / 包装 業界 動向 など |
| 食品業界（取引先） | 食品業界 / 食品メーカー 値上げ / 惣菜・弁当 業界 / 食品スーパー 出店 など |
| 地域（島根・山口・鳥取・佐賀） | 各県 企業 OR 経済 / 島根・鳥取 食品 / 山口・佐賀 食品 |
| その他タカハシ包装センター関連 | タカハシ包装センター / 包装資材 商社・卸 / 食品 物流 人手不足 / 容器包装リサイクル法 |

クエリを変えたいときは `index.ts` の `THEMES` を編集して再デプロイする。

### 手動実行（DB から）

```sql
select net.http_post(
  url := 'https://hjthckukwqskycthzqkk.supabase.co/functions/v1/collect-google-news',
  headers := jsonb_build_object(
    'Content-Type', 'application/json',
    'Authorization', 'Bearer ' || (select decrypted_secret from vault.decrypted_secrets where name = 'anon_key')
  ),
  body := '{}'::jsonb,
  timeout_milliseconds := 150000
);
```

### 実行状況の確認

```sql
-- cron の実行履歴
select jobid, status, start_time, end_time, return_message
from cron.job_run_details order by start_time desc limit 10;

-- 今日取り込まれた件数（テーマ別）
select theme, count(*) from public.news_items
where collected_at >= current_date group by theme;
```
