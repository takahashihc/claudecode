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
