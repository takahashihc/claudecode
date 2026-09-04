// Google ニュース RSS から 5 テーマの最新ニュースを取得し、news_items に保存する。
// pg_cron から毎朝 7:00 (JST) に呼び出される。
import { createClient } from "npm:@supabase/supabase-js@2";

// 直近何日分を取得するか（Google ニュースの when: 演算子）
const LOOKBACK_DAYS = 2;

// テーマごとの検索クエリ。1 クエリ = 1 回の RSS 取得。
const THEMES: { theme: string; queries: string[] }[] = [
  {
    theme: "原料価格（原油・ナフサ・樹脂）",
    queries: [
      "原油 価格",
      "ナフサ 価格",
      "ポリスチレン OR PS樹脂 価格",
      "PET樹脂 OR ペット樹脂 価格",
      "ポリエチレン OR ポリプロピレン 値上げ",
      "プラスチック 原料 値上げ",
      "石油化学 価格 改定",
    ],
  },
  {
    theme: "食品包装業界",
    queries: [
      "食品 包装資材",
      "食品包装 容器",
      "包装 業界 動向",
      "プラスチック容器 食品",
      "紙容器 OR 環境配慮 包装",
    ],
  },
  {
    theme: "食品業界（取引先）",
    queries: [
      "食品業界",
      "食品メーカー 値上げ",
      "食品 工場 新設",
      "惣菜 OR 弁当 業界",
      "食品スーパー 出店",
      "水産加工 OR 菓子メーカー",
    ],
  },
  {
    theme: "地域（島根・山口・鳥取・佐賀）",
    queries: [
      "島根県 企業 OR 経済",
      "山口県 企業 OR 経済",
      "鳥取県 企業 OR 経済",
      "佐賀県 企業 OR 経済",
      "島根 OR 鳥取 食品",
      "山口 OR 佐賀 食品",
    ],
  },
  {
    theme: "その他タカハシ包装センター関連",
    queries: [
      "タカハシ包装センター",
      "包装資材 商社",
      "包装資材 卸",
      "食品 物流 2024年問題 OR 人手不足",
      "容器包装 リサイクル法",
    ],
  },
];

function rssUrl(query: string): string {
  const q = `${query} when:${LOOKBACK_DAYS}d`;
  return `https://news.google.com/rss/search?q=${encodeURIComponent(q)}&hl=ja&gl=JP&ceid=JP:ja`;
}

function decodeEntities(str: string): string {
  return str
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

function stripCdata(str: string): string {
  return str.replace(/^<!\[CDATA\[/, "").replace(/\]\]>$/, "");
}

function stripTags(html: string): string {
  return html.replace(/<[^>]*>/g, "").trim();
}

interface RssItem {
  title: string;
  link: string;
  pubDate: string | null;
  source: string | null;
  summary: string | null;
}

function parseRssItems(xml: string): RssItem[] {
  const items: RssItem[] = [];
  const itemBlocks = xml.split("<item>").slice(1);

  for (const block of itemBlocks) {
    const titleMatch = block.match(/<title>([\s\S]*?)<\/title>/);
    const linkMatch = block.match(/<link>([\s\S]*?)<\/link>/);
    const pubDateMatch = block.match(/<pubDate>([\s\S]*?)<\/pubDate>/);
    const sourceMatch = block.match(/<source[^>]*>([\s\S]*?)<\/source>/);
    const descMatch = block.match(/<description>([\s\S]*?)<\/description>/);

    if (!titleMatch || !linkMatch) continue;

    items.push({
      title: decodeEntities(stripCdata(titleMatch[1])).trim(),
      link: decodeEntities(stripCdata(linkMatch[1])).trim(),
      pubDate: pubDateMatch ? pubDateMatch[1].trim() : null,
      source: sourceMatch ? decodeEntities(stripCdata(sourceMatch[1])).trim() : null,
      summary: descMatch ? stripTags(decodeEntities(stripCdata(descMatch[1]))).trim() : null,
    });
  }

  return items;
}

async function fetchRss(url: string): Promise<string> {
  const headers = {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    Accept: "application/rss+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7",
  };

  let lastError: unknown = null;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const res = await fetch(url, { headers });
      if (res.ok) return await res.text();
      lastError = new Error(`RSS fetch failed: ${res.status}`);
    } catch (err) {
      lastError = err;
    }
    await new Promise((r) => setTimeout(r, 1500 * (attempt + 1)));
  }
  throw lastError;
}

function toIso(pubDate: string | null): string | null {
  if (!pubDate) return null;
  const d = new Date(pubDate);
  return isNaN(d.getTime()) ? null : d.toISOString();
}

Deno.serve(async (req: Request) => {
  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );

  // 任意: {"themes": ["食品包装業界"]} で対象テーマを絞れる（テスト用）
  let onlyThemes: string[] | null = null;
  try {
    const body = await req.json();
    if (Array.isArray(body?.themes)) onlyThemes = body.themes;
  } catch {
    // body なしは通常運用
  }

  const results: Record<string, { fetched: number; inserted: number; errors: string[] }> = {};

  for (const { theme, queries } of THEMES) {
    if (onlyThemes && !onlyThemes.includes(theme)) continue;
    const summary = { fetched: 0, inserted: 0, errors: [] as string[] };
    results[theme] = summary;

    // 同一テーマ内の重複を除く（複数クエリで同じ記事が出るため）
    const seen = new Map<string, RssItem & { query: string }>();

    for (const query of queries) {
      try {
        const xml = await fetchRss(rssUrl(query));
        const items = parseRssItems(xml);
        summary.fetched += items.length;
        for (const item of items) {
          if (!seen.has(item.link)) seen.set(item.link, { ...item, query });
        }
      } catch (err) {
        summary.errors.push(`${query}: ${String(err)}`);
      }
      // Google 側の連続アクセス制限を避けるための小休止
      await new Promise((r) => setTimeout(r, 400));
    }

    const rows = [...seen.values()].map((item) => ({
      theme,
      title: item.title,
      link: item.link,
      source: item.source,
      published_at: toIso(item.pubDate),
      summary: item.summary,
      query: item.query,
    }));

    if (rows.length > 0) {
      const { data, error } = await supabase
        .from("news_items")
        .upsert(rows, { onConflict: "theme,link", ignoreDuplicates: true })
        .select("id");
      if (error) summary.errors.push(`upsert: ${error.message}`);
      else summary.inserted = data?.length ?? 0;
    }
  }

  const hasError = Object.values(results).some((r) => r.errors.length > 0);
  return new Response(
    JSON.stringify({ ok: !hasError, ran_at: new Date().toISOString(), results }, null, 2),
    { status: hasError ? 207 : 200, headers: { "Content-Type": "application/json" } },
  );
});
