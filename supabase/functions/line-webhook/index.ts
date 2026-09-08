// LINE Messaging API Webhook 受信関数
// 「競り市案内」グループに投稿された写真・テキストを Supabase Storage / DB に保存する。
//
// 必要な Secrets（Supabase Dashboard → Edge Functions → Secrets）
//   LINE_CHANNEL_SECRET        … 署名検証用
//   LINE_CHANNEL_ACCESS_TOKEN  … 画像本体の取得・表示名取得用
// SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY は Supabase が自動で注入する。

import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "@supabase/supabase-js";

const BUCKET = "line-catch-reports";
const TABLE = "line_catch_reports";
const SIGNED_URL_TTL_SEC = 60 * 60 * 24 * 7; // 7日

const channelSecret = Deno.env.get("LINE_CHANNEL_SECRET") ?? "";
const accessToken = Deno.env.get("LINE_CHANNEL_ACCESS_TOKEN") ?? "";
const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

type LineSource = { type: "user" | "group" | "room"; userId?: string; groupId?: string; roomId?: string };
type LineMessage = {
  id: string;
  type: string;
  text?: string;
  fileName?: string;
  contentProvider?: { type: "line" | "external"; originalContentUrl?: string };
};
type LineEvent = {
  type: string;
  webhookEventId?: string;
  timestamp: number;
  source?: LineSource;
  message?: LineMessage;
};

// ---- 署名検証 -------------------------------------------------------------
async function verifySignature(body: string, signature: string | null): Promise<boolean> {
  if (!channelSecret || !signature) return false;
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(channelSecret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const mac = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(body));
  const expected = btoa(String.fromCharCode(...new Uint8Array(mac)));
  if (expected.length !== signature.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ signature.charCodeAt(i);
  return diff === 0;
}

// ---- LINE API ヘルパー ----------------------------------------------------
const lineHeaders = () => ({ Authorization: `Bearer ${accessToken}` });

async function fetchDisplayName(source?: LineSource): Promise<string | null> {
  if (!source?.userId || !accessToken) return null;
  const url =
    source.type === "group" && source.groupId
      ? `https://api.line.me/v2/bot/group/${source.groupId}/member/${source.userId}`
      : source.type === "room" && source.roomId
      ? `https://api.line.me/v2/bot/room/${source.roomId}/member/${source.userId}`
      : `https://api.line.me/v2/bot/profile/${source.userId}`;
  try {
    const res = await fetch(url, { headers: lineHeaders() });
    if (!res.ok) return null;
    const json = await res.json();
    return typeof json.displayName === "string" ? json.displayName : null;
  } catch {
    return null;
  }
}

async function fetchContent(message: LineMessage): Promise<{ bytes: Uint8Array; contentType: string } | null> {
  const url =
    message.contentProvider?.type === "external" && message.contentProvider.originalContentUrl
      ? message.contentProvider.originalContentUrl
      : `https://api-data.line.me/v2/bot/message/${message.id}/content`;
  const headers = message.contentProvider?.type === "external" ? {} : lineHeaders();
  const res = await fetch(url, { headers });
  if (!res.ok) {
    console.error(`content fetch failed: ${res.status} ${message.id}`);
    return null;
  }
  const contentType = res.headers.get("content-type") ?? "application/octet-stream";
  return { bytes: new Uint8Array(await res.arrayBuffer()), contentType };
}

function extFor(contentType: string, fileName?: string): string {
  if (fileName && fileName.includes(".")) return fileName.slice(fileName.lastIndexOf(".") + 1).toLowerCase();
  const map: Record<string, string> = {
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/gif": "gif",
    "image/webp": "webp",
    "video/mp4": "mp4",
    "audio/m4a": "m4a",
    "audio/x-m4a": "m4a",
    "application/pdf": "pdf",
  };
  return map[contentType.split(";")[0].trim()] ?? "bin";
}

// ---- イベント処理 ----------------------------------------------------------
async function handleEvent(ev: LineEvent): Promise<void> {
  const source = ev.source;
  const base = {
    event_type: ev.type,
    source_type: source?.type ?? null,
    group_id: source?.groupId ?? source?.roomId ?? null,
    user_id: source?.userId ?? null,
    line_timestamp: new Date(ev.timestamp).toISOString(),
  };

  // Bot のグループ参加・退出などは groupId を知るために記録しておく
  if (ev.type !== "message" || !ev.message) {
    await supabase.from(TABLE).upsert(
      { ...base, message_id: ev.webhookEventId ?? `${ev.type}-${ev.timestamp}` },
      { onConflict: "message_id", ignoreDuplicates: true },
    );
    return;
  }

  const msg = ev.message;
  const userName = await fetchDisplayName(source);
  const row: Record<string, unknown> = {
    ...base,
    message_id: msg.id,
    message_type: msg.type,
    user_name: userName,
    text: msg.text ?? null,
  };

  if (["image", "video", "audio", "file"].includes(msg.type)) {
    const content = await fetchContent(msg);
    if (content) {
      const day = row.line_timestamp!.toString().slice(0, 10);
      const path = `${base.group_id ?? base.user_id ?? "unknown"}/${day}/${msg.id}.${extFor(content.contentType, msg.fileName)}`;
      const { error: upErr } = await supabase.storage
        .from(BUCKET)
        .upload(path, content.bytes, { contentType: content.contentType, upsert: true });
      if (upErr) {
        console.error("storage upload failed", upErr);
      } else {
        row.storage_path = path;
        row.content_type = content.contentType;
        const { data: signed } = await supabase.storage.from(BUCKET).createSignedUrl(path, SIGNED_URL_TTL_SEC);
        row.signed_url = signed?.signedUrl ?? null;
      }
    }
  }

  const { error } = await supabase.from(TABLE).upsert(row, { onConflict: "message_id", ignoreDuplicates: true });
  if (error) console.error("db upsert failed", error);
}

// ---- HTTP ハンドラ ----------------------------------------------------------
Deno.serve(async (req: Request) => {
  if (req.method === "GET") {
    return new Response("line-webhook ok", { status: 200 });
  }
  if (req.method !== "POST") {
    return new Response("method not allowed", { status: 405 });
  }

  const body = await req.text();
  if (!(await verifySignature(body, req.headers.get("x-line-signature")))) {
    console.warn("invalid signature");
    return new Response("invalid signature", { status: 401 });
  }

  let events: LineEvent[] = [];
  try {
    events = (JSON.parse(body).events ?? []) as LineEvent[];
  } catch {
    return new Response("bad request", { status: 400 });
  }

  // LINE はレスポンスを待つため、保存処理はバックグラウンドで続行して即 200 を返す
  const work = Promise.allSettled(events.map(handleEvent)).then((results) => {
    for (const r of results) if (r.status === "rejected") console.error("event failed", r.reason);
  });
  // deno-lint-ignore no-explicit-any
  const rt = (globalThis as any).EdgeRuntime;
  if (rt?.waitUntil) rt.waitUntil(work);
  else await work;

  return new Response(JSON.stringify({ received: events.length }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
});
