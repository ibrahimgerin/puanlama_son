import { Redis } from "@upstash/redis";

const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
const redis = url && token ? new Redis({ url, token }) : null;

function unwrap(stored) {
  if (stored === null || stored === undefined) return null;
  const parsed = typeof stored === "string" ? JSON.parse(stored) : stored;
  return parsed && typeof parsed === "object" && "raw" in parsed ? parsed.raw : null;
}

export default async function handler(req, res) {
  if (!redis) {
    return res.status(500).json({
      error: "Veritabanı bağlı değil. Vercel → Storage → Upstash Redis bağla, sonra redeploy et.",
    });
  }
  try {
    if (req.method === "GET") {
      // Bulk read: supports one or more comma-separated prefixes in a single
      // request, e.g. ?bulk=r:,pr:  — this lets the frontend combine several
      // polls into one HTTP round trip, which matters a lot once many
      // concurrent users are polling at once.
      if (req.query.bulk !== undefined) {
        const prefixes = (req.query.bulk || "")
          .toString()
          .split(",")
          .map((p) => p.trim())
          .filter(Boolean);
        if (prefixes.length === 0) return res.status(200).json({ entries: {} });
        const keySets = await Promise.all(prefixes.map((p) => redis.keys(`${p}*`)));
        const keys = [...new Set(keySets.flat())];
        if (keys.length === 0) return res.status(200).json({ entries: {} });
        const values = await redis.mget(...keys);
        const entries = {};
        keys.forEach((k, i) => {
          const v = unwrap(values[i]);
          if (v !== null) entries[k] = v;
        });
        return res.status(200).json({ entries });
      }

      // List keys: GET /api/storage?list=prefix
      if (req.query.list !== undefined) {
        const prefix = (req.query.list || "").toString();
        const keys = await redis.keys(`${prefix}*`);
        return res.status(200).json({ keys });
      }

      // Single read: GET /api/storage?key=xxx
      const key = (req.query.key || "").toString();
      if (!key) return res.status(400).json({ error: "key gerekli" });
      const value = unwrap(await redis.get(key));
      return res.status(200).json({ value });
    }

    if (req.method === "POST") {
      const { key, value } = req.body || {};
      if (!key) return res.status(400).json({ error: "key gerekli" });
      await redis.set(key, JSON.stringify({ raw: value }));
      return res.status(200).json({ ok: true });
    }

    if (req.method === "DELETE") {
      const key = (req.query.key || "").toString();
      if (!key) return res.status(400).json({ error: "key gerekli" });
      await redis.del(key);
      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ error: "Desteklenmeyen metod" });
  } catch (err) {
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
