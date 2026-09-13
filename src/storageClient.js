const BASE = "/api/storage";

async function get(key) {
  const res = await fetch(`${BASE}?key=${encodeURIComponent(key)}`);
  if (!res.ok) throw new Error("storage get failed: " + res.status);
  const data = await res.json();
  if (data.value === null || data.value === undefined) throw new Error("not found");
  return { key, value: data.value, shared: true };
}

async function set(key, value) {
  const res = await fetch(BASE, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ key, value }),
  });
  if (!res.ok) return null;
  return { key, value, shared: true };
}

async function list(prefix) {
  const res = await fetch(`${BASE}?list=${encodeURIComponent(prefix || "")}`);
  if (!res.ok) throw new Error("storage list failed: " + res.status);
  const data = await res.json();
  return { keys: data.keys || [], prefix, shared: true };
}

// Bulk read — pass a single prefix ("r:") or several comma-separated
// prefixes ("r:,pr:") to fetch multiple key ranges in one request.
async function bulk(prefix) {
  const res = await fetch(`${BASE}?bulk=${encodeURIComponent(prefix || "")}`);
  if (!res.ok) throw new Error("storage bulk failed: " + res.status);
  const data = await res.json();
  return data.entries || {};
}

async function del(key) {
  const res = await fetch(`${BASE}?key=${encodeURIComponent(key)}`, { method: "DELETE" });
  if (!res.ok) throw new Error("storage delete failed: " + res.status);
  return { key, deleted: true, shared: true };
}

const storage = { get, set, list, bulk, delete: del };
export default storage;
