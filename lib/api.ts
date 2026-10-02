export async function api<T = any>(action: string, payload: Record<string, unknown> = {}, adminToken?: string): Promise<T> {
  const res = await fetch("/api/backend", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action, payload, adminToken: adminToken || "" })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data?.error) {
    throw new Error(data?.message || `Request gagal (${res.status})`);
  }
  return data as T;
}
