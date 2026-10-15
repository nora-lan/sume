import { getStore } from "@netlify/blobs";

const STORE_NAME = "contract-signing";
const RECORD_KEY = "hanheng-channel-account-change-2026";
function json(data: unknown, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" },
  });
}

export default async function signing(request: Request) {
  if (request.method !== "GET" && request.method !== "PUT") return json({ error: "Method not allowed" }, 405);
  const origin = request.headers.get("origin");
  if (origin && origin !== new URL(request.url).origin) return json({ error: "Cross-origin requests are not allowed" }, 403);

  try {
    const store = getStore(STORE_NAME, { region: "ap-southeast-1" });
    if (request.method === "GET") {
      const data = await store.get(RECORD_KEY, { type: "json" });
      return json({ data: data ?? null });
    }

    const contentLength = Number(request.headers.get("content-length") || 0);
    if (contentLength > 5_000_000) return json({ error: "Payload too large" }, 413);
    const input = await request.json() as { signatures?: unknown; positions?: unknown; history?: unknown };
    if (!Array.isArray(input.signatures) || input.signatures.length !== 2 ||
        !input.signatures.every((item) => item === null || (typeof item === "string" && item.startsWith("data:image/png;base64,") && item.length < 700_000)) ||
        !Array.isArray(input.positions) || input.positions.length !== 2 ||
        !input.positions.every((item) => item && typeof item.x === "number" && Number.isFinite(item.x) && item.x >= 0 && item.x <= 1 &&
          typeof item.y === "number" && Number.isFinite(item.y) && item.y >= 0 && item.y <= 1) ||
        (input.history !== undefined && (!Array.isArray(input.history) || input.history.length > 100 ||
          !input.history.every((item) => item && typeof item.id === "string" && Number.isInteger(item.signerIndex) &&
            item.signerIndex >= 0 && item.signerIndex < 2 && typeof item.signature === "string" &&
            item.signature.startsWith("data:image/png;base64,") && item.signature.length < 700_000 &&
            typeof item.createdAt === "string")))) {
      return json({ error: "Invalid signing data" }, 400);
    }

    await store.setJSON(RECORD_KEY, {
      signatures: input.signatures,
      positions: input.positions,
      history: input.history ?? [],
    });
    return json({ ok: true });
  } catch (error) {
    console.error("Signing data persistence failed", error);
    return json({ error: "Signing data storage is unavailable" }, 500);
  }
}
