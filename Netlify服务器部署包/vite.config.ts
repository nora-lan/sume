import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

function localSigningApi() {
  const dataPath = resolve(process.cwd(), ".local-data", "signing.json");
  return {
    name: "local-signing-api",
    configureServer(server: import("vite").ViteDevServer) {
      server.middlewares.use("/api/signing", async (request, response) => {
        response.setHeader("Content-Type", "application/json; charset=utf-8");
        response.setHeader("Cache-Control", "no-store");
        try {
          if (request.method === "GET") {
            try {
              response.end(JSON.stringify({ data: JSON.parse(await readFile(dataPath, "utf8")) }));
            } catch (error) {
              if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
              response.end(JSON.stringify({ data: null }));
            }
            return;
          }
          if (request.method !== "PUT") {
            response.statusCode = 405;
            response.end(JSON.stringify({ error: "Method not allowed" }));
            return;
          }
          const chunks: Buffer[] = [];
          for await (const chunk of request) chunks.push(Buffer.from(chunk));
          const body = Buffer.concat(chunks).toString("utf8");
          if (body.length > 5_000_000) {
            response.statusCode = 413;
            response.end(JSON.stringify({ error: "Payload too large" }));
            return;
          }
          const data = JSON.parse(body);
          await mkdir(resolve(process.cwd(), ".local-data"), { recursive: true });
          await writeFile(dataPath, JSON.stringify(data), "utf8");
          response.end(JSON.stringify({ ok: true }));
        } catch {
          response.statusCode = 400;
          response.end(JSON.stringify({ error: "Unable to process signing data" }));
        }
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), localSigningApi()],
});
