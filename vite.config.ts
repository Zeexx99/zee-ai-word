/// <reference types="vitest/config" />
import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import { appendFileSync, readFileSync } from "node:fs";
import { getHttpsServerOptions } from "office-addin-dev-certs";

type Middlewares = {
  use: (path: string, fn: (req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse) => void) => void;
};

/** Local only: the task pane posts its errors here and they are appended to client-log.txt. */
function clientLog(): Plugin {
  const file = fileURLToPath(new URL("./client-log.txt", import.meta.url));
  const attach = (middlewares: Middlewares) => {
    middlewares.use("/__zee_log", (req, res) => {
      let body = "";
      req.on("data", (c) => (body += c));
      req.on("end", () => {
        try {
          appendFileSync(file, `[${new Date().toISOString()}] ${body.slice(0, 8000)}\n`);
        } catch {
          /* ignore */
        }
        res.statusCode = 204;
        res.end();
      });
    });
  };
  return {
    name: "zee-client-log",
    configureServer(server) {
      attach(server.middlewares);
    },
    configurePreviewServer(server) {
      attach(server.middlewares);
    },
  };
}

// Office requires the task pane to be served over HTTPS.
// office-addin-dev-certs creates and trusts a localhost certificate (setup.bat / `npm run certs`).
// Port 3200: 3000 is often taken by other dev servers and 3100 is used by the Zee AI Excel add-in,
// so both add-ins can run side by side.
const PORT = 3200;

export default defineConfig(async ({ command }) => {
  // No certificates for `vite build` (CI / GitHub Pages) or for Vitest (which uses "serve" internally).
  const https = command === "serve" && !process.env.VITEST ? await getHttpsServerOptions() : undefined;
  const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as { version: string };
  return {
    // Relative base: the same build works on https://localhost:3200/ and https://<user>.github.io/zee-ai-word/
    base: "./",
    plugins: [react(), clientLog()],
    // Shown in Settings → About, so you can tell which build Word actually loaded.
    define: {
      __ZEE_VERSION__: JSON.stringify(pkg.version),
      __ZEE_BUILD__: JSON.stringify(new Date().toISOString()),
    },
    server: { port: PORT, strictPort: true, https },
    preview: { port: PORT, strictPort: true, https },
    test: { environment: "node", include: ["src/__tests__/**/*.test.ts"] },
    build: {
      outDir: "dist",
      // No source maps in the shipped build: they were ~14 MB (mostly the icon set) and made every
      // GitHub Pages deploy and first load heavier. Set ZEE_SOURCEMAP=1 to get them for debugging.
      sourcemap: process.env.ZEE_SOURCEMAP === "1",
      chunkSizeWarningLimit: 900,
      rollupOptions: {
        input: { taskpane: fileURLToPath(new URL("./taskpane.html", import.meta.url)) },
        output: {
          // Vendor libraries in separate, cacheable files. App code is NOT lazy-loaded: a lazy chunk
          // breaks when Office still holds a cached page from an older deploy.
          manualChunks: {
            fluent: ["@fluentui/react-components"],
            icons: ["@fluentui/react-icons"],
            markdown: ["react-markdown", "remark-gfm", "rehype-highlight", "lowlight"],
          },
        },
      },
    },
  };
});
