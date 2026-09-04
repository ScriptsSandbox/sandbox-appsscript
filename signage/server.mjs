import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { createDisplayDataService } from "./server/data-service.mjs";
import { loadEnv } from "./server/env.mjs";

const root = fileURLToPath(new URL(".", import.meta.url));
await loadEnv(join(root, ".env"));
const port = Number(process.env.PORT || 4173);
const dataService = createDisplayDataService({ root });
const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".gif": "image/gif",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

function sendJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

async function serveFile(requestPath, response) {
  const relativePath = requestPath === "/" ? "index.html" : requestPath.replace(/^\/+/, "");
  const filePath = resolve(root, normalize(relativePath));
  if (!filePath.startsWith(resolve(root) + "/")) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  try {
    const fileStat = await stat(filePath);
    if (!fileStat.isFile()) throw new Error("Not a file");
    response.writeHead(200, {
      "Content-Type": mimeTypes[extname(filePath)] || "application/octet-stream",
      "Cache-Control": extname(filePath) === ".html" ? "no-cache" : "public, max-age=300",
    });
    createReadStream(filePath).pipe(response);
  } catch {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
  }
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url, `http://${request.headers.host || "localhost"}`);
  if (url.pathname === "/api/display-data") {
    try {
      sendJson(response, 200, await dataService.getDisplayData());
    } catch (error) {
      sendJson(response, 503, { error: "Display data unavailable", detail: error.message });
    }
    return;
  }
  if (url.pathname === "/api/health") {
    sendJson(response, 200, dataService.getHealth());
    return;
  }
  await serveFile(url.pathname, response);
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Sandbox signage running at http://127.0.0.1:${port}`);
});
