import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { serve } from "bun";
import html from "@/frontend/index.excel.html";
import { chatRoute } from "@/server/routes/chat";

// Load certificates
const certPath = join(homedir(), ".office-addin-dev-certs");
const cert = readFileSync(join(certPath, "localhost.crt"));
const key = readFileSync(join(certPath, "localhost.key"));

const server = serve({
  port: process.env.PORT,
  tls: { cert, key },
  development: { hmr: true, console: true },
  routes: {
    "/*": html,
    "/api/chat": chatRoute,
  },
});

console.log(
  `Excel Add-in dev server running at https://localhost:${server.port}`,
);
console.log(
  "Make sure to generate certificates: `bun run office-addin-dev-certs install`",
);
