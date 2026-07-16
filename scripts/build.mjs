import { spawnSync } from "node:child_process";
import { resolve } from "node:path";

const isNetlify =
  process.env.NETLIFY === "true" || Boolean(process.env.NETLIFY_BUILD_BASE);

const command = isNetlify ? process.execPath : "bash";
const args = isNetlify
  ? [resolve("node_modules/next/dist/bin/next"), "build"]
  : ["scripts/build-verified.sh"];

console.log(
  isNetlify
    ? "Building the Next.js application for Netlify..."
    : "Building the Cloudflare Worker application for Sites...",
);

const result = spawnSync(command, args, {
  cwd: process.cwd(),
  env: process.env,
  stdio: "inherit",
});

if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}

process.exit(result.status ?? 1);
