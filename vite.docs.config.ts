import { fileURLToPath, URL } from "node:url";
import { readFileSync } from "node:fs";
import { defineConfig } from "vite";

const { version } = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as { version: string };
const siteUrl = new URL(process.env.DOCS_SITE_URL ?? "https://tcy-xsy.github.io/cyexcel/");
if (!siteUrl.pathname.endsWith("/")) siteUrl.pathname += "/";

export default defineConfig(({ command, isPreview }) => ({
  root: "docs-site",
  base: isPreview ? siteUrl.pathname : "./",
  define: {
    __CYEXCEL_VERSION__: JSON.stringify(version),
    __DOCS_SITE_URL__: JSON.stringify(siteUrl.href),
    __DOCS_BASE_PATH__: JSON.stringify(command === "build" || isPreview ? siteUrl.pathname : "/")
  },
  resolve: {
    alias: [
      {
        find: /^@cyexcel\//,
        replacement: `${fileURLToPath(new URL("./src", import.meta.url))}/`
      }
    ]
  },
  server: {
    host: "127.0.0.1",
    port: 4174
  },
  preview: {
    host: "127.0.0.1",
    port: 4174
  },
  build: {
    outDir: "../dist-docs",
    emptyOutDir: true
  }
}));
