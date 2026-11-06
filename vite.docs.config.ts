import { fileURLToPath, URL } from "node:url";
import { readFileSync } from "node:fs";
import { defineConfig } from "vite";

const { version } = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as { version: string };

export default defineConfig({
  root: "docs-site",
  base: "./",
  define: {
    __CYEXCEL_VERSION__: JSON.stringify(version)
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
});
