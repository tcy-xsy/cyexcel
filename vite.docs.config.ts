import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";

export default defineConfig({
  root: "docs-site",
  base: "./",
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
