import { readFile, mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createServer } from "vite";

const output = resolve("dist-docs");
const site = new URL(process.env.DOCS_SITE_URL ?? "https://tcy-xsy.github.io/cyexcel/");
if (!site.pathname.endsWith("/")) site.pathname += "/";
const template = await readFile(resolve(output, "index.html"), "utf8");

// Vite loads the shared TypeScript content in Node; demo mount callbacks never run.
const loader = await createServer({
  configFile: resolve("vite.docs.config.ts"),
  server: { middlewareMode: true },
  appType: "custom",
  define: { __DOCS_BASE_PATH__: JSON.stringify(site.pathname) }
});
let pages;
try {
  const { renderStaticPages } = await loader.ssrLoadModule("/src/static.ts");
  pages = renderStaticPages();
} finally {
  await loader.close();
}
if (pages.length === 0) throw new Error("No documentation pages generated");
if (new Set(pages.map(page => page.path)).size !== pages.length) throw new Error("Duplicate documentation paths");

for (const page of pages) {
  if (!page.path.startsWith(site.pathname)) throw new Error(`Invalid documentation path: ${page.path}`);
  const relative = page.path.slice(site.pathname.length);
  const depth = relative.split("/").filter(Boolean).length;
  const prefix = depth === 0 ? "./" : "../".repeat(depth);
  const html = template
    .replace('lang="zh-CN"', `lang="${page.language}"`)
    .replace(/<title>[^<]*<\/title>/, "")
    .replaceAll('="./assets/', `="${prefix}assets/`)
    .replace("</head>", () => `${page.head}\n</head>`)
    .replace('<div id="app"></div>', () => `<div id="app">${page.body}</div>`);
  const directory = resolve(output, relative);
  await mkdir(directory, { recursive: true });
  await writeFile(resolve(directory, "index.html"), html);
}

const escapeXml = text => text.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
const entries = pages.map(({ canonical, path }) => {
  const relative = path.slice(site.pathname.length).replace(/^en\//, "");
  const zh = escapeXml(new URL(relative, site).href);
  const en = escapeXml(new URL(`en/${relative}`, site).href);
  return `  <url><loc>${escapeXml(canonical)}</loc><xhtml:link rel="alternate" hreflang="zh-CN" href="${zh}"/><xhtml:link rel="alternate" hreflang="en" href="${en}"/><xhtml:link rel="alternate" hreflang="x-default" href="${zh}"/></url>`;
});
await writeFile(resolve(output, "sitemap.xml"), `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:xhtml="http://www.w3.org/1999/xhtml">\n${entries.join("\n")}\n</urlset>\n`);
await writeFile(resolve(output, "robots.txt"), `User-agent: *\nAllow: /\n\nSitemap: ${new URL("sitemap.xml", site).href}\n`);
console.log(`Generated ${pages.length} static documentation pages, sitemap.xml, and robots.txt without a browser.`);
