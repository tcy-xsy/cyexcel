import { docHref, type Locale } from "./i18n";
import { escapeHtml } from "./components/content";

function meta(attribute: "name" | "property", key: string, content: string): void {
  let node = document.head.querySelector<HTMLMetaElement>(`meta[${attribute}="${key}"]`);
  if (node === null) {
    node = document.createElement("meta");
    node.setAttribute(attribute, key);
    document.head.append(node);
  }
  node.content = content;
}

function link(rel: string, href: string, language?: string): void {
  const selector = language === undefined ? `link[rel="${rel}"]` : `link[rel="${rel}"][hreflang="${language}"]`;
  let node = document.head.querySelector<HTMLLinkElement>(selector);
  if (node === null) {
    node = document.createElement("link");
    node.rel = rel;
    if (language !== undefined) node.hreflang = language;
    document.head.append(node);
  }
  node.href = href;
}

export function createSeo(locale: Locale, route: string, pageTitle: string, description: string) {
  const site = new URL(__DOCS_SITE_URL__);
  const urlFor = (language: Locale): string => new URL(docHref(language, route).slice(__DOCS_BASE_PATH__.length), site).href;
  const url = urlFor(locale);
  const title = route === "/"
    ? locale === "en" ? "cyexcel | JavaScript Spreadsheet Component" : "cyexcel | JavaScript 电子表格组件"
    : pageTitle;
  const metadata: Array<{ attribute: "name" | "property"; key: string; content: string }> = [
    { attribute: "name", key: "description", content: description },
    { attribute: "name", key: "robots", content: "index, follow" },
    { attribute: "property", key: "og:type", content: "website" },
    { attribute: "property", key: "og:site_name", content: "cyexcel" },
    { attribute: "property", key: "og:title", content: title },
    { attribute: "property", key: "og:description", content: description },
    { attribute: "property", key: "og:url", content: url },
    { attribute: "property", key: "og:locale", content: locale === "en" ? "en_US" : "zh_CN" },
    { attribute: "property", key: "og:locale:alternate", content: locale === "en" ? "zh_CN" : "en_US" },
    { attribute: "name", key: "twitter:card", content: "summary" },
    { attribute: "name", key: "twitter:title", content: title },
    { attribute: "name", key: "twitter:description", content: description }
  ];
  const alternates = [
    { language: "zh-CN", href: urlFor("zh") },
    { language: "en", href: urlFor("en") },
    { language: "x-default", href: urlFor("zh") }
  ];
  const schema = JSON.stringify({
    "@context": "https://schema.org",
    "@type": "TechArticle",
    headline: title,
    description,
    url,
    inLanguage: locale === "en" ? "en" : "zh-CN",
    isPartOf: { "@type": "WebSite", name: "cyexcel", url: site.href },
    about: {
      "@type": "SoftwareSourceCode",
      name: "cyexcel",
      description: "A framework-independent JavaScript spreadsheet component with Canvas rendering, formulas, data validation, and XLSX export.",
      codeRepository: "https://github.com/tcy-xsy/cyexcel",
      programmingLanguage: "TypeScript",
      license: "https://opensource.org/license/mit"
    }
  }).replace(/</g, "\\u003c");
  return { title, url, metadata, alternates, schema };
}

export function renderSeoHead(seo: ReturnType<typeof createSeo>): string {
  return [
    `<title>${escapeHtml(seo.title)}</title>`,
    ...seo.metadata.map(({ attribute, key, content }) => `<meta ${attribute}="${escapeHtml(key)}" content="${escapeHtml(content)}">`),
    `<link rel="canonical" href="${escapeHtml(seo.url)}">`,
    ...seo.alternates.map(({ language, href }) => `<link rel="alternate" hreflang="${language}" href="${escapeHtml(href)}">`),
    `<script type="application/ld+json">${seo.schema}</script>`
  ].join("\n");
}

export function updateSeo(seo: ReturnType<typeof createSeo>): void {
  document.title = seo.title;
  for (const { attribute, key, content } of seo.metadata) meta(attribute, key, content);
  link("canonical", seo.url);
  for (const { language, href } of seo.alternates) link("alternate", href, language);
  let schema = document.head.querySelector<HTMLScriptElement>('script[type="application/ld+json"]');
  if (schema === null) {
    schema = document.createElement("script");
    schema.type = "application/ld+json";
    document.head.append(schema);
  }
  schema.textContent = seo.schema;
}
