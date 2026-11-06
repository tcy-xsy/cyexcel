import { createDocPage } from "./document";
import { docHref } from "./i18n";
import { navGroups } from "./nav";
import { renderSeoHead } from "./seo";

export function renderStaticPages() {
  return navGroups.flatMap(group => group.links).flatMap(link => (["zh", "en"] as const).map(locale => {
    const { shell, seo } = createDocPage(locale, link.route);
    return {
      path: docHref(locale, link.route),
      canonical: seo.url,
      language: locale === "en" ? "en" : "zh-CN",
      head: renderSeoHead(seo),
      body: shell.toHtml()
    };
  }));
}
