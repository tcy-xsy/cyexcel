import { element } from "./components/doc";
import { navGroups } from "./nav";
import { docHref, setLocale, translate, type Locale } from "./i18n";
import { createSeo } from "./seo";

export function createDocPage(locale: Locale, path: string) {
  setLocale(locale);
  const page = navGroups.flatMap(group => group.links).find(link => link.route === path) ?? navGroups[0]!.links[0]!;
  const route = page.route;
  const shell = element("div", "docs-shell");
  const header = element("header", "docs-header");
  const brand = element("a", "brand");
  brand.attributes.href = docHref(locale, "/");
  brand.append(element("span", "brand-mark", "cy"), element("span", undefined, "cyexcel"));
  const menu = element("button", "menu-toggle", "菜单");
  menu.attributes.type = "button";
  menu.attributes["aria-label"] = translate("切换导航");
  const headerMeta = element("div", "header-meta");
  const language = element("a", "language-toggle");
  language.setText(locale === "en" ? "中文" : "English");
  language.attributes.href = docHref(locale === "en" ? "zh" : "en", route);
  language.attributes.hreflang = locale === "en" ? "zh-CN" : "en";
  headerMeta.append(language, element("span", "header-version", __CYEXCEL_VERSION__));
  header.append(brand, headerMeta, menu);

  const body = element("div", "docs-body");
  const sidebar = element("nav", "docs-sidebar");
  for (const group of navGroups) {
    const groupRoot = element("div", "sidebar-group");
    groupRoot.append(element("h2", "sidebar-title", group.title));
    for (const link of group.links) {
      const anchor = element("a", "sidebar-link", link.label);
      anchor.attributes.href = docHref(locale, link.route);
      if (link.route === route) {
        anchor.addClass("active");
        anchor.attributes["aria-current"] = "page";
      }
      groupRoot.append(anchor);
    }
    sidebar.append(groupRoot);
  }
  menu.onClick = () => {
    const expanded = menu.element.getAttribute("aria-expanded") === "true";
    sidebar.element.hidden = expanded;
    menu.element.setAttribute("aria-expanded", expanded ? "false" : "true");
  };
  sidebar.onMount = () => {
    if (window.innerWidth <= 900) {
      sidebar.element.hidden = true;
      menu.element.setAttribute("aria-expanded", "false");
    }
  };

  const main = element("main", "docs-main");
  const article = element("article", "docs-page");
  page.module(article);
  main.append(article);
  body.append(sidebar, main);
  shell.append(header, body, element("footer", "footer", "cyexcel · framework-agnostic spreadsheet component"));
  const description = article.findByClass("page-description")?.textContent;
  if (description === undefined) throw new Error(`Missing description: ${route}`);
  const seo = createSeo(locale, route, `${translate(page.label)} · cyexcel`, description);
  return { shell, seo, route };
}
