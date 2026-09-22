import "./styles/docs.css";
import { navGroups } from "./nav";
import { element } from "./components/doc";
import type { PageModule } from "./router";

const app = document.querySelector<HTMLElement>("#app");

if (app === null) {
  throw new Error("missing #app");
}

const shell = element("div", "docs-shell");
const header = element("header", "docs-header");
const brand = element("a", "brand");
brand.href = "#/";
brand.append(element("span", "brand-mark", "cy"), element("span", undefined, "cyexcel"));
const menuToggle = element("button", "menu-toggle", "菜单");
menuToggle.type = "button";
menuToggle.setAttribute("aria-label", "切换导航");
const headerMeta = element("div", "header-meta");
headerMeta.append(element("span", "header-version", "0.0.0"));
header.append(brand, headerMeta, menuToggle);

const body = element("div", "docs-body");
const sidebar = element("nav", "docs-sidebar");
const main = element("main", "docs-main");
const pageRoot = element("article", "docs-page");
main.append(pageRoot);
body.append(sidebar, main);

const footer = element("footer", "footer", "cyexcel · framework-agnostic spreadsheet component");
shell.append(header, body, footer);
app.append(shell);

let cleanup: (() => void) | undefined;

function readRoute(): string {
  const raw = window.location.hash.replace(/^#/, "");
  if (raw === "" || raw === "/") {
    return "/";
  }
  const route = raw.startsWith("/") ? raw : `/${raw}`;
  return route.replace(/\/+$/, "") || "/";
}

function findRoute(route: string): { label: string; module: PageModule } {
  for (const group of navGroups) {
    const link = group.links.find((entry) => entry.route === route);
    if (link !== undefined) {
      return link;
    }
  }
  return navGroups[0]?.links[0] ?? { label: "Overview", module: () => {} };
}

function renderSidebar(route: string): void {
  sidebar.replaceChildren();
  for (const group of navGroups) {
    const groupRoot = element("div", "sidebar-group");
    groupRoot.append(element("h2", "sidebar-title", group.title));
    for (const link of group.links) {
      const anchor = element("a", "sidebar-link", link.label);
      anchor.href = `#${link.route}`;
      if (link.route === route) {
        anchor.classList.add("active");
      }
      groupRoot.append(anchor);
    }
    sidebar.append(groupRoot);
  }
}

function render(): void {
  const route = readRoute();
  const page = findRoute(route);
  cleanup?.();
  cleanup = undefined;
  pageRoot.replaceChildren();
  renderSidebar(route);
  document.title = `${page.label} · cyexcel`;
  const result = page.module(pageRoot);
  if (typeof result === "function") {
    cleanup = result;
  }
  if (window.innerWidth <= 900) {
    sidebar.hidden = true;
    menuToggle.setAttribute("aria-expanded", "false");
  }
}

menuToggle.addEventListener("click", () => {
  const expanded = menuToggle.getAttribute("aria-expanded") === "true";
  sidebar.hidden = expanded;
  menuToggle.setAttribute("aria-expanded", expanded ? "false" : "true");
});

window.addEventListener("hashchange", () => {
  render();
  window.scrollTo({ top: 0 });
});

render();
