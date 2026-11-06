import "./styles/docs.css";
import { createDocPage } from "./document";
import { currentDocRoute, docHref } from "./i18n";
import { updateSeo } from "./seo";

const app = document.querySelector<HTMLElement>("#app");
if (app === null) throw new Error("missing #app");

let cleanup: (() => void) | undefined;

function render(): void {
  const location = currentDocRoute();
  const page = createDocPage(location.locale, location.path);
  if (window.location.hash.startsWith("#/")) {
    window.history.replaceState(null, "", docHref(location.locale, page.route));
  }
  cleanup?.();
  cleanup = undefined;
  document.documentElement.lang = location.locale === "en" ? "en" : "zh-CN";
  cleanup = page.shell.mount(app!);
  updateSeo(page.seo);
}

window.addEventListener("hashchange", () => {
  render();
  window.scrollTo({ top: 0 });
});

render();
