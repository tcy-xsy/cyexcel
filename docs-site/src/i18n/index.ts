import { english } from "./en";

export type Locale = "zh" | "en";

let locale: Locale = "zh";

export function setLocale(value: Locale): void {
  locale = value;
}

export function getLocale(): Locale {
  return locale;
}

export function parseDocRoute(hash: string): { locale: Locale; path: string } {
  const route = `/${hash.replace(/^#\/?/, "")}`.replace(/\/+$/, "") || "/";
  if (route === "/en" || route.startsWith("/en/")) {
    return { locale: "en", path: route.slice(3) || "/" };
  }
  return { locale: "zh", path: route };
}

export function docHref(language: Locale, path: string): string {
  return `${__DOCS_BASE_PATH__}${language === "en" ? "en/" : ""}${path === "/" ? "" : `${path.slice(1)}/`}`;
}

export function currentDocRoute(): { locale: Locale; path: string } {
  if (window.location.hash.startsWith("#/")) return parseDocRoute(window.location.hash);
  return parseDocRoute(`#/${window.location.pathname.slice(__DOCS_BASE_PATH__.length)}`);
}

// Longest matches preserve complete sentences and translate concatenated prose
// or dynamic log fragments without replacing text inside an English result.
const sources = Object.keys(english).sort((a, b) => b.length - a.length);
const pattern = new RegExp(sources.map(source => source.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"), "gu");

export function translate(text: string): string {
  if (locale === "zh") return text;
  if (Object.hasOwn(english, text)) return english[text]!;
  return text.replace(pattern, source => english[source]!);
}
