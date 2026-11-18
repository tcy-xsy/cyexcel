import { enUS } from "./en-US.js";
import { zhCN } from "./zh-CN.js";
import type { Locale, Translator } from "./types.js";

export { enUS, zhCN };
export type { Locale, LocaleDefinition, MessageKey, Translator } from "./types.js";

/** Each instance captures its own dictionary; no mutable global locale. */
export function createTranslator(locale: Locale = "zh-CN"): Translator {
  if (typeof locale === "string" && locale !== "zh-CN" && locale !== "en-US") {
    throw new TypeError(`Unknown locale "${locale}". Pass a locale definition for custom languages.`);
  }
  const definition = typeof locale === "string"
    ? locale === "en-US" ? enUS : zhCN
    : locale;
  if (!definition || typeof definition.name !== "string" || !definition.name.trim() ||
      !definition.messages || typeof definition.messages !== "object" || Array.isArray(definition.messages)) {
    throw new TypeError("A locale definition requires a non-empty name and a messages dictionary.");
  }
  for (const value of Object.values(definition.messages)) {
    if (value !== undefined && typeof value !== "string") {
      throw new TypeError("Locale messages must be strings.");
    }
  }
  const messages = { ...enUS.messages, ...definition.messages };
  return (key, ...values) => {
    const message = messages[key] ?? enUS.messages[key];
    return message.replace(/\{(\d+)\}/g, (placeholder, index: string) => {
      const value = values[Number(index)];
      return value === undefined ? placeholder : String(value);
    });
  };
}

export const defaultTranslator: Translator = createTranslator();
