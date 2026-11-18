import { translate } from "../i18n";

/** Only apply to documentation fixtures, including list-option color keys. */
export function localizeDemoData<T>(value: T): T {
  if (typeof value === "string") return translate(value) as T;
  if (Array.isArray(value)) return value.map(item => localizeDemoData(item)) as T;
  if (value !== null && typeof value === "object" && Object.getPrototypeOf(value) === Object.prototype) {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [translate(key), localizeDemoData(item)])) as T;
  }
  return value;
}
