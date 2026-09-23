/**
 * Excel's "AutoFormat As You Type" rule: a cell becomes a hyperlink only
 * when its *entire* trimmed content is a URL or email address, not when one
 * appears inside a longer sentence. We follow the same rule — it's what
 * lets this run purely off the already-formatted display text, with no risk
 * of turning ordinary prose into a link by accident.
 */
const URL_PATTERN = /^(https?:\/\/|www\.)\S+$/i;
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/**
 * Returns the URL a cell's full display text should open to, or `undefined`
 * if it doesn't look like one. A bare "www." address gets an "https://"
 * prefix and an email gets a "mailto:" prefix — `text` itself is never
 * mutated, this is purely for navigation/paint-style decisions.
 */
export function detectLinkUrl(text: string): string | undefined {
  const trimmed = text.trim();
  if (trimmed === "") {
    return undefined;
  }
  if (URL_PATTERN.test(trimmed)) {
    return trimmed.toLowerCase().startsWith("www.") ? `https://${trimmed}` : trimmed;
  }
  if (EMAIL_PATTERN.test(trimmed)) {
    return `mailto:${trimmed}`;
  }
  return undefined;
}
