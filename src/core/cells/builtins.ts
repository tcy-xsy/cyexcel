import type { CellTypeRegistry } from "./CellTypeRegistry";

const JSON_PREVIEW_LIMIT = 80;
const ARRAY_PREVIEW_COUNT = 3;

function formatText(value: unknown): string {
  return value === undefined || value === null ? "" : String(value);
}

function isValidDate(value: unknown): value is Date {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * Registers the built-in CellTypes listed in doc/architecture.md §8. Each
 * has a distinct, deterministic `format()` (no Intl locale formatting, to
 * keep output stable across test environments). `array`/`json` only render a
 * preview, per architecture.md's "数组和 JSON 默认只绘制预览内容".
 *
 * `compare` is defined by exactly the types whose natural order differs from
 * the order their values' string forms give, because that string order is
 * what `createDefaultSortCompare()` falls back to when a type defines no
 * `compare`. `text`/`select` are already in string order; `array`/`json`
 * have no meaningful total order at all.
 */
export function registerBuiltinCellTypes(registry: CellTypeRegistry): void {
  registry.register("text", {
    format: (value) => formatText(value)
  });

  registry.register("number", {
    format: (value) => (typeof value === "number" && Number.isFinite(value) ? String(value) : ""),
    parse: (text) => Number(text),
    compare: (left, right) => (left as number) - (right as number)
  });

  registry.register("boolean", {
    format: (value) => (value ? "true" : "false"),
    // Editor-agnostic: whatever text a text-based editor produces (default
    // or otherwise) has to round-trip correctly even if no dedicated
    // checkbox editor is ever wired up for this column — see doc/progress.md's
    // editor-pluggability patch.
    parse: (text) => text.trim().toLowerCase() === "true",
    serialize: (value) => (value ? "true" : "false"),
    // By truthiness, matching format()/parse(), not by the value's string
    // form: a "boolean" column may hold 1/0 as readily as true/false, and
    // sorting has to agree with what is drawn. String order would read
    // "0" < "1" < "false" < "true" and interleave the two groups.
    compare: (left, right) => Number(Boolean(left)) - Number(Boolean(right))
  });

  registry.register("date", {
    format: (value) => (isValidDate(value) ? value.toISOString().slice(0, 10) : ""),
    parse: (text) => new Date(text),
    serialize: (value) => (isValidDate(value) ? value.toISOString().slice(0, 10) : ""),
    // Chronological. `String(date)` starts with the weekday name, so without
    // this a date column would sort Fri < Mon < Sat — visibly wrong, and the
    // reason this exists at all (see the Step 22 sort review).
    //
    // Values that are not valid dates carry no instant to order by, so they
    // group together ahead of the real ones. That grouping is direction-
    // agnostic, so it flips under "desc"; genuinely empty cells (null /
    // undefined) never reach here — createDefaultSortCompare() keeps those
    // last in both directions before delegating.
    compare: (left, right) => {
      const leftTime = isValidDate(left) ? left.getTime() : Number.NaN;
      const rightTime = isValidDate(right) ? right.getTime() : Number.NaN;
      if (Number.isNaN(leftTime) || Number.isNaN(rightTime)) {
        if (Number.isNaN(leftTime) && Number.isNaN(rightTime)) {
          return 0;
        }
        return Number.isNaN(leftTime) ? -1 : 1;
      }
      return leftTime < rightTime ? -1 : leftTime > rightTime ? 1 : 0;
    }
  });

  registry.register("select", {
    // `options` is column/type metadata (see CellType.options), not an
    // editor choice — a column shows the matching label when configured,
    // and degrades to the raw value otherwise (no options configured, or a
    // value that isn't in the configured list).
    format: (value, context) => {
      const match = context.column.options?.find((option) => option.value === value);
      return match ? (match.label ?? match.value) : formatText(value);
    }
  });

  registry.register("array", {
    format: (value) => {
      if (!Array.isArray(value)) {
        return "";
      }
      const preview = value.slice(0, ARRAY_PREVIEW_COUNT).map((item) => formatText(item));
      const suffix = value.length > ARRAY_PREVIEW_COUNT ? ", ..." : "";
      return `[${value.length}] ${preview.join(", ")}${suffix}`;
    },
    // JSON array syntax, not comma-joined text — has to round-trip
    // arbitrary nested values (numbers, objects, ...), not just flat
    // strings. A TextAreaCellEditor with a validate() callback (see
    // doc/progress.md's editor-pluggability patch) is the primary defense
    // against committing malformed JSON; this catch is a second line of
    // defense for columns that never opted into one — same "degrade to an
    // empty-looking value, don't throw" precedent as "number"'s parse().
    parse: (text) => {
      try {
        const parsed = JSON.parse(text);
        return Array.isArray(parsed) ? parsed : [];
      } catch {
        return [];
      }
    },
    serialize: (value) => (Array.isArray(value) ? JSON.stringify(value) : "[]")
  });

  registry.register("json", {
    format: (value) => {
      const serialized = JSON.stringify(value);
      if (serialized === undefined) {
        return "";
      }
      return serialized.length > JSON_PREVIEW_LIMIT
        ? `${serialized.slice(0, JSON_PREVIEW_LIMIT - 3)}...`
        : serialized;
    },
    parse: (text) => {
      try {
        return JSON.parse(text);
      } catch {
        return undefined;
      }
    },
    serialize: (value) => {
      const serialized = JSON.stringify(value);
      return serialized === undefined ? "" : serialized;
    }
  });
}
