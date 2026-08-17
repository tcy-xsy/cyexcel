import type { ColumnDefinition } from "./types";

/**
 * Captures caller-owned column metadata at an ownership boundary. Stateful
 * renderer/editor functions stay referenced; declarative metadata is copied
 * and frozen so later caller mutations cannot change live or history roots.
 */
export function snapshotColumnDefinitions(
  columns: readonly ColumnDefinition[]
): readonly ColumnDefinition[] {
  return Object.freeze(
    columns.map((column) =>
      Object.freeze({
        ...column,
        ...(typeof column.header === "object" && column.header !== null
          ? { header: Object.freeze({ ...column.header }) }
          : {}),
        ...(column.options !== undefined
          ? { options: Object.freeze(column.options.map((option) => Object.freeze({ ...option }))) }
          : {}),
        ...(column.clipboard !== undefined
          ? { clipboard: Object.freeze({ ...column.clipboard }) }
          : {})
      })
    )
  );
}
