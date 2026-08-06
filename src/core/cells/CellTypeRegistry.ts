import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellType, CellTypeDefinition } from "./types";

function baseDefaults<T>(): CellType<T> {
  return {
    format: (value) => (value === undefined || value === null ? "" : String(value))
  };
}

/**
 * Named, registrable CellType definitions with single-inheritance via
 * `extends`. `resolve()` flattens the chain once and caches the result by
 * name — repeated resolves (e.g. once per painted cell) never re-merge.
 */
export class CellTypeRegistry {
  readonly #definitions = new Map<string, CellTypeDefinition>();
  readonly #resolved = new Map<string, CellType>();

  register<T>(name: string, definition: CellTypeDefinition<T>): void {
    this.#definitions.set(name, definition as CellTypeDefinition);
    this.#resolved.delete(name);
  }

  resolve(name: string): CellType {
    const cached = this.#resolved.get(name);
    if (cached) {
      return cached;
    }

    const resolved = this.#resolveChain(name, new Set());
    this.#resolved.set(name, resolved);
    return resolved;
  }

  #resolveChain(name: string, visiting: Set<string>): CellType {
    const definition = this.#definitions.get(name);
    if (!definition) {
      throw new CyExcelError(
        CyExcelErrorCode.UnknownCellType,
        `Unknown cell type "${name}" — register it with registerColumnType() before referencing it`
      );
    }
    if (visiting.has(name)) {
      throw new CyExcelError(
        CyExcelErrorCode.UnknownCellType,
        `Cell type "${name}" has a circular extends chain`
      );
    }
    visiting.add(name);

    const parent = definition.extends
      ? this.#resolveChain(definition.extends, visiting)
      : baseDefaults();

    const { extends: _extends, ...ownFields } = definition;
    return { ...parent, ...ownFields };
  }
}
