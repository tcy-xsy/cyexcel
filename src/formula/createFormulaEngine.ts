import FormulaParser from "fast-formula-parser";
import type { CellValue } from "../core/sheet/types";
import type { FormulaCell, FormulaEngine, FormulaEngineFactory, FormulaModelAccess } from "../core/formula/types";
import { DependencyIndex } from "./DependencyIndex";
import type { DependencyRange } from "./DependencyIndex";

export const SUPPORTED_FORMULA_FUNCTIONS = Object.freeze([
  "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTIF", "SUMIF", "IF", "IFERROR", "ROUND", "ABS", "CONCATENATE"
]);
const allowedFunctions = new Set<string>(SUPPORTED_FORMULA_FUNCTIONS);
const CURRENT_SHEET = "__cyexcel_current_sheet__";
const keyOf = (row: number, column: number): string => `${row}:${column}`;
const cellOf = (key: string): FormulaCell => { const [row, column] = key.split(":").map(Number); return { row: row!, column: column! }; };

interface FunctionArgument { readonly value: unknown; readonly isArray: boolean; readonly isCellRef?: boolean; readonly isRangeRef?: boolean }
function extremum(args: FunctionArgument[], minimum: boolean): number {
  let result: number | undefined;
  for (const arg of args) {
    const pending = [arg.value];
    while (pending.length) {
      const value = pending.pop();
      if (Array.isArray(value)) { for (const item of value) pending.push(item); continue; }
      if (value instanceof FormulaParser.FormulaError) throw value;
      let number: number;
      if (typeof value === "number") number = value;
      else if (!arg.isArray && !arg.isCellRef && !arg.isRangeRef && value !== null && value !== undefined) {
        number = Number(value);
        if (!Number.isFinite(number)) throw new FormulaParser.FormulaError("#VALUE!");
      } else continue;
      result = result === undefined ? number : minimum ? Math.min(result, number) : Math.max(result, number);
    }
  }
  return result ?? 0;
}
interface Node {
  readonly source: string;
  readonly dependencies: readonly DependencyRange[];
  dirty: boolean;
  value: CellValue | undefined;
  error: string | undefined;
}
export interface FormulaOptions {
  /** Maximum cells materialized by one range reference. Defaults to 100000. */
  readonly maxRangeCells?: number;
  /** Prevents excessively deep dependency chains from overflowing the JS stack. Defaults to 256. */
  readonly maxDepth?: number;
}

class Calculator implements FormulaEngine {
  readonly #model: FormulaModelAccess;
  readonly #maxRangeCells: number;
  readonly #maxDepth: number;
  readonly #nodes = new Map<string, Node>();
  readonly #dependents = new DependencyIndex();
  readonly #visiting = new Set<string>();
  #depParser: InstanceType<typeof FormulaParser.DepParser> | undefined;
  #parser: FormulaParser | undefined;

  constructor(model: FormulaModelAccess, options: FormulaOptions) {
    this.#model = model;
    this.#maxRangeCells = options.maxRangeCells ?? 100_000;
    this.#maxDepth = options.maxDepth ?? 256;
  }

  #validateRef(ref: { sheet?: string; row?: number; col?: number }): void {
    if ((ref.sheet !== undefined && ref.sheet !== CURRENT_SHEET) ||
        (ref.row !== undefined && (ref.row < 1 || ref.row > 1048576)) ||
        (ref.col !== undefined && (ref.col < 1 || ref.col > 16384))) {
      throw new FormulaParser.FormulaError("#REF!");
    }
  }

  #range(ref: { sheet?: string; from: { row: number; col: number }; to: { row: number; col: number } }): DependencyRange {
    this.#validateRef({ ...ref.from, ...(ref.sheet === undefined ? {} : { sheet: ref.sheet }) });
    this.#validateRef(ref.to);
    const rows = ref.to.row - ref.from.row + 1, columns = ref.to.col - ref.from.col + 1;
    if (rows <= 0 || columns <= 0 || rows * columns > this.#maxRangeCells) throw new FormulaParser.FormulaError("#NUM!");
    return { key: "", startRow: ref.from.row - 1, endRow: ref.to.row - 1, startColumn: ref.from.col - 1, endColumn: ref.to.col - 1 };
  }

  #read(row: number, column: number): unknown {
    const value = this.getValue(row, column);
    const key = keyOf(row, column);
    const error = this.#visiting.has(key) ? "#CYCLE!" : this.#nodes.get(key)?.error;
    return error === undefined ? value ?? null : new FormulaParser.FormulaError(error);
  }

  #getParser(): FormulaParser {
    if (!this.#parser) {
      this.#parser = new FormulaParser({
        onCell: ref => { this.#validateRef(ref); return this.#read(ref.row - 1, ref.col - 1); },
        onRange: ref => {
          const range = this.#range(ref);
          return Array.from({ length: range.endRow - range.startRow + 1 }, (_, row) =>
            Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, column) => this.#read(range.startRow + row, range.startColumn + column)));
        }
      });
      // These two functions are stubs in fast-formula-parser 1.0.19.
      this.#parser.functions.MIN = (...args: FunctionArgument[]) => extremum(args, true);
      this.#parser.functions.MAX = (...args: FunctionArgument[]) => extremum(args, false);
      // Keep the first release deterministic and synchronous. Dynamic references,
      // volatile functions and functions that perform network IO are not enabled.
      for (const name of Object.keys(this.#parser.functions)) {
        if (!allowedFunctions.has(name)) delete this.#parser.functions[name];
      }
    }
    return this.#parser;
  }

  getValue(row: number, column: number): CellValue | undefined {
    const record = this.#model.getRecord(row, column);
    if (record?.f === undefined) return record?.v;
    const key = keyOf(row, column);
    if (this.#visiting.has(key)) return "#CYCLE!";
    if (this.#visiting.size >= this.#maxDepth) throw new FormulaParser.FormulaError("#NUM!");
    let node = this.#nodes.get(key);
    if (node?.source === record.f && !node.dirty) return node.value;
    this.#visiting.add(key);
    try {
      const formula = record.f.startsWith("=") ? record.f.slice(1) : record.f;
      if (!formula.trim()) throw new FormulaParser.FormulaError("#ERROR!");
      const position = { row: row + 1, col: column + 1, sheet: CURRENT_SHEET };
      if (node?.source !== record.f) {
        this.#depParser ??= new FormulaParser.DepParser();
        const dependencies = this.#depParser.parse(formula, position).map(ref => {
          if ("from" in ref) return { ...this.#range(ref), key };
          this.#validateRef(ref);
          return { key, startRow: ref.row - 1, endRow: ref.row - 1, startColumn: ref.col - 1, endColumn: ref.col - 1 };
        });
        node = { source: record.f, dependencies, dirty: true, value: undefined, error: undefined };
        this.#nodes.set(key, node);
        this.#dependents.set(key, dependencies);
      }
      // Resolve referenced formulas before entering the shared parser. Its parse
      // state is mutable and cannot safely be entered recursively from onCell.
      for (const range of node.dependencies) {
        for (let r = range.startRow; r <= range.endRow; r += 1) {
          for (let c = range.startColumn; c <= range.endColumn; c += 1) {
            if (this.#model.getRecord(r, c)?.f !== undefined) this.getValue(r, c);
          }
        }
      }
      const parsed = this.#getParser().parse(formula, position);
      // A direct reference to an empty cell has numeric value zero in Excel.
      const result = parsed === null || parsed === undefined ? 0 : parsed;
      if (result instanceof FormulaParser.FormulaError) throw result;
      if (typeof result !== "number" && typeof result !== "string" && typeof result !== "boolean") throw new FormulaParser.FormulaError("#VALUE!");
      if (typeof result === "number" && !Number.isFinite(result)) throw new FormulaParser.FormulaError("#NUM!");
      node.value = result;
      node.error = undefined;
      node.dirty = false;
      return result;
    } catch (caught) {
      let error = caught;
      while (error instanceof FormulaParser.FormulaError && error.details instanceof FormulaParser.FormulaError) error = error.details;
      const code = error instanceof FormulaParser.FormulaError ? error.error : "#ERROR!";
      if (node === undefined || node.source !== record.f) {
        node = { source: record.f, dependencies: [], dirty: false, value: code, error: code };
        this.#nodes.set(key, node);
      }
      node.value = code;
      node.error = code;
      node.dirty = false;
      return code;
    } finally {
      this.#visiting.delete(key);
    }
  }

  getError(row: number, column: number): string | undefined {
    this.getValue(row, column);
    return this.#nodes.get(keyOf(row, column))?.error;
  }

  invalidate(row: number, column: number): readonly FormulaCell[] {
    const queue = [keyOf(row, column)], seen = new Set(queue);
    for (let i = 0; i < queue.length; i += 1) {
      const cell = cellOf(queue[i]!);
      for (const dependent of this.#dependents.query(cell.row, cell.column)) {
        if (!seen.has(dependent)) { seen.add(dependent); queue.push(dependent); }
      }
      const node = this.#nodes.get(queue[i]!);
      if (node) node.dirty = true;
    }
    // An edited formula will rebuild its dependency edges on its next read.
    const key = queue[0]!;
    this.#nodes.delete(key);
    this.#dependents.delete(key);
    return queue.slice(1).map(cellOf);
  }

  dispose(): void {
    this.#nodes.clear();
    this.#dependents.clear();
    this.#visiting.clear();
    this.#parser = undefined;
    this.#depParser = undefined;
  }
}

/** MIT-licensed parser adapter. The core only depends on FormulaEngineFactory. */
export function createFormulaEngine(options: FormulaOptions = {}): FormulaEngineFactory {
  for (const value of [options.maxRangeCells ?? 100_000, options.maxDepth ?? 256]) {
    if (!Number.isSafeInteger(value) || value < 1) throw new RangeError("Formula limits must be positive integers");
  }
  return model => new Calculator(model, options);
}
