declare module "fast-formula-parser" {
  interface CellRef { row: number; col: number; sheet?: string }
  interface RangeRef { from: { row: number; col: number }; to: { row: number; col: number }; sheet?: string }
  class _FormulaError extends Error {
    constructor(error: string);
    readonly error: string;
    readonly details?: unknown;
  }
  class _DependencyParser {
    parse(formula: string, position: CellRef): (CellRef | RangeRef)[];
  }
  class FormulaParser {
    constructor(options: { onCell: (ref: CellRef) => unknown; onRange: (ref: RangeRef) => unknown[][] });
    functions: Record<string, unknown>;
    parse(formula: string, position: CellRef): unknown;
    static readonly FormulaError: typeof _FormulaError;
    static readonly DepParser: typeof _DependencyParser;
  }
  export default FormulaParser;
}
