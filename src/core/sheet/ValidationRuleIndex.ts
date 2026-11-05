import type { CellRange } from "../types";
import type { DataValidationRule } from "./validations";

interface Entry {
  readonly range: CellRange;
  readonly rule: DataValidationRule;
  readonly order: number;
}

interface Node {
  readonly row: number;
  readonly crossing: readonly Entry[];
  readonly maxEndColumn: readonly number[];
  readonly left: Node | undefined;
  readonly right: Node | undefined;
}

function build(entries: Entry[]): Node | undefined {
  if (entries.length === 0) {
    return undefined;
  }
  const midpoints = entries.map(({ range }) => (range.startRow + range.endRow) / 2).sort((a, b) => a - b);
  const row = midpoints[Math.floor(midpoints.length / 2)] as number;
  const left: Entry[] = [];
  const right: Entry[] = [];
  const crossing: Entry[] = [];
  for (const entry of entries) {
    if (entry.range.endRow < row) {
      left.push(entry);
    } else if (entry.range.startRow > row) {
      right.push(entry);
    } else {
      crossing.push(entry);
    }
  }
  crossing.sort((a, b) => a.range.startColumn - b.range.startColumn);
  let largest = -1;
  const maxEndColumn = crossing.map(({ range }) => (largest = Math.max(largest, range.endColumn)));
  return { row, crossing, maxEndColumn, left: build(left), right: build(right) };
}

/** Immutable rule arrays share an index; ordinary edits do not rebuild it. */
const indexes = new WeakMap<readonly DataValidationRule[], Node>();
const eligibility = new WeakMap<readonly DataValidationRule[], boolean>();

export function canIndexValidationRules(rules: readonly DataValidationRule[]): boolean {
  if (!Object.isFrozen(rules)) return false;
  let eligible = eligibility.get(rules);
  if (eligible === undefined) {
    eligible = rules.every(rule => Object.isFrozen(rule) && Object.isFrozen(rule.ranges) &&
      rule.ranges.every(range => Object.isFrozen(range)));
    eligibility.set(rules, eligible);
  }
  return eligible;
}

export function findIndexedValidationRule(
  rules: readonly DataValidationRule[], row: number, column: number
): DataValidationRule | undefined {
  let node = indexes.get(rules);
  if (node === undefined) {
    node = build(rules.flatMap((rule, order) => rule.ranges.map((range) => ({ range, rule, order }))));
    if (node === undefined) {
      return undefined;
    }
    indexes.set(rules, node);
  }
  let best: Entry | undefined;
  while (node !== undefined) {
    let lo = 0;
    let hi = node.crossing.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if ((node.crossing[mid] as Entry).range.startColumn <= column) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }
    // Disjoint ranges visit at most one candidate. The prefix bound also
    // preserves first-rule precedence for callers with overlapping input.
    for (let i = lo - 1; i >= 0 && (node.maxEndColumn[i] as number) >= column; i -= 1) {
      const entry = node.crossing[i] as Entry;
      const range = entry.range;
      if (range.endColumn >= column && range.startRow <= row && range.endRow >= row &&
          (best === undefined || entry.order < best.order)) {
        best = entry;
      }
    }
    node = row < node.row ? node.left : row > node.row ? node.right : undefined;
  }
  return best?.rule;
}
