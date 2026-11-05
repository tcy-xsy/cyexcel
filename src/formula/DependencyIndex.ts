import type { CellRange } from "../core/types";

export interface DependencyRange extends CellRange { readonly key: string }
interface Node {
  readonly row: number;
  readonly crossing: DependencyRange[];
  readonly maxColumn: number[];
  readonly left: Node | undefined;
  readonly right: Node | undefined;
}

function build(ranges: DependencyRange[]): Node | undefined {
  if (!ranges.length) return undefined;
  const centers = ranges.map(r => (r.startRow + r.endRow) / 2).sort((a, b) => a - b);
  const row = centers[Math.floor(centers.length / 2)]!;
  const left: DependencyRange[] = [], right: DependencyRange[] = [], crossing: DependencyRange[] = [];
  for (const range of ranges) {
    if (range.endRow < row) left.push(range);
    else if (range.startRow > row) right.push(range);
    else crossing.push(range);
  }
  crossing.sort((a, b) => a.startColumn - b.startColumn);
  let max = -1;
  return { row, crossing, maxColumn: crossing.map(r => (max = Math.max(max, r.endColumn))), left: build(left), right: build(right) };
}

/** Range dependencies stay as rectangles; A1:A100000 does not create 100000 edges. */
export class DependencyIndex {
  readonly #ranges = new Map<string, readonly DependencyRange[]>();
  #root: Node | undefined;
  #dirty = false;

  set(key: string, ranges: readonly DependencyRange[]): void {
    this.#ranges.set(key, ranges);
    this.#dirty = true;
  }

  delete(key: string): void {
    if (this.#ranges.delete(key)) this.#dirty = true;
  }

  clear(): void {
    this.#ranges.clear();
    this.#root = undefined;
    this.#dirty = false;
  }

  query(row: number, column: number): Set<string> {
    if (this.#dirty) {
      this.#root = build([...this.#ranges.values()].flat());
      this.#dirty = false;
    }
    const result = new Set<string>();
    let node = this.#root;
    while (node) {
      let lo = 0, hi = node.crossing.length;
      while (lo < hi) {
        const mid = (lo + hi) >>> 1;
        if (node.crossing[mid]!.startColumn <= column) lo = mid + 1;
        else hi = mid;
      }
      for (let i = lo - 1; i >= 0 && node.maxColumn[i]! >= column; i -= 1) {
        const range = node.crossing[i]!;
        if (range.startRow <= row && range.endRow >= row && range.endColumn >= column) result.add(range.key);
      }
      node = row < node.row ? node.left : row > node.row ? node.right : undefined;
    }
    return result;
  }
}
