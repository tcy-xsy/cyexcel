import { assertValidRowOrder } from "./rowOrder";

const DENSE_LEAF_SIZE = 512;

interface NodeBase {
  readonly rowCount: number;
  readonly totalHeight: number;
  readonly depth: number;
}

interface UniformLeaf extends NodeBase {
  readonly kind: "uniform";
  readonly height: number;
}

interface DenseLeaf extends NodeBase {
  readonly kind: "dense";
  readonly values: Float64Array;
  readonly offsets: Float64Array;
}

interface Branch extends NodeBase {
  readonly kind: "branch";
  readonly left: Node;
  readonly right: Node;
}

type Leaf = UniformLeaf | DenseLeaf;
type Node = Leaf | Branch;

function assertCount(name: string, value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative safe integer: ${value}`);
  }
}

function assertHeight(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

function uniformLeaf(rowCount: number, height: number): UniformLeaf | null {
  if (rowCount === 0) {
    return null;
  }
  if (!Number.isSafeInteger(rowCount)) {
    throw new RangeError("row metrics rowCount must remain a safe integer");
  }
  const totalHeight = rowCount * height;
  if (!Number.isFinite(totalHeight)) {
    throw new RangeError("row metrics totalHeight must remain finite");
  }
  return Object.freeze({
    kind: "uniform" as const,
    rowCount,
    totalHeight,
    depth: 1,
    height
  });
}

function denseLeaf(values: Float64Array): DenseLeaf | null {
  if (values.length === 0) {
    return null;
  }
  const offsets = new Float64Array(values.length + 1);
  for (let index = 0; index < values.length; index += 1) {
    offsets[index + 1] = (offsets[index] as number) + (values[index] as number);
  }
  const totalHeight = offsets[values.length] as number;
  if (!Number.isFinite(totalHeight)) {
    throw new RangeError("row metrics totalHeight must remain finite");
  }
  return Object.freeze({
    kind: "dense" as const,
    rowCount: values.length,
    totalHeight,
    depth: 1,
    values,
    offsets
  });
}

function branch(left: Node, right: Node): Branch {
  const rowCount = left.rowCount + right.rowCount;
  const totalHeight = left.totalHeight + right.totalHeight;
  if (!Number.isSafeInteger(rowCount)) {
    throw new RangeError("row metrics rowCount must remain a safe integer");
  }
  if (!Number.isFinite(totalHeight)) {
    throw new RangeError("row metrics totalHeight must remain finite");
  }
  return Object.freeze({
    kind: "branch" as const,
    rowCount,
    totalHeight,
    depth: Math.max(left.depth, right.depth) + 1,
    left,
    right
  });
}

function asDenseValues(leaf: Leaf): Float64Array {
  if (leaf.kind === "dense") {
    return leaf.values;
  }
  const values = new Float64Array(leaf.rowCount);
  values.fill(leaf.height);
  return values;
}

function mergeLeaves(left: Leaf, right: Leaf): Leaf | null {
  if (left.kind === "uniform" && right.kind === "uniform" && left.height === right.height) {
    return uniformLeaf(left.rowCount + right.rowCount, left.height);
  }
  if (left.rowCount + right.rowCount > DENSE_LEAF_SIZE) {
    return null;
  }
  const values = new Float64Array(left.rowCount + right.rowCount);
  values.set(asDenseValues(left));
  values.set(asDenseValues(right), left.rowCount);
  return denseLeaf(values);
}

function collapse(node: Node): Node {
  if (node.kind !== "branch") {
    return node;
  }
  if (node.left.kind !== "branch" && node.right.kind !== "branch") {
    return mergeLeaves(node.left, node.right) ?? node;
  }
  return node;
}

function join(left: Node | null, right: Node | null): Node | null {
  if (left === null) {
    return right;
  }
  if (right === null) {
    return left;
  }
  if (left.kind !== "branch" && right.kind !== "branch") {
    const merged = mergeLeaves(left, right);
    if (merged !== null) {
      return merged;
    }
  }
  if (left.depth > right.depth + 1 && left.kind === "branch") {
    return rebalance(branch(left.left, join(left.right, right) as Node));
  }
  if (right.depth > left.depth + 1 && right.kind === "branch") {
    return rebalance(branch(join(left, right.left) as Node, right.right));
  }
  return branch(left, right);
}

function rebalance(node: Branch): Node {
  const balance = node.left.depth - node.right.depth;
  if (balance > 1 && node.left.kind === "branch") {
    const left = node.left;
    if (left.right.depth > left.left.depth && left.right.kind === "branch") {
      const pivot = left.right;
      return collapse(branch(
        collapse(branch(left.left, pivot.left)),
        collapse(branch(pivot.right, node.right))
      ));
    }
    return collapse(branch(left.left, collapse(branch(left.right, node.right))));
  }
  if (balance < -1 && node.right.kind === "branch") {
    const right = node.right;
    if (right.left.depth > right.right.depth && right.left.kind === "branch") {
      const pivot = right.left;
      return collapse(branch(
        collapse(branch(node.left, pivot.left)),
        collapse(branch(pivot.right, right.right))
      ));
    }
    return collapse(branch(collapse(branch(node.left, right.left)), right.right));
  }
  return collapse(node);
}

function splitLeaf(leaf: Leaf, index: number): readonly [Node | null, Node | null] {
  if (index === 0) {
    return [null, leaf];
  }
  if (index === leaf.rowCount) {
    return [leaf, null];
  }
  if (leaf.kind === "uniform") {
    return [uniformLeaf(index, leaf.height), uniformLeaf(leaf.rowCount - index, leaf.height)];
  }
  return [
    denseLeaf(leaf.values.slice(0, index)),
    denseLeaf(leaf.values.slice(index))
  ];
}

function split(node: Node | null, index: number): readonly [Node | null, Node | null] {
  if (node === null) {
    return [null, null];
  }
  if (index === 0) {
    return [null, node];
  }
  if (index === node.rowCount) {
    return [node, null];
  }
  if (node.kind !== "branch") {
    return splitLeaf(node, index);
  }
  if (index < node.left.rowCount) {
    const [before, after] = split(node.left, index);
    return [before, join(after, node.right)];
  }
  if (index > node.left.rowCount) {
    const [before, after] = split(node.right, index - node.left.rowCount);
    return [join(node.left, before), after];
  }
  return [node.left, node.right];
}

function heightAt(node: Node, row: number): number {
  let current = node;
  let index = row;
  while (current.kind === "branch") {
    if (index < current.left.rowCount) {
      current = current.left;
    } else {
      index -= current.left.rowCount;
      current = current.right;
    }
  }
  return current.kind === "uniform" ? current.height : (current.values[index] as number);
}

function offsetAt(node: Node | null, row: number): number {
  let current = node;
  let index = row;
  let total = 0;
  while (current !== null && current.kind === "branch") {
    if (index <= current.left.rowCount) {
      current = current.left;
    } else {
      total += current.left.totalHeight;
      index -= current.left.rowCount;
      current = current.right;
    }
  }
  if (current === null || index === 0) {
    return total;
  }
  if (current.kind === "uniform") {
    return total + index * current.height;
  }
  return total + (current.offsets[index] as number);
}

function rowAtOffset(node: Node | null, offset: number): number {
  if (node === null || Number.isNaN(offset) || offset <= 0) {
    return 0;
  }
  if (offset >= node.totalHeight) {
    return node.rowCount;
  }
  let current = node;
  let top = 0;
  let row = 0;
  while (current.kind === "branch") {
    const boundary = top + current.left.totalHeight;
    if (offset < boundary) {
      current = current.left;
    } else {
      top = boundary;
      row += current.left.rowCount;
      current = current.right;
    }
  }
  let low = 0;
  let high = current.rowCount;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    const localOffset = current.kind === "uniform"
      ? middle * current.height
      : (current.offsets[middle] as number);
    if (top + localOffset <= offset) {
      low = middle;
    } else {
      high = middle - 1;
    }
  }
  return row + low;
}

function setHeight(node: Node, row: number, value: number): Node {
  if (node.kind === "branch") {
    if (row < node.left.rowCount) {
      return join(setHeight(node.left, row, value), node.right) as Node;
    }
    return join(node.left, setHeight(node.right, row - node.left.rowCount, value)) as Node;
  }
  if (node.kind === "dense") {
    if (node.values[row] === value) {
      return node;
    }
    const values = node.values.slice();
    values[row] = value;
    return denseLeaf(values) as Node;
  }
  if (node.height === value) {
    return node;
  }

  const pageStart = Math.floor(row / DENSE_LEAF_SIZE) * DENSE_LEAF_SIZE;
  const pageCount = Math.min(DENSE_LEAF_SIZE, node.rowCount - pageStart);
  const values = new Float64Array(pageCount);
  values.fill(node.height);
  values[row - pageStart] = value;
  return join(
    join(uniformLeaf(pageStart, node.height), denseLeaf(values)),
    uniformLeaf(node.rowCount - pageStart - pageCount, node.height)
  ) as Node;
}

function appendDense(node: Node | null, target: Float64Array, start: number): number {
  if (node === null) {
    return start;
  }
  if (node.kind === "branch") {
    return appendDense(node.right, target, appendDense(node.left, target, start));
  }
  if (node.kind === "uniform") {
    target.fill(node.height, start, start + node.rowCount);
  } else {
    target.set(node.values, start);
  }
  return start + node.rowCount;
}

function buildBalanced(nodes: readonly Node[], start = 0, end = nodes.length): Node | null {
  if (start >= end) {
    return null;
  }
  if (end - start === 1) {
    return nodes[start] as Node;
  }
  const middle = Math.floor((start + end) / 2);
  return branch(buildBalanced(nodes, start, middle) as Node, buildBalanced(nodes, middle, end) as Node);
}

/**
 * Immutable row geometry backed by a persistent rope. Uniform spans stay
 * compressed; edits copy one 512-row leaf plus the path to it.
 */
export class RowMetrics {
  readonly #root: Node | null;
  readonly rowCount: number;
  readonly totalHeight: number;

  private constructor(root: Node | null) {
    this.#root = root;
    this.rowCount = root?.rowCount ?? 0;
    this.totalHeight = root?.totalHeight ?? 0;
    Object.freeze(this);
  }

  static uniform(rowCount: number, height: number): RowMetrics {
    assertCount("rowCount", rowCount);
    assertHeight("height", height);
    return new RowMetrics(uniformLeaf(rowCount, height));
  }

  static create(
    rowCount: number,
    defaultHeight: number,
    overrides: ReadonlyMap<number, number> = new Map()
  ): RowMetrics {
    const metrics = RowMetrics.uniform(rowCount, defaultHeight);
    return metrics.withHeights(overrides);
  }

  static fromDenseValues(source: ArrayLike<number>): RowMetrics {
    assertCount("values.length", source.length);
    const leaves: Node[] = [];
    for (let start = 0; start < source.length; start += DENSE_LEAF_SIZE) {
      const count = Math.min(DENSE_LEAF_SIZE, source.length - start);
      const values = new Float64Array(count);
      let uniform = true;
      let first = 0;
      for (let offset = 0; offset < count; offset += 1) {
        const value = source[start + offset] as number;
        assertHeight(`values[${start + offset}]`, value);
        values[offset] = value;
        if (offset === 0) {
          first = value;
        } else if (value !== first) {
          uniform = false;
        }
      }
      const leaf = uniform ? uniformLeaf(count, first) : denseLeaf(values);
      if (leaf !== null) {
        const previous = leaves.at(-1);
        if (previous?.kind === "uniform" && leaf.kind === "uniform" && previous.height === leaf.height) {
          leaves[leaves.length - 1] = uniformLeaf(previous.rowCount + leaf.rowCount, leaf.height) as Node;
        } else {
          leaves.push(leaf);
        }
      }
    }
    return new RowMetrics(buildBalanced(leaves));
  }

  heightAt(row: number): number {
    if (!Number.isSafeInteger(row) || row < 0 || row >= this.rowCount || this.#root === null) {
      return 0;
    }
    return heightAt(this.#root, row);
  }

  offsetAt(row: number): number {
    if (!Number.isSafeInteger(row) || row < 0 || row > this.rowCount) {
      return 0;
    }
    return offsetAt(this.#root, row);
  }

  rowAtOffset(offset: number): number {
    return rowAtOffset(this.#root, offset);
  }

  withHeight(row: number, height: number): RowMetrics {
    if (!Number.isSafeInteger(row) || row < 0 || row >= this.rowCount || this.#root === null) {
      throw new RangeError(`row is outside row metrics: ${row}`);
    }
    assertHeight("height", height);
    const root = setHeight(this.#root, row, height);
    return root === this.#root ? this : new RowMetrics(root);
  }

  withHeights(updates: ReadonlyMap<number, number>): RowMetrics {
    const entries = [...updates];
    for (const [row, height] of entries) {
      if (!Number.isSafeInteger(row) || row < 0 || row >= this.rowCount) {
        throw new RangeError(`row is outside row metrics: ${row}`);
      }
      assertHeight("height", height);
    }
    let root = this.#root;
    if (root === null) {
      return this;
    }
    for (const [row, height] of entries) {
      root = setHeight(root, row, height);
    }
    return root === this.#root ? this : new RowMetrics(root);
  }

  insert(start: number, fragment: RowMetrics): RowMetrics {
    if (!Number.isSafeInteger(start) || start < 0 || start > this.rowCount) {
      throw new RangeError(`insert start is outside row metrics: ${start}`);
    }
    if (fragment.rowCount === 0) {
      return this;
    }
    const [before, after] = split(this.#root, start);
    return new RowMetrics(join(join(before, fragment.#root), after));
  }

  delete(start: number, count: number): {
    readonly metrics: RowMetrics;
    readonly deleted: RowMetricsFragment;
  } {
    this.#assertRange(start, count);
    const [before, tail] = split(this.#root, start);
    const [deleted, after] = split(tail, count);
    return Object.freeze({
      metrics: new RowMetrics(join(before, after)),
      deleted: new RowMetrics(deleted)
    });
  }

  move(start: number, count: number, target: number): RowMetrics {
    this.#assertRange(start, count);
    if (!Number.isSafeInteger(target) || target < 0 || target > this.rowCount) {
      throw new RangeError(`move target is outside row metrics: ${target}`);
    }
    if (target > start && target < start + count) {
      throw new RangeError("move target cannot be inside the moved block");
    }
    const [before, tail] = split(this.#root, start);
    const [moved, after] = split(tail, count);
    const remaining = join(before, after);
    const adjustedTarget = target > start ? target - count : target;
    const [targetBefore, targetAfter] = split(remaining, adjustedTarget);
    return new RowMetrics(join(join(targetBefore, moved), targetAfter));
  }

  slice(start: number, count: number): RowMetricsFragment {
    this.#assertRange(start, count);
    const [, tail] = split(this.#root, start);
    const [fragment] = split(tail, count);
    return new RowMetrics(fragment);
  }

  exportDenseValues(): Float64Array {
    const values = new Float64Array(this.rowCount);
    appendDense(this.#root, values, 0);
    return values;
  }

  /**
   * Applies a row permutation (`order[newPosition] = oldPosition`).
   *
   * Short-circuits when the whole tree is a single uniform leaf: every row
   * has the same height, so any permutation of them is the same tree. That
   * is the million-row default case, and without this check reordering would
   * materialise an 8MB dense array to produce an identical result.
   */
  reorder(order: Int32Array): RowMetrics {
    assertValidRowOrder(order, this.rowCount);
    if (this.#root === null || this.#root.kind === "uniform") {
      return this;
    }
    const current = this.exportDenseValues();
    const next = new Float64Array(order.length);
    for (let index = 0; index < order.length; index += 1) {
      next[index] = current[order[index] as number] as number;
    }
    return RowMetrics.fromDenseValues(next);
  }

  /**
   * Zeroes the given rows' height — the filtering mechanism (ADR 0003 §3.2):
   * a filtered row contributes 0 to every derived offset/total, and a
   * zero-height run collapses into one `totalHeight === 0` node via the
   * `fromDenseValues` adjacent-uniform-leaf merge, so `rowAtOffset`/
   * `offsetAt` skip a whole hidden run in O(log n) with no other changes.
   *
   * Pure: `this` is returned untouched when the set is empty or every named
   * row is already 0.
   */
  withFilteredRows(filteredRows: ReadonlySet<number>): RowMetrics {
    if (filteredRows.size === 0) {
      return this;
    }
    for (const row of filteredRows) {
      if (!Number.isSafeInteger(row) || row < 0 || row >= this.rowCount) {
        throw new RangeError(`row is outside row metrics: ${row}`);
      }
    }
    const values = this.exportDenseValues();
    let changed = false;
    for (const row of filteredRows) {
      if (values[row] !== 0) {
        values[row] = 0;
        changed = true;
      }
    }
    return changed ? RowMetrics.fromDenseValues(values) : this;
  }

  #assertRange(start: number, count: number): void {
    assertCount("start", start);
    assertCount("count", count);
    if (count === 0 || start + count > this.rowCount) {
      throw new RangeError("row metrics range is empty or outside the axis");
    }
  }
}

export type RowMetricsFragment = RowMetrics;
