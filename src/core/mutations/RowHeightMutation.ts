import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { LayoutEngine, LayoutRoot, RowHeightProvenance } from "../layout";
import type { Patch } from "../types";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";

export interface RowHeightMutationUpdate {
  readonly row: number;
  readonly height: number;
  readonly provenance: RowHeightProvenance;
}

export type RowHeightMutationPolicy = "overwrite" | "preserve-manual";

function invalidUpdate(message: string): CyExcelError {
  return new CyExcelError(CyExcelErrorCode.InvalidOptions, message);
}

function snapshotUpdates(updates: readonly RowHeightMutationUpdate[]): readonly RowHeightMutationUpdate[] {
  const rows = new Set<number>();
  return Object.freeze(updates.map((update, index) => {
    if (!Number.isSafeInteger(update.row) || update.row < 0) {
      throw invalidUpdate(`Invalid row-height mutation row at index ${index}: ${update.row}`);
    }
    if (!Number.isFinite(update.height) || update.height <= 0) {
      throw invalidUpdate(`Invalid row-height mutation height at index ${index}: ${update.height}`);
    }
    if (
      update.provenance !== "default" &&
      update.provenance !== "auto" &&
      update.provenance !== "manual"
    ) {
      throw invalidUpdate(
        `Invalid row-height mutation provenance at index ${index}: ${String(update.provenance)}`
      );
    }
    if (rows.has(update.row)) {
      throw invalidUpdate(`Duplicate row-height mutation update for row ${update.row}`);
    }
    rows.add(update.row);
    return Object.freeze({
      row: update.row,
      height: update.height,
      provenance: update.provenance
    });
  }));
}

function patchesFor(
  updates: readonly RowHeightMutationUpdate[],
  root: LayoutRoot
): readonly Extract<Patch, { type: "row-layout" }>[] {
  if (updates.length === 0) {
    return Object.freeze([]);
  }
  const mutableRows: [number, ...number[]] = [updates[0]!.row];
  let startRow = mutableRows[0];
  for (let index = 1; index < updates.length; index += 1) {
    const row = updates[index]!.row;
    mutableRows.push(row);
    startRow = Math.min(startRow, row);
  }
  // Preserve the frozen Command's update order in the semantic fact.
  const rows: readonly [number, ...number[]] = Object.freeze(mutableRows);
  return Object.freeze([Object.freeze({
    type: "row-layout",
    rows,
    range: Object.freeze({
      startRow,
      startColumn: 0,
      endRow: root.rowCount - 1,
      endColumn: Math.max(0, root.columnCount - 1)
    })
  })]);
}

/** Atomic row-height batch whose prepare composes through the transaction-local Layout root. */
export class RowHeightMutation implements Mutation {
  readonly type = "row-heights";
  readonly updates: readonly RowHeightMutationUpdate[];
  readonly policy: RowHeightMutationPolicy;
  readonly changes: readonly Patch[];
  readonly #layout: LayoutEngine;

  constructor(
    layout: LayoutEngine,
    updates: readonly RowHeightMutationUpdate[],
    policy: RowHeightMutationPolicy = "overwrite",
    changes?: readonly Patch[]
  ) {
    if (policy !== "overwrite" && policy !== "preserve-manual") {
      throw invalidUpdate(`Invalid row-height mutation policy: ${String(policy)}`);
    }
    this.#layout = layout;
    this.updates = snapshotUpdates(updates);
    this.policy = policy;
    this.changes = changes === undefined
      ? patchesFor(this.updates, layout.root)
      : Object.freeze([...changes]);
    Object.freeze(this);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const previousRoot = context.roots.get(this.#layout, () => this.#layout.root);
    const updates = this.updates.filter(
      (update) => {
        const previousProvenance = previousRoot.rowHeightProvenance.get(update.row) ?? "default";
        if (this.policy === "preserve-manual" && previousProvenance === "manual") {
          return false;
        }
        return previousRoot.rowMetrics.heightAt(update.row) !== update.height ||
          previousProvenance !== update.provenance;
      }
    );
    const changedUpdates = snapshotUpdates(updates);
    const nextRoot = changedUpdates.length === 0
      ? previousRoot
      : this.#layout.prepareSetRowHeights(
          new Map(changedUpdates.map((update) => [update.row, Object.freeze({
            height: update.height,
            provenance: update.provenance
          })])),
          previousRoot
        );
    if (nextRoot !== previousRoot) {
      context.roots.set(this.#layout, nextRoot);
    }

    const changes = patchesFor(changedUpdates, nextRoot);
    const inverseUpdates = snapshotUpdates(changedUpdates.map((update) => ({
      row: update.row,
      height: previousRoot.rowMetrics.heightAt(update.row),
      provenance: previousRoot.rowHeightProvenance.get(update.row) ?? "default"
    })));
    // Receipt replay is exact: only the initial auto-measurement prepare protects manual rows.
    const forward = new RowHeightMutation(this.#layout, changedUpdates, "overwrite", changes);
    const inverse = new RowHeightMutation(this.#layout, inverseUpdates, "overwrite", changes);
    let commitStarted = false;
    let rolledBack = false;

    return {
      hasEffect: nextRoot !== previousRoot,
      forward,
      inverse,
      estimatedBytes: changedUpdates.length === 0 ? 0 : 64 + changedUpdates.length * 32,
      commit: () => {
        if (this.#layout.root !== previousRoot) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Layout root changed after row-height prepare");
        }
        if (nextRoot === previousRoot) {
          return;
        }
        commitStarted = true;
        this.#layout.replaceRoot(nextRoot);
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        this.#layout.replaceRoot(previousRoot);
      },
      dispose: () => {}
    };
  }
}
