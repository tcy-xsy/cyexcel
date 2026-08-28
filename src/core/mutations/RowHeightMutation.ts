import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { LayoutEngine, LayoutRoot, RowHeightProvenance } from "../layout";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { Patch } from "../types";
import { rowHeightOf, rowHeightProvenanceOf, withRowHeight } from "./rowHeightMeta";
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

/**
 * Atomic row-height batch whose prepare composes through the transaction-local Layout root.
 *
 * 一条变更改的是**两处**状态：权威（`rowData[row].h`/`.hp`，走 `set-row-meta`）与
 * 索引（`RowMetrics`，走事务局部的布局根）。它们由同一个参与者一起 commit、一起
 * rollback，所以不存在「索引前进了而权威没有」的中间态——ADR 0004 阶段 4B T8 的
 * 不变量（设计稿 §4.2）靠的正是这一点。
 */
export class RowHeightMutation implements Mutation {
  readonly type = "row-heights";
  readonly updates: readonly RowHeightMutationUpdate[];
  readonly policy: RowHeightMutationPolicy;
  readonly changes: readonly Patch[];
  readonly #layout: LayoutEngine;
  readonly #sheet: SheetModel;

  constructor(
    layout: LayoutEngine,
    sheet: SheetModel,
    updates: readonly RowHeightMutationUpdate[],
    policy: RowHeightMutationPolicy = "overwrite",
    changes?: readonly Patch[]
  ) {
    if (policy !== "overwrite" && policy !== "preserve-manual") {
      throw invalidUpdate(`Invalid row-height mutation policy: ${String(policy)}`);
    }
    this.#layout = layout;
    this.#sheet = sheet;
    this.updates = snapshotUpdates(updates);
    this.policy = policy;
    this.changes = changes === undefined
      ? patchesFor(this.updates, layout.root)
      : Object.freeze([...changes]);
    Object.freeze(this);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const previousRoot = context.roots.get(this.#layout, () => this.#layout.root);
    const defaultRowHeight = this.#layout.rowHeight;
    // 读的是影子里的权威，不是 `previousRoot` 的索引：同一个事务里前一条变更已经
    // 改过它了，而索引与权威在这一刻按不变量是同一个答案——既然如此，就读那个说了算的。
    const previousMeta = new Map<number, ReturnType<MutationPrepareContext["readRowMeta"]>>();
    const metaOf = (row: number) => {
      if (!previousMeta.has(row)) {
        previousMeta.set(row, context.readRowMeta(this.#sheet, row));
      }
      return previousMeta.get(row);
    };

    const updates = this.updates.filter((update) => {
      const meta = metaOf(update.row);
      const previousProvenance = rowHeightProvenanceOf(meta);
      if (this.policy === "preserve-manual" && previousProvenance === "manual") {
        return false;
      }
      return rowHeightOf(meta, defaultRowHeight) !== update.height ||
        previousProvenance !== update.provenance;
    });
    const changedUpdates = snapshotUpdates(updates);

    const staged: PreparedSheetMutation[] = [];
    let nextRoot = previousRoot;
    try {
      for (const update of changedUpdates) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-row-meta",
          row: update.row,
          meta: withRowHeight(metaOf(update.row), update.height, update.provenance)
        }));
      }
      nextRoot = changedUpdates.length === 0
        ? previousRoot
        : this.#layout.prepareSetRowHeights(
            new Map(changedUpdates.map((update) => [update.row, update.height])),
            previousRoot
          );
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }
    if (nextRoot !== previousRoot) {
      context.roots.set(this.#layout, nextRoot);
    }

    const changes = patchesFor(changedUpdates, nextRoot);
    const inverseUpdates = snapshotUpdates(changedUpdates.map((update) => {
      const meta = metaOf(update.row);
      return {
        row: update.row,
        height: rowHeightOf(meta, defaultRowHeight),
        provenance: rowHeightProvenanceOf(meta)
      };
    }));
    // Receipt replay is exact: only the initial auto-measurement prepare protects manual rows.
    const forward = new RowHeightMutation(this.#layout, this.#sheet, changedUpdates, "overwrite", changes);
    const inverse = new RowHeightMutation(this.#layout, this.#sheet, inverseUpdates, "overwrite", changes);
    let commitStarted = false;
    let rolledBack = false;

    return {
      // 只改来源、不改高度也是有效果的：`nextRoot === previousRoot`（索引不动），
      // 但权威里的 `hp` 变了。所以效果看的是 `changedUpdates`，不是根的同一性。
      hasEffect: changedUpdates.length > 0,
      forward,
      inverse,
      estimatedBytes: changedUpdates.length === 0
        ? 0
        : 64 + changedUpdates.length * 32 +
          staged.reduce((total, prepared) => total + prepared.estimatedBytes, 0),
      commit: () => {
        if (this.#layout.root !== previousRoot) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Layout root changed after row-height prepare");
        }
        if (changedUpdates.length === 0) {
          return;
        }
        commitStarted = true;
        for (const prepared of staged) {
          prepared.commit();
        }
        if (nextRoot !== previousRoot) {
          this.#layout.replaceRoot(nextRoot);
        }
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        this.#layout.replaceRoot(previousRoot);
        for (let index = staged.length - 1; index >= 0; index -= 1) {
          staged[index]!.rollback();
        }
      },
      dispose: () => {
        for (const prepared of staged) {
          prepared.dispose();
        }
      }
    };
  }
}
