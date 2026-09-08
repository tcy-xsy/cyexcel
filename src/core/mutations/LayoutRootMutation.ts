import { CyExcelError, CyExcelErrorCode } from "../errors";
import { invertRowOrder, type LayoutEngine, type LayoutRoot, type RowHeightProvenance } from "../layout";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { ColumnMeta, RowMeta } from "../sheet/types";
import type { Patch } from "../types";
import { geometricColumnWidthOf, withColumnHidden, withColumnWidth } from "./columnWidthMeta";
import { rowHeightOf, rowHeightProvenanceOf, withRowHeight } from "./rowHeightMeta";
import type { Mutation, MutationPrepareContext, PreparedMutationParticipant } from "./types";

export type LayoutRootAction =
  | {
      readonly type: "set-row-height";
      readonly row: number;
      readonly height: number;
      readonly provenance: RowHeightProvenance;
    }
  | { readonly type: "set-column-width"; readonly column: number; readonly width: number }
  | { readonly type: "hide-column"; readonly column: number }
  | { readonly type: "show-column"; readonly column: number }
  /**
   * 把一列的元数据整条写成 `meta`（`undefined` 即删掉条目）。
   *
   * 它是**另外三条列动作共同的逆**，也是唯一能表达「这一列本来就没有条目」的动作
   * （ADR 0004 阶段 4B T8b）。行轴那边不需要这样一条：`provenance === "default"` 已经
   * 是「没有显式行高」的写法，`withRowHeight` 见到它就删 `h`/`hp`。列宽没有 provenance
   * 这个第三态，逆若还写 `set-column-width`，就会在一列本来没有条目的地方凭空造出
   * `{ w: 默认宽 }`——撤销之后再改默认列宽，那一列就永远留在旧默认值上，正是 §5.1
   * 要防的那种「往返不逐字段一致」。
   *
   * 它对自己也是逆（写回上一条 `meta`），所以撤销/重做在这条上是对称的。
   */
  | {
      readonly type: "set-column-meta";
      readonly column: number;
      readonly meta: ColumnMeta | undefined;
    }
  /**
   * A whole-axis row permutation (`order[newPosition] = oldPosition`).
   *
   * It belongs here rather than behind an `AxisChangeDescriptor` because a
   * permutation has no `start`/`count`/`target` to put in one, and because a
   * reorder is exactly what this union is for: a root replacement whose
   * inverse is another action of the same kind (here, the inverse
   * permutation) rather than something that has to be read back off the
   * previous root.
   */
  | { readonly type: "reorder-rows"; readonly order: Int32Array }
  /**
   * Replaces the entire `filteredRows` membership (ADR 0003 §3.2/§3.3) — a
   * whole-set replacement, like `reorder-rows` a root replacement whose
   * inverse is another action of the same kind (the previous set), not
   * something read back off the previous root. Stages no data mutation at
   * all: that absence *is* "过滤没有写数据路径" made structural rather than
   * just documented.
   */
  | { readonly type: "set-filtered-rows"; readonly filteredRows: ReadonlySet<number> };

function patchFor(action: LayoutRootAction, root: LayoutRoot): Patch {
  switch (action.type) {
    case "set-row-height":
      return {
        type: "row-layout",
        rows: Object.freeze([action.row]),
        range: {
          startRow: action.row,
          startColumn: 0,
          endRow: root.rowCount - 1,
          endColumn: Math.max(0, root.columnCount - 1)
        }
      };
    case "reorder-rows":
      return { type: "row-reorder", order: action.order };
    case "set-filtered-rows":
      return { type: "row-visibility" };
    case "set-column-width":
    case "hide-column":
    case "show-column":
    case "set-column-meta":
      return { type: "column-layout", column: action.column };
  }
}

/** 这条动作要把 `columnData[column]` 写成什么——`undefined` 即删掉条目。 */
function nextColumnMetaFor(
  action: LayoutRootAction,
  previous: ColumnMeta | undefined
): ColumnMeta | undefined {
  switch (action.type) {
    case "set-column-width":
      return withColumnWidth(previous, action.width);
    case "hide-column":
      return withColumnHidden(previous, true);
    case "show-column":
      return withColumnHidden(previous, false);
    case "set-column-meta":
      return action.meta;
    default:
      return previous;
  }
}

/**
 * 两条列元数据是不是同一份。
 *
 * 逐字段比而不是 `===`：`withColumnWidth` 每次都造新对象，引用比较会把「设成它已经
 * 是的那个宽度」算成一次有效果的变更，进历史、发补丁。判据要与行轴那边一致——看
 * **权威变没变**，不看有没有造出新对象。
 */
function sameColumnMeta(a: ColumnMeta | undefined, b: ColumnMeta | undefined): boolean {
  if (a === undefined || b === undefined) {
    return a === b;
  }
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const key of keys) {
    if (a[key as keyof ColumnMeta] !== b[key as keyof ColumnMeta]) {
      return false;
    }
  }
  return true;
}

/** History budgeting: only a permutation or a filter set carries a payload worth counting. */
function estimatedBytesFor(action: LayoutRootAction): number {
  if (action.type === "reorder-rows") {
    return 96 + action.order.length * 4;
  }
  if (action.type === "set-filtered-rows") {
    return 96 + action.filteredRows.size * 4;
  }
  return 96;
}

function prepareAction(
  layout: LayoutEngine,
  root: LayoutRoot,
  action: LayoutRootAction,
  nextColumnMeta: ColumnMeta | undefined
): LayoutRoot {
  switch (action.type) {
    case "set-row-height":
      // 只喂高度：来源写进 `rowData[row].hp`，那半边由 prepare() 自己发出去
      // （设计稿 §4.3）。
      return layout.prepareSetRowHeights(new Map([[action.row, action.height]]), root);
    case "set-column-width":
    case "hide-column":
    case "show-column":
    case "set-column-meta":
      // 四条列动作在索引这一侧是同一件事：把权威算成几何宽度写进去。隐藏是 0，
      // 没有 `w` 的列取配置里的缺省宽——这正是 §4.2 那条不变量的列轴版本
      // （`layout.getColumnWidth(c) === geometricColumnWidthOf(meta, defaultColumnWidth)`）。
      return layout.prepareSetColumnWidth(
        action.column,
        geometricColumnWidthOf(nextColumnMeta, layout.defaultColumnWidth),
        root
      );
    case "reorder-rows":
      return layout.prepareReorderRows(action.order, root);
    case "set-filtered-rows":
      return layout.prepareSetFilteredRows(action.filteredRows, root);
  }
}

function inverseAction(
  action: LayoutRootAction,
  previousRowMeta: RowMeta | undefined,
  previousColumnMeta: ColumnMeta | undefined,
  defaultRowHeight: number,
  previousFilteredRows: ReadonlySet<number>
): LayoutRootAction {
  switch (action.type) {
    case "set-row-height":
      // 高度和来源都从**权威**（影子里的 `rowData` 条目）读，不再从根上读：根只是
      // 索引，而且它已经不存来源了（ADR 0004 阶段 4B T8）。
      return {
        type: "set-row-height",
        row: action.row,
        height: rowHeightOf(previousRowMeta, defaultRowHeight),
        provenance: rowHeightProvenanceOf(previousRowMeta)
      };
    case "set-column-width":
    case "hide-column":
    case "show-column":
    case "set-column-meta":
      // 四条列动作共用一条逆：把整条 `ColumnMeta` 写回原样。原来那三条各自算逆
      // （宽度从恢复宽度读、隐藏按当前状态在 hide/show 之间挑一条）现在收成这一条，
      // 而且是精确的——「本来没有条目」也表达得出来。
      return { type: "set-column-meta", column: action.column, meta: previousColumnMeta };
    case "reorder-rows":
      // The only action whose inverse needs nothing from `root`: undoing a
      // permutation is applying the inverse permutation.
      return { type: "reorder-rows", order: invertRowOrder(action.order) };
    case "set-filtered-rows":
      // Whole-set replacement: the inverse is just "write back the previous
      // set" (O(1) — this never re-scans the model), same shape as
      // set-column-meta's own inverse.
      return { type: "set-filtered-rows", filteredRows: previousFilteredRows };
  }
}

/**
 * Declarative Layout root replacement whose commit/rollback are reference swaps only.
 *
 * 除了 `reorder-rows`，每一条动作都要写模型：行高的权威在 `rowData`、列宽与隐藏的
 * 权威在 `columnData`，索引分别是 `RowMetrics` 与 `ColumnAxisSnapshot`。两半必须
 * 同进同退，所以 prepare 期顺带暂存一条 `set-row-meta` / `set-column-meta`
 * （设计稿 §4.2/§4.3/§4.4）。`reorder-rows` 是例外：`rowData` 是登记在册的行侧表，
 * 数据侧那条 `reorder-rows` 已经把 `h`/`hp` 一起重排了，这里只补索引。
 */
export class LayoutRootMutation implements Mutation {
  readonly type = "layout-root";
  readonly changes: readonly Patch[];
  readonly #layout: LayoutEngine;
  readonly #sheet: SheetModel;
  readonly #action: LayoutRootAction;

  constructor(
    layout: LayoutEngine,
    sheet: SheetModel,
    action: LayoutRootAction,
    changes?: readonly Patch[]
  ) {
    this.#layout = layout;
    this.#sheet = sheet;
    this.#action = Object.freeze({ ...action });
    this.changes = Object.freeze([...(changes ?? [patchFor(action, layout.root)])]);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const action = this.#action;
    const previousRoot = context.roots.get(this.#layout, () => this.#layout.root);
    const previousRowMeta = action.type === "set-row-height"
      ? context.readRowMeta(this.#sheet, action.row)
      : undefined;
    const isColumnAction = action.type === "set-column-width" ||
      action.type === "hide-column" ||
      action.type === "show-column" ||
      action.type === "set-column-meta";
    const previousColumnMeta = isColumnAction
      ? context.readColumnMeta(this.#sheet, action.column)
      : undefined;
    const nextColumnMeta = nextColumnMetaFor(action, previousColumnMeta);

    // 高度和来源都没变就一条也不暂存：否则一次「设成它已经是的那个高度」会变成
    // 一条有效果的变更，进历史、发补丁。这与旧实现「根没换就没有效果」等价，只是
    // 判据换成了权威。列那边同理，判据是整条 `ColumnMeta` 变没变。
    const rowMetaChanges = action.type === "set-row-height" && (
      rowHeightOf(previousRowMeta, this.#layout.rowHeight) !== action.height ||
      rowHeightProvenanceOf(previousRowMeta) !== action.provenance
    );
    const columnMetaChanges = isColumnAction &&
      !sameColumnMeta(previousColumnMeta, nextColumnMeta);

    const staged: PreparedSheetMutation[] = [];
    let nextRoot: LayoutRoot;
    try {
      if (action.type === "set-row-height" && rowMetaChanges) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-row-meta",
          row: action.row,
          meta: withRowHeight(previousRowMeta, action.height, action.provenance)
        }));
      }
      if (isColumnAction && columnMetaChanges) {
        staged.push(context.prepareData(this.#sheet, {
          type: "set-column-meta",
          column: action.column,
          meta: nextColumnMeta
        }));
      }
      nextRoot = prepareAction(this.#layout, previousRoot, action, nextColumnMeta);
    } catch (error) {
      for (const prepared of staged) {
        prepared.dispose();
      }
      throw error;
    }
    context.roots.set(this.#layout, nextRoot);

    // 列这一侧：几何没动就没有可发的补丁。
    //
    // `ColumnMeta` 里对消费方可见的只有几何宽度（`hd` 已经被折进那个数），所以根没换就
    // 真的什么都没变——发一条 `column-layout` 是在说谎，而且不是白说：插一列时结构那条
    // `column-structure` 已经让整条列轴重画了，再补一条纯属重复；改一列**隐藏着的**列的
    // 恢复宽度则是屏幕上一个像素都没动。
    //
    // 行那一侧不能照搬：`RowMeta` 的 `hp`（来源）是消费方要看的，一次「高度不变、只把
    // 这行锁成 manual」必须发得出去，否则自动行高会继续来抢。所以那边照旧发。
    const publishesPatch = nextRoot !== previousRoot || !isColumnAction;
    const changes = Object.freeze(publishesPatch ? [patchFor(action, nextRoot)] : []);
    const forward = new LayoutRootMutation(this.#layout, this.#sheet, action, changes);
    // No `changes` for the inverse: `changes` describes the *forward* patch.
    // Every other action's patch happens to be direction-independent, and
    // prepare() recomputes `changes` before the inverse's own is ever read,
    // so passing it was harmless — but a reorder's patch carries the
    // permutation itself, so handing the inverse the forward one loads a gun
    // for any future code that reads `mutation.changes` on an unprepared
    // inverse and gets the wrong direction for a sort undo. The
    // constructor's `patchFor(action, layout.root)` fallback derives it from
    // the inverse action instead.
    const inverse = new LayoutRootMutation(
      this.#layout,
      this.#sheet,
      inverseAction(action, previousRowMeta, previousColumnMeta, this.#layout.rowHeight, previousRoot.filteredRows)
    );
    let commitStarted = false;
    let rolledBack = false;

    return {
      // 一次只改来源的行高写不动根，但它确实改了权威里的 `hp`。
      hasEffect: nextRoot !== previousRoot || staged.length > 0,
      forward,
      inverse,
      estimatedBytes: estimatedBytesFor(action),
      commit: () => {
        if (this.#layout.root !== previousRoot) {
          throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Layout root changed after mutation prepare");
        }
        commitStarted = true;
        for (const prepared of staged) {
          prepared.commit();
        }
        this.#layout.replaceRoot(nextRoot);
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
