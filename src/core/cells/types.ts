import type { CellEditor } from "../editors";
import type { DomPointerOwnership } from "../interaction/domPointerOwnership";
import type { CellRect } from "../layout";
import type { Canvas2DLike } from "../render";
import type { CellValue } from "../sheet/types";

/**
 * Encapsulates a value's display, parsing, comparison and editing behavior.
 * Matches doc/architecture.md §8's published shape. `createEditor`/`serialize`
 * are part of that published contract but are not wired into EditorManager or
 * a clipboard pipeline yet (Step 8 scope — see the Step 8 progress.md entry's
 * remaining risks); they exist so custom types can already declare them
 * without a breaking shape change later.
 */
export interface CellFormatContext {
  readonly row: number;
  readonly column: RuntimeColumn;
}

/** An enumerated value's candidate list — currently only meaningful for the "select" type's `format()`. */
export interface SelectOption {
  readonly value: string;
  /** Defaults to `value` when not given. */
  readonly label?: string;
}

/**
 * `T` 的默认值是 `CellValue`（`string | number | boolean`），不是 `unknown`——
 * 单元格能存的就是这三种（见 `sheet/types.ts` 的 `CellValue`），`parse()` 的返回值
 * 因此也只能是它们之一。这条约束让"快照往返无损"与"导出 xlsx 是直译"从类型上成立。
 */
export interface CellType<T extends CellValue = CellValue> {
  format(value: T | undefined, context: CellFormatContext): string;
  paint?(context: CellRenderContext<T>): void;
  parse?(text: string, context: CellFormatContext): T;
  compare?(left: T, right: T): number;
  /**
   * Whether values of this type have a meaningful order at all. Defaults to
   * `true`; a type sets it `false` to opt out of sorting entirely, which is
   * different from having no `compare` (that only means "string order is
   * already right"). A column can override it either way.
   */
  readonly sortable?: boolean;
  /**
   * Whether values of this type can be filtered at all. Defaults to `true`,
   * same shape as `sortable` — a column can override it either way.
   */
  readonly filterable?: boolean;
  createEditor?(context: CellFormatContext): CellEditor;
  serialize?(value: T | undefined, context: CellFormatContext): string;
  /** Metadata only — does not itself select an editor. See doc/progress.md's editor-pluggability patch. */
  readonly options?: readonly SelectOption[];
}

/** A registrable named type; `extends` names a previously-registered type to inherit from. */
export interface CellTypeDefinition<T extends CellValue = CellValue> extends Partial<CellType<T>> {
  readonly extends?: string;
}

/**
 * Read-only paint/hitTest context handed to a CanvasCellRenderer. Deliberately
 * structurally independent from render's CellPaintContext (no import from
 * "../render") to avoid a render <-> cells dependency cycle — see the Step 8
 * plan's "关键设计决策 3". Exposes nothing that could mutate the Store.
 */
export interface CellRenderContext<T extends CellValue = CellValue> {
  readonly ctx: Canvas2DLike;
  readonly rect: CellRect;
  readonly row: number;
  readonly column: RuntimeColumn;
  readonly value: T | undefined;
}

export interface CellHitTestContext<T extends CellValue = CellValue> extends CellRenderContext<T> {
  readonly point: { readonly x: number; readonly y: number };
}

/**
 * Minimal outcome vocabulary for Step 8 — just enough to prove "hitTest
 * returns a CellAction, which gets converted into a Command" (Step 8's 4th
 * pre-written test). Not the full action set implied by later Hook targets
 * (sort/filter/menu); those get added when the step that needs them lands.
 */
export type CellAction = { readonly type: "set-value"; readonly value: CellValue } | { readonly type: "start-edit" };

export interface CanvasCellRenderer<T extends CellValue = CellValue> {
  readonly kind: "canvas";
  paint(context: CellRenderContext<T>): void;
  hitTest?(context: CellHitTestContext<T>): CellAction | void;
}

export interface DomCellContext<T extends CellValue = CellValue> {
  readonly row: number;
  readonly column: RuntimeColumn;
  readonly value: T | undefined;
  readonly rect: CellRect;
  readonly actions: {
    readonly commit: (value: T) => void;
  };
}

export interface DomCellRenderer<T extends CellValue = CellValue> {
  mount(container: HTMLElement, context: DomCellContext<T>): void;
  update(context: DomCellContext<T>): void;
  destroy(): void;
}

export interface DomCellRendererFactory<T extends CellValue = CellValue> {
  readonly kind: "dom";
  /** Terminal ownership for events retargeted to the mounted host, including closed-shadow events. */
  readonly pointerOwnership: DomPointerOwnership;
  create(): DomCellRenderer<T>;
}

/**
 * Header rendering/hit-test state a column can display — always an
 * externally-supplied static value (there's no sort/filter Command yet, so
 * nothing ever computes this itself; see the Step 9 plan's "不做的事"). Just
 * enough for `CanvasHeaderRenderer.paint()` to have something meaningful to
 * draw for "correctly renders title, sort, and filter state" (Step 9's 1st
 * pre-written test).
 */
export interface HeaderState {
  readonly sorted?: "asc" | "desc" | false;
  readonly filtered?: boolean;
  /**
   * Whether this column offers sorting at all. `false` removes the header's
   * sort affordance and makes `defaultHeaderHitTest` refuse the trailing zone,
   * so the column keeps its whole width for click-to-select. Omitted means
   * sortable — a renderer used standalone, with no state supplied, still shows
   * the affordance.
   */
  readonly sortable?: boolean;
  /**
   * Whether this column offers filtering at all. Same shape as `sortable`:
   * `false` removes the header's filter affordance and makes
   * `defaultHeaderHitTest` refuse that zone. Omitted means filterable.
   */
  readonly filterable?: boolean;
}

/**
 * Read-only paint/hitTest context for a leaf column's header cell.
 * Structurally independent from `CellRenderContext` (no `value` — a header
 * represents a column, not a row's data) but follows the same shape
 * conventions as Step 8's cell renderer contracts. `column` is the leaf
 * column index, not a `RuntimeColumn` — headers only ever need `title`
 * (already resolved by HeaderManager, included below) plus which column
 * this is, and keeping it to a plain index means the `headers` module never
 * needs to depend on ColumnManager's resolution pipeline.
 */
export interface HeaderRenderContext {
  readonly ctx: Canvas2DLike;
  readonly rect: CellRect;
  readonly column: number;
  readonly title: string;
  readonly state: HeaderState;
}

export interface HeaderHitTestContext extends HeaderRenderContext {
  readonly point: { readonly x: number; readonly y: number };
}

/**
 * Canvas header interaction outcomes. CyExcelCore emits a `header-action`
 * event after the input receipt closes; hosts can translate it into their
 * own sort/filter command without coupling the renderer to model mutation.
 */
export type HeaderAction = { readonly type: "toggle-sort" } | { readonly type: "toggle-filter" };

export interface CanvasHeaderRenderer {
  readonly kind: "canvas";
  paint(context: HeaderRenderContext): void;
  hitTest?(context: HeaderHitTestContext): HeaderAction | void;
}

export interface DomHeaderContext {
  readonly column: number;
  readonly title: string;
  readonly rect: CellRect;
  readonly state: HeaderState;
}

export interface DomHeaderRenderer {
  mount(container: HTMLElement, context: DomHeaderContext): void;
  update(context: DomHeaderContext): void;
  destroy(): void;
}

export interface DomHeaderRendererFactory {
  readonly kind: "dom";
  /** Terminal ownership for events retargeted to the mounted host, including closed-shadow events. */
  readonly pointerOwnership: DomPointerOwnership;
  create(): DomHeaderRenderer;
}

export interface HeaderDefinition {
  readonly title?: string;
  readonly renderer?: CanvasHeaderRenderer | DomHeaderRendererFactory;
}

/**
 * Per-column clipboard behavior — the only clipboard-related field a column
 * gets in Step 10. `doc/api-design.md`'s full vision has independent
 * `readonly`/`hidden` column fields, but neither is implemented anywhere
 * yet (Step 8/9 deliberately left them out — see their progress.md
 * entries), so this deliberately doesn't add them. A column wanting
 * "readonly" clipboard behavior sets `paste: false`; one wanting "hidden"
 * behavior sets `copy: false`. Both default to `true`.
 */
export interface ColumnClipboardPolicy {
  readonly copy?: boolean;
  readonly paste?: boolean;
}

/**
 * A column's declared configuration. Deliberately minimal — only fields
 * Step 8/9's pre-written tests and exit criteria exercise. See the Step 8
 * plan's "不做的事" for the fuller doc/customization-and-hooks.md field set
 * (readonly/pinned/hidden/validator/comparator/filter/clipboard/meta/...)
 * intentionally left out until the Step that actually needs them.
 *
 * `header`'s contract types live here rather than in `../headers` — that
 * module's HeaderManager needs to read `ColumnDefinition.header`, and
 * `../headers` importing from `cells` while `cells` also imported a
 * `HeaderDefinition` type from `../headers` would be a real import cycle.
 * `../headers/types.ts` re-exports these for anyone browsing that module.
 */
export interface ColumnDefinition<T extends CellValue = CellValue> {
  /*
   * 没有 `id`，也没有 `field`——列在这个模型里没有身份（ADR 0004 §3.8.2）。
   * 一列**就是**它的下标：公开 API 本来就全是下标制（`getCell(row, column)`、
   * `setColumnWidth(column, width)`、`sortByColumn(column)`，没有一个吃 `id`），
   * 而 `field` 随对象数组一起降级成了导入那一刻的参数
   * （`rowsToWorkbookSnapshot` 的 `fields`，§3.2/§4.5），导入之后存储里只有
   * `(row, column)`，再没有字段名可对。
   *
   * 代价是「结构操作后的列身份」：插一列之后下标全移，持有某列引用的调用方会
   * 错位。Excel 就是这样（公式靠 remap 而不是 id），仓库里 `Patch` 的
   * `column-structure` + `remapIndex()` 也已经在做这件事。ADR 已接受。
   */
  /** Column width in logical pixels; defaults to 120. */
  readonly width?: number;
  /** Lower/upper bound a column-resize drag is clamped to (Step 9). */
  readonly minWidth?: number;
  readonly maxWidth?: number;
  /** Registered CellType name; defaults to "text". */
  readonly type?: string;
  readonly renderer?: CanvasCellRenderer<T> | DomCellRendererFactory<T>;
  /** A plain string is shorthand for `{ title: string }`. */
  readonly header?: string | HeaderDefinition;
  /** Overrides the resolved CellType's `options` for this column specifically. */
  readonly options?: readonly SelectOption[];
  /**
   * Overrides the resolved CellType's `createEditor` for this column
   * specifically — the explicit, per-column way to opt into a particular
   * editor without first registering a whole named CellType for it. Not set
   * by any built-in type; see doc/progress.md's editor-pluggability patch.
   */
  readonly createEditor?: (context: CellFormatContext) => CellEditor;
  /** See ColumnClipboardPolicy. Defaults to `{ copy: true, paste: true }`. */
  readonly clipboard?: ColumnClipboardPolicy;
  /**
   * Wraps text onto multiple lines instead of the default single-line
   * `fillText()` — only takes effect for the built-in formatted-text paint
   * path (`createColumnAwarePainter`'s `drawFormattedText()`); a column
   * with a custom `renderer` owns its own painting entirely and this has no
   * effect on it. Default `false` — unset behavior (single line) is
   * unchanged, same "unset is a no-op" precedent as every other optional
   * column field here.
   */
  readonly wrap?: boolean;
  /**
   * Overrides the cell type's own answer to "can this be sorted". Set it
   * `false` to drop the header's sort affordance and refuse click-to-sort for
   * this column; set it `true` to sort a type that opted out. Defaults to the
   * type's `sortable`, which defaults to `true`.
   */
  readonly sortable?: boolean;
  /**
   * Overrides the cell type's own answer to "can this be filtered". Same
   * shape as `sortable` — set it `false` to drop the header's filter
   * affordance for this column. Defaults to the type's `filterable`, which
   * defaults to `true`.
   */
  readonly filterable?: boolean;
}

/**
 * The resolved, cached output of ColumnManager's merge chain. Never
 * re-merged during painting — see ColumnManager.
 */
export interface RuntimeColumn<T extends CellValue = CellValue> extends CellType<T> {
  readonly width: number;
  readonly typeName: string;
  readonly renderer?: CanvasCellRenderer<T> | DomCellRendererFactory<T>;
  readonly clipboard: Required<ColumnClipboardPolicy>;
  readonly wrap: boolean;
  /** Resolved from the column, then the cell type, then `true`. */
  readonly sortable: boolean;
  /** Resolved from the column, then the cell type, then `true`. */
  readonly filterable: boolean;
  /** Bumped only if columns are ever re-resolved; Step 8 resolves once at construction. */
  readonly version: number;
}
