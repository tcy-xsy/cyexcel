import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRect } from "../layout";
import type { SceneCellFragment, SceneSnapshot } from "../scene";
import type { CellRange } from "../types";
import type { Canvas2DLike } from "./Canvas2DLike";
import type { BorderSide, CellStyle } from "../style/types";
import type { DirtyState } from "./RenderScheduler";
import type { CellValue } from "../sheet/types";
import { CAPSULE_RADIUS, computeCapsule, type CapsuleGeometry, type CellDropdown } from "./dropdownGeometry";
import {
  CELL_FONT,
  CELL_PADDING,
  FROZEN_SEPARATOR_COLOR,
  FROZEN_SEPARATOR_THICKNESS,
  DROPDOWN_ARROW_COLOR,
  DROPDOWN_ARROW_ZONE,
  DROPDOWN_CAPSULE_COLOR,
  GRID_LINE_COLOR,
  INVALID_MARK_COLOR,
  INVALID_MARK_SIZE,
  TEXT_COLOR
} from "./theme";

export interface CellPaintContext {
  readonly ctx: Canvas2DLike;
  readonly rect: CellRect;
  readonly row: number;
  readonly column: number;
  readonly value: CellValue | undefined;
  /**
   * 这一格生效的样式（列默认 + 格样式合并后的结果），没有则 `undefined`。
   * 自定义 renderer 拿得到它，但**要不要用是 renderer 自己的事**——核心不会替一个
   * 接管了绘制的 renderer 去套字体或底色。
   */
  readonly style?: CellStyle;
  /**
   * 这一格的图片 URL，没有则 `undefined`。跟 `style` 一样是给 painter（自定义
   * 或内置）用的信息，画不画是 painter 自己的事——`createColumnAwarePainter`
   * 拿到它会让图片优先于列配置画出来。
   */
  readonly image?: string;
}

/**
 * Contract is synchronous (return value ignored except to detect a Promise
 * misuse — see `checkSyncRenderers`). Typed `unknown` rather than `void` so
 * that check isn't silently defeated by TypeScript's lenient void-return
 * position.
 */
export type CellPainter = (context: CellPaintContext) => unknown;

export interface ContentLayerRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly getScene: () => SceneSnapshot;
  readonly getCellValue: (row: number, column: number) => CellValue | undefined;
  /** 与 getCellValue 对称：渲染层继续不认识 SheetModel，样式也从回调进来。 */
  readonly getCellStyle?: (row: number, column: number) => CellStyle | undefined;
  /** 与 getCellStyle 同样的理由：单元格级图片也从回调进来，不认识 SheetModel。 */
  readonly getCellImage?: (row: number, column: number) => string | undefined;
  readonly painter?: CellPainter;
  readonly onError?: (error: unknown, cell: { row: number; column: number }) => void;
  /** Throw immediately if a painter returns a thenable. Default: true. */
  readonly checkSyncRenderers?: boolean;
  /** Draw a light gridline around every cell after it's painted. Default: true. */
  readonly showGridLines?: boolean;
  /**
   * Returns every merge intersecting `range` — a range query, not a per-cell
   * point query, so a fragment covering no merges (the common case, and the
   * only case when the sheet has zero merges) costs exactly one call, not
   * one call per visible cell. See `#paintFragment`.
   */
  readonly getMergesInRange?: (range: CellRange) => readonly CellRange[];
  /**
   * 这一格的值是否违反了它所在的数据验证规则。画在 painter 之外（网格线和边框之后），
   * 所以自定义 renderer 不必各自处理；合并单元格按合并后的矩形算右上角。
   * 每帧对每个可见格调用一次，实现必须便宜——没有规则时应当是一次数组长度判断。
   */
  readonly isCellInvalid?: (row: number, column: number) => boolean;
  /**
   * 这一格的 list 下拉要怎么画（数据验证的 `list` 规则）：`arrow` 文字加右侧箭头，`text`
   * 只有文字，`custom` 带底色的胶囊标签（文字和箭头都画在胶囊里）；颜色见 `CellDropdown`。
   * 和"无效"角标一样画在 painter 之外，常驻、不看选区——同 Univer 的 `DropdownWidget`：
   * 箭头是绘制的一部分，点击由画布的命中测试处理。每帧对每个可见格调用一次，实现必须便宜。
   */
  readonly getCellDropdown?: (row: number, column: number) => CellDropdown | undefined;
}

function defaultPainter(context: CellPaintContext): void {
  const { ctx, rect, value } = context;
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = CELL_FONT;
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  ctx.fillText(
    value === undefined || value === null ? "" : String(value),
    rect.x + CELL_PADDING,
    rect.y + rect.height / 2
  );
}

function isThenable(value: unknown): boolean {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as { then?: unknown }).then === "function"
  );
}

/**
 * One of up to four rectangular row×column combinations a paint pass needs
 * to handle separately — frozen rows/columns are always shown regardless
 * of scroll, so they can't share one "visible range" with the scrollable
 * body the way a single un-frozen table could. See doc/architecture.md
 * §6.1 for why this is modeled as up to four independent regions rather
 * than one merged range.
 */
/**
 * Paints cell values onto a Canvas2DLike surface, clipped to whatever is
 * currently visible per SceneSnapshot. Never
 * depends on SheetModel/TableStore directly — only a getCellValue callback
 * — so it can be unit-tested without constructing a real data layer.
 */
export class ContentLayerRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #getScene: () => SceneSnapshot;
  readonly #getCellValue: (row: number, column: number) => CellValue | undefined;
  readonly #getCellStyle: ((row: number, column: number) => CellStyle | undefined) | undefined;
  readonly #getCellImage: ((row: number, column: number) => string | undefined) | undefined;
  readonly #painter: CellPainter;
  readonly #onError:
    | ((error: unknown, cell: { row: number; column: number }) => void)
    | undefined;
  readonly #checkSyncRenderers: boolean;
  readonly #showGridLines: boolean;
  readonly #getMergesInRange: ContentLayerRendererOptions["getMergesInRange"];
  readonly #isCellInvalid: ContentLayerRendererOptions["isCellInvalid"];
  readonly #getCellDropdown: ContentLayerRendererOptions["getCellDropdown"];

  constructor(options: ContentLayerRendererOptions) {
    this.#ctx = options.ctx;
    this.#getScene = options.getScene;
    this.#getCellValue = options.getCellValue;
    this.#getCellStyle = options.getCellStyle;
    this.#getCellImage = options.getCellImage;
    this.#painter = options.painter ?? defaultPainter;
    this.#onError = options.onError;
    this.#checkSyncRenderers = options.checkSyncRenderers ?? true;
    this.#showGridLines = options.showGridLines ?? true;
    this.#getMergesInRange = options.getMergesInRange;
    this.#isCellInvalid = options.isCellInvalid;
    this.#getCellDropdown = options.getCellDropdown;
  }

  paintContent(dirty: DirtyState, scene: SceneSnapshot = this.#getScene()): void {
    if (dirty.kind === "none") {
      return;
    }
    if (dirty.kind === "full") {
      // Quadrant painting only clears each individual cell's own rect, at
      // its *current* position — fine as long as content only ever grows
      // or keeps the same footprint, but a structural change that shrinks
      // total content width/height (Step 17: hideColumn()/deleteColumns())
      // leaves whatever was painted at the old, now-uncovered trailing
      // edge on screen forever otherwise. A full repaint is exactly the
      // case where clearing the whole viewport first is both correct and
      // cheap (it's about to repaint every visible cell anyway).
      if (scene.width > 0 && scene.height > 0) {
        this.#ctx.clearRect(0, 0, scene.width, scene.height);
      }
      for (const fragment of scene.getVisibleFragments()) {
        this.#paintFragment(scene, fragment);
      }
    } else {
      for (const range of dirty.ranges) {
        for (const fragment of scene.getVisibleFragmentsForRange(range)) {
          this.#paintFragment(scene, fragment);
        }
      }
    }

    this.#paintFrozenSeparators(scene);
  }

  /**
   * A heavier line marking the fixed screen boundary between the frozen
   * band and the scrollable body — always at the same pixel position
   * regardless of scroll, since that's what "frozen" means. Drawn after
   * every paint pass (not just "full") because a regions repaint can clear
   * pixels right up to that boundary without necessarily repainting it.
   */
  #paintFrozenSeparators(scene: SceneSnapshot): void {
    const frozenCorner = scene.getPanes("body").find((pane) => pane.id === "frozen-corner");
    if (!frozenCorner) {
      return;
    }

    if (frozenCorner.rows.count > 0 && frozenCorner.clipRect.height > 0) {
      const y = frozenCorner.clipRect.y + frozenCorner.clipRect.height;
      this.#ctx.fillStyle = FROZEN_SEPARATOR_COLOR;
      this.#ctx.fillRect(
        0,
        y - FROZEN_SEPARATOR_THICKNESS / 2,
        scene.width,
        FROZEN_SEPARATOR_THICKNESS
      );
    }

    if (frozenCorner.columns.count > 0 && frozenCorner.clipRect.width > 0) {
      const x = frozenCorner.clipRect.x + frozenCorner.clipRect.width;
      this.#ctx.fillStyle = FROZEN_SEPARATOR_COLOR;
      this.#ctx.fillRect(
        x - FROZEN_SEPARATOR_THICKNESS / 2,
        0,
        FROZEN_SEPARATOR_THICKNESS,
        scene.height
      );
    }
  }

  #paintFragment(scene: SceneSnapshot, fragment: SceneCellFragment): void {
    const clipRect = fragment.pane.clipRect;
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
    this.#ctx.clip();

    const lastRow = fragment.rows.start + fragment.rows.count - 1;
    const lastColumn = fragment.columns.start + fragment.columns.count - 1;

    // Every (row, column) this pass has already painted as part of a merged
    // cell's single spanning rect — a fresh Set on every #paintFragment call,
    // never a field: a persistent Set would go stale (new merges created
    // after the first frame would never get re-scanned into it) and could
    // leak entries from one pane's fragment into another's.
    const mergeCoveredCells = new Set<string>();
    const intersectingMerges = this.#getMergesInRange?.({
      startRow: fragment.rows.start,
      endRow: lastRow,
      startColumn: fragment.columns.start,
      endColumn: lastColumn
    });
    if (intersectingMerges && intersectingMerges.length > 0) {
      for (const merge of intersectingMerges) {
        const mergeFragment = scene
          .getFragmentsForRange(merge)
          .find((candidate) => candidate.pane.id === fragment.pane.id);
        if (!mergeFragment) {
          continue;
        }
        for (const row of scene.getRenderableRows({
          start: Math.max(merge.startRow, fragment.rows.start),
          count: Math.min(merge.endRow, lastRow) - Math.max(merge.startRow, fragment.rows.start) + 1
        })) {
          for (
            let column = Math.max(merge.startColumn, fragment.columns.start);
            column <= Math.min(merge.endColumn, lastColumn);
            column += 1
          ) {
            mergeCoveredCells.add(`${row},${column}`);
          }
        }
        this.#paintCellAt(merge.startRow, merge.startColumn, mergeFragment.rect);
      }
    }

    for (const row of scene.getRenderableRows(fragment.rows)) {
      for (let column = fragment.columns.start; column <= lastColumn; column += 1) {
        if (mergeCoveredCells.has(`${row},${column}`)) {
          continue;
        }
        this.#paintCell(scene, row, column, fragment);
      }
    }

    this.#ctx.restore();
  }

  #paintCell(scene: SceneSnapshot, row: number, column: number, fragment: SceneCellFragment): void {
    const rect = scene.getCellRectInPane({ row, column }, fragment.pane);
    this.#paintCellAt(row, column, rect);
  }

  #paintCellAt(row: number, column: number, rect: CellRect): void {
    if (rect.width <= 0 || rect.height <= 0) {
      return;
    }
    this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);

    const value = this.#getCellValue(row, column);
    const style = this.#getCellStyle?.(row, column);
    const image = this.#getCellImage?.(row, column);
    const dropdown = this.#getCellDropdown?.(row, column);
    // 背景在文字之前：Canvas 是后画的盖先画的，顺序反了底色会把字糊掉。
    // 放在 painter 之外，所以自定义 renderer 不必自己记得画底色。格子自己设了背景色就用它；
    // 没有的话，`arrow` / `text` 模式的选项颜色当作单元格底色（`custom` 的颜色在胶囊上）。
    const background =
      style?.background ?? (dropdown !== undefined && dropdown.mode !== "custom" ? dropdown.color : undefined);
    if (background !== undefined) {
      this.#ctx.fillStyle = background;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }

    // `custom` 模式：文字和箭头都在胶囊里，这一格不交给 painter（同 Univer 的
    // `skipDefaultFontRender`）。胶囊挤不下（格子太窄或太矮）时退回 `arrow` 模式。
    let capsule: CapsuleGeometry | null = null;
    let capsuleText = "";
    if (dropdown?.mode === "custom") {
      capsuleText = value === undefined || value === null ? "" : String(value);
      capsule = computeCapsule(rect, this.#measureTextWidth(capsuleText));
    }
    if (dropdown !== undefined && capsule !== null) {
      this.#paintCapsule(capsule, capsuleText, dropdown.color);
      this.#paintGridLine(rect);
      this.#paintBorder(rect, style);
      this.#paintInvalidMark(row, column, rect);
      return;
    }
    const arrow = dropdown !== undefined && dropdown.mode !== "text";
    // list 单元格右侧给下拉箭头留出位置（同 Univer 的 `calcCellAutoWidth` 给箭头预留
    // 宽度）：painter 拿到的 rect 右边缘内缩一个箭头区，文字的对齐、换行宽度都按这个窄
    // rect 算，并且裁剪到它——过长的文字在箭头前截断，不会画到箭头下面。格子比箭头区还窄
    // 时不预留（否则整格什么都画不出来）。背景、网格线、边框仍按完整的 rect。
    // `text` 模式没有箭头，也就不预留。
    const paintRect: CellRect =
      arrow && rect.width > DROPDOWN_ARROW_ZONE
        ? { x: rect.x, y: rect.y, width: rect.width - DROPDOWN_ARROW_ZONE, height: rect.height }
        : rect;
    // 每个单元格的绘制都裁剪到它自己的矩形（list 单元格是扣掉箭头区之后的窄矩形）。
    // 不裁剪的话，过长的文字会画出单元格右边界：中间的列看不出来，因为右邻格随后的
    // `clearRect` 会盖掉溢出的部分，最后一列右边没有邻格，字就露在表格外面了。
    // Excel 的规则是"右邻格为空才允许溢出"，这里没有实现，统一裁剪更可预期。
    const clipToPaintRect = (): void => {
      this.#ctx.beginPath();
      this.#ctx.rect(paintRect.x, paintRect.y, paintRect.width, paintRect.height);
      this.#ctx.clip();
    };
    const context: CellPaintContext = {
      ctx: this.#ctx,
      rect: paintRect,
      row,
      column,
      value,
      ...(style !== undefined ? { style } : {}),
      ...(image !== undefined ? { image } : {})
    };

    // Isolated in its own save/restore so a painter (built-in default or a
    // custom CanvasCellRenderer) that changes ctx state — font, fillStyle,
    // textAlign, ... — and never restores it itself can't leak that state
    // into the next cell painted on this same canvas. Canvas 2D state is
    // otherwise persistent across draw calls; without this, one column's
    // renderer choosing e.g. a different font silently changes how every
    // later cell that never sets its own font gets drawn.
    this.#ctx.save();
    clipToPaintRect();
    let result: unknown;
    try {
      result = this.#painter(context);
    } catch (error) {
      this.#ctx.restore();
      this.#onError?.(error, { row, column });
      this.#ctx.save();
      clipToPaintRect();
      defaultPainter(context);
      this.#ctx.restore();
      this.#paintGridLine(rect);
      this.#paintBorder(rect, style);
      this.#paintInvalidMark(row, column, rect);
      this.#paintDropdownArrow(arrow, rect);
      return;
    }
    this.#ctx.restore();

    if (this.#checkSyncRenderers && isThenable(result)) {
      throw new CyExcelError(
        CyExcelErrorCode.AsyncRendererNotSupported,
        `Cell renderer for (${row}, ${column}) returned a Promise; cell renderers must be synchronous`
      );
    }

    this.#paintGridLine(rect);
    this.#paintBorder(rect, style);
    this.#paintInvalidMark(row, column, rect);
    this.#paintDropdownArrow(arrow, rect);
  }

  /**
   * 右侧居中的向下小三角。箭头中心落在命中区（`DROPDOWN_ARROW_ZONE`）的正中，格子比
   * 命中区还窄时贴着格子边缘。自己 save/restore，理由同无效角标。
   */
  #paintDropdownArrow(arrow: boolean, rect: CellRect): void {
    if (!arrow) {
      return;
    }
    const zone = Math.min(DROPDOWN_ARROW_ZONE, rect.width);
    this.#paintArrowGlyph(rect.x + rect.width - zone / 2, rect.y + rect.height / 2);
  }

  #paintArrowGlyph(centerX: number, centerY: number): void {
    this.#ctx.save();
    this.#ctx.fillStyle = DROPDOWN_ARROW_COLOR;
    this.#ctx.beginPath();
    this.#ctx.moveTo(centerX - 3.5, centerY - 2);
    this.#ctx.lineTo(centerX + 3.5, centerY - 2);
    this.#ctx.lineTo(centerX, centerY + 2.5);
    this.#ctx.closePath();
    this.#ctx.fill();
    this.#ctx.restore();
  }

  /** 胶囊内文字用的宽度：与画字用同一个字体，几何（画和命中）才对得上。 */
  #measureTextWidth(text: string): number {
    if (text === "") {
      return 0;
    }
    this.#ctx.save();
    this.#ctx.font = CELL_FONT;
    const width = this.#ctx.measureText(text).width;
    this.#ctx.restore();
    return width;
  }

  /**
   * `custom` 模式的胶囊：圆角底色、裁剪过的文字、箭头。Canvas2DLike 没有 `roundRect`，
   * 用四段 `arc` 拼圆角矩形。自己 save/restore，不把 fillStyle、font 漏给下一格。
   */
  #paintCapsule(geometry: CapsuleGeometry, text: string, color: string | undefined): void {
    const { x, y, width, height } = geometry.capsule;
    const r = Math.min(CAPSULE_RADIUS, height / 2, width / 2);
    const ctx = this.#ctx;
    ctx.save();
    ctx.fillStyle = color ?? DROPDOWN_CAPSULE_COLOR;
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + width - r, y);
    ctx.arc(x + width - r, y + r, r, -Math.PI / 2, 0);
    ctx.lineTo(x + width, y + height - r);
    ctx.arc(x + width - r, y + height - r, r, 0, Math.PI / 2);
    ctx.lineTo(x + r, y + height);
    ctx.arc(x + r, y + height - r, r, Math.PI / 2, Math.PI);
    ctx.lineTo(x, y + r);
    ctx.arc(x + r, y + r, r, Math.PI, Math.PI * 1.5);
    ctx.closePath();
    ctx.fill();
    ctx.restore();

    if (text !== "" && geometry.textMaxWidth > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(geometry.textX, y, geometry.textMaxWidth, height);
      ctx.clip();
      ctx.fillStyle = TEXT_COLOR;
      ctx.font = CELL_FONT;
      ctx.textBaseline = "middle";
      ctx.textAlign = "left";
      ctx.fillText(text, geometry.textX, y + height / 2);
      ctx.restore();
    }
    this.#paintArrowGlyph(geometry.arrowCenter.x, geometry.arrowCenter.y);
  }

  /** 右上角的红色小三角。自己 save/restore：不能把 fillStyle 漏给下一格。 */
  #paintInvalidMark(row: number, column: number, rect: CellRect): void {
    if (this.#isCellInvalid === undefined || !this.#isCellInvalid(row, column)) {
      return;
    }
    const size = Math.min(INVALID_MARK_SIZE, rect.width, rect.height);
    const right = rect.x + rect.width;
    this.#ctx.save();
    this.#ctx.fillStyle = INVALID_MARK_COLOR;
    this.#ctx.beginPath();
    this.#ctx.moveTo(right - size, rect.y);
    this.#ctx.lineTo(right, rect.y);
    this.#ctx.lineTo(right, rect.y + size);
    this.#ctx.closePath();
    this.#ctx.fill();
    this.#ctx.restore();
  }

  /**
   * 边框画在网格线**之后**，所以盖得住网格线。
   *
   * 相邻两格共一条像素线（A1 的右边框与 B1 的左边框是同一条），规则是**后画的胜**
   * ——绘制按行列顺序推进，于是右边/下边那一格说了算。Excel 的真实规则是"最后一次
   * 设置的胜"，那要记录每条边的设置时序，不追。
   */
  #paintBorder(rect: CellRect, style: CellStyle | undefined): void {
    const border = style?.border;
    if (border === undefined) {
      return;
    }
    const ctx = this.#ctx;
    const line = (side: BorderSide | undefined, x1: number, y1: number, x2: number, y2: number): void => {
      if (side === undefined) {
        return;
      }
      ctx.strokeStyle = side.color;
      ctx.lineWidth = side.width ?? 1;
      ctx.beginPath();
      ctx.moveTo(x1, y1);
      ctx.lineTo(x2, y2);
      ctx.stroke();
    };
    const right = rect.x + rect.width;
    const bottom = rect.y + rect.height;
    line(border.top, rect.x, rect.y, right, rect.y);
    line(border.right, right, rect.y, right, bottom);
    line(border.bottom, rect.x, bottom, right, bottom);
    line(border.left, rect.x, rect.y, rect.x, bottom);
  }

  #paintGridLine(rect: CellRect): void {
    if (!this.#showGridLines) {
      return;
    }
    this.#ctx.strokeStyle = GRID_LINE_COLOR;
    this.#ctx.lineWidth = 1;
    this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
  }

}
