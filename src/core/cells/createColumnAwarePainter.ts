import type { CellPaintContext, CellPainter } from "../render";
import { CELL_LINE_HEIGHT } from "../render/theme";
import { formatCellText } from "../style/formatCellText";
import { resolveDrawStyle, type DrawStyle } from "../style/resolveDrawStyle";
import { wrapText } from "../render/textWrap";
import type { CellRenderContext, RuntimeColumn } from "./types";

/**
 * 文字的起点 x 与 `textAlign` 要一起定：Canvas 的 `textAlign` 决定的是"给定的 x
 * 算左边、中间还是右边"，所以居中时 x 必须是格子中线，而不是左内边距。
 */
function textOriginX(rect: CellPaintContext["rect"], draw: DrawStyle): number {
  switch (draw.align) {
    case "center":
      return rect.x + rect.width / 2;
    case "right":
      return rect.x + rect.width - draw.padding;
    case "left":
      return rect.x + draw.padding;
  }
}

/** 单行文字的基线 y。多行（wrap）不走这里——那条路径自己从顶部往下排。 */
function textBaselineY(rect: CellPaintContext["rect"], draw: DrawStyle): number {
  switch (draw.verticalAlign) {
    case "top":
      return rect.y + draw.padding + CELL_LINE_HEIGHT / 2;
    case "bottom":
      return rect.y + rect.height - draw.padding - CELL_LINE_HEIGHT / 2;
    case "middle":
      return rect.y + rect.height / 2;
  }
}

/**
 * 下划线与删除线：Canvas 没有原生支持，只能自己量出文字宽度再画一条线。
 * 线的起点跟着 `textAlign` 走，所以要按对齐方式把 x 折算回"文字左边缘"。
 */
function drawTextDecoration(
  ctx: CellPaintContext["ctx"],
  text: string,
  originX: number,
  baselineY: number,
  draw: DrawStyle
): void {
  if (!draw.underline && !draw.strike) {
    return;
  }
  const width = ctx.measureText(text).width;
  const left =
    draw.align === "center" ? originX - width / 2 : draw.align === "right" ? originX - width : originX;
  ctx.strokeStyle = draw.color;
  ctx.lineWidth = 1;
  const stroke = (y: number): void => {
    ctx.beginPath();
    ctx.moveTo(left, y);
    ctx.lineTo(left + width, y);
    ctx.stroke();
  };
  if (draw.underline) {
    // 基线下方一点——压在基线上会和字母的下缘糊在一起。
    stroke(Math.round(baselineY) + 6);
  }
  if (draw.strike) {
    stroke(Math.round(baselineY));
  }
}

function drawFormattedText(context: CellPaintContext, column: RuntimeColumn): void {
  const { ctx, rect, row, value, style } = context;
  // 样式的 numFmt 优先，没有才退回列的 CellType.format（Step 24 的向后兼容路径）。
  const text = formatCellText(value, style, () => column.format(value, { row, column }));
  const draw = resolveDrawStyle(style);
  ctx.fillStyle = draw.color;
  ctx.font = draw.font;
  ctx.textBaseline = "middle";
  ctx.textAlign = draw.align;

  if (!column.wrap) {
    const originX = textOriginX(rect, draw);
    const baselineY = textBaselineY(rect, draw);
    ctx.fillText(text, originX, baselineY);
    drawTextDecoration(ctx, text, originX, baselineY, draw);
    return;
  }

  // Multi-line: top-aligned (not vertically centered like the single-line
  // path above — centering N unknown-ahead-of-time lines would need the
  // total block height before drawing the first line, which is exactly
  // what auto-height already computed the row tall enough for, so top
  // alignment reads naturally without that extra bookkeeping here). Lines
  // past rect.height are simply never drawn — #paintQuadrant()'s existing
  // ctx.clip() call already prevents them from bleeding into the next row
  // even without this early exit, but skipping the fillText() calls is
  // cheap and avoids wasted work for a cell whose row wasn't grown to fit.
  const lines = wrapText(ctx, text, rect.width - draw.padding * 2);
  const x = textOriginX(rect, draw);
  let y = rect.y + draw.padding + CELL_LINE_HEIGHT / 2;
  for (const line of lines) {
    if (y - CELL_LINE_HEIGHT / 2 >= rect.y + rect.height) {
      break;
    }
    ctx.fillText(line, x, y);
    drawTextDecoration(ctx, line, x, y, draw);
    y += CELL_LINE_HEIGHT;
  }
}

/**
 * Bridges render's column-agnostic CellPainter extension point
 * (ContentLayerRendererOptions.painter) to Step 8's per-column renderer
 * resolution, without any changes to ContentLayerRenderer itself — errors
 * thrown by a custom CanvasCellRenderer.paint() are left to propagate so
 * ContentLayerRenderer's existing try/catch + onError + default-painter
 * fallback (already tested) handles them, exactly as it does for any other
 * custom `painter`.
 *
 * Takes a getter, not a snapshot array: structural mutation commits replace
 * the manager's cached root, and every later paint must observe that root.
 */
export function createColumnAwarePainter(getColumns: () => readonly RuntimeColumn[]): CellPainter {
  return (context: CellPaintContext) => {
    const columns = getColumns();
    const column = columns[context.column];
    if (!column) {
      return;
    }

    const renderer = column.renderer;
    if (renderer?.kind === "dom") {
      // The DOM overlay (DOMCellManager) owns this cell entirely.
      return;
    }

    if (renderer?.kind === "canvas") {
      const renderContext: CellRenderContext = {
        ctx: context.ctx,
        rect: context.rect,
        row: context.row,
        column,
        value: context.value
      };
      renderer.paint(renderContext);
      return;
    }

    drawFormattedText(context, column);
  };
}
