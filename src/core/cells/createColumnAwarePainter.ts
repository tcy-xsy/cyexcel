import type { CellPaintContext, CellPainter } from "../render";
import { CELL_FONT, CELL_LINE_HEIGHT, CELL_PADDING, TEXT_COLOR } from "../render/theme";
import { wrapText } from "../render/textWrap";
import type { CellRenderContext, RuntimeColumn } from "./types";

function drawFormattedText(context: CellPaintContext, column: RuntimeColumn): void {
  const { ctx, rect, row, value } = context;
  const text = column.format(value, { row, column });
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = CELL_FONT;
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";

  if (!column.wrap) {
    ctx.fillText(text, rect.x + CELL_PADDING, rect.y + rect.height / 2);
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
  const lines = wrapText(ctx, text, rect.width - CELL_PADDING * 2);
  const x = rect.x + CELL_PADDING;
  let y = rect.y + CELL_PADDING + CELL_LINE_HEIGHT / 2;
  for (const line of lines) {
    if (y - CELL_LINE_HEIGHT / 2 >= rect.y + rect.height) {
      break;
    }
    ctx.fillText(line, x, y);
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
