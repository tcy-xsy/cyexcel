import type { CellPaintContext, CellPainter } from "../render";
import { CELL_FONT, CELL_PADDING, TEXT_COLOR } from "../render/theme";
import type { CellRenderContext, RuntimeColumn } from "./types";

function drawFormattedText(context: CellPaintContext, column: RuntimeColumn): void {
  const { ctx, rect, row, value } = context;
  const text = column.format(value, { row, column });
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = CELL_FONT;
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  ctx.fillText(text, rect.x + CELL_PADDING, rect.y + rect.height / 2);
}

/**
 * Bridges render's column-agnostic CellPainter extension point
 * (ContentLayerRendererOptions.painter) to Step 8's per-column renderer
 * resolution, without any changes to ContentLayerRenderer itself — errors
 * thrown by a custom CanvasCellRenderer.paint() are left to propagate so
 * ContentLayerRenderer's existing try/catch + onError + default-painter
 * fallback (already tested) handles them, exactly as it does for any other
 * custom `painter`.
 */
export function createColumnAwarePainter(columns: readonly RuntimeColumn[]): CellPainter {
  return (context: CellPaintContext) => {
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
