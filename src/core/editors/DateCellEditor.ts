import { CELL_FONT, CELL_PADDING, SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "../render/theme";
import type { CellEditor, CellEditorContext } from "./CellEditor";
import { positionEditorElement } from "./positionEditorElement";
import type { CellValue } from "../sheet/types";
import { serialFromISODate, serialToISODate } from "../style/dateSerial";

/**
 * 格子里存的是 Excel 序列号（数字），`<input type="date">` 要的是 `yyyy-mm-dd`。
 * 不是序列号的值（比如还没设过的空格）就交给浏览器显示空。
 */
function toDateInputValue(value: CellValue | undefined): string {
  if (typeof value !== "number") {
    return "";
  }
  return serialToISODate(value) ?? "";
}

/**
 * Optional built-in editor for date-shaped values — a native
 * `<input type="date">`, using the browser's own date picker instead of a
 * hand-built calendar widget. Not wired to any CellType automatically; a
 * column opts into it explicitly via `createEditor()`.
 */
export function createDateCellEditor(): CellEditor {
  let input: HTMLInputElement | null = null;

  return {
    mount(container: HTMLElement, context: CellEditorContext): void {
      const element = document.createElement("input");
      element.type = "date";
      element.value = toDateInputValue(context.initialValue);
      element.style.position = "absolute";
      positionEditorElement(element, context.rect);
      element.style.boxSizing = "border-box";
      element.style.margin = "0";
      element.style.padding = `0 ${CELL_PADDING}px`;
      element.style.border = `${SELECTION_STROKE_WIDTH}px solid ${SELECTION_STROKE_COLOR}`;
      element.style.outline = "none";
      element.style.background = "#ffffff";
      element.style.font = CELL_FONT;
      container.appendChild(element);
      input = element;
    },
    updateGeometry(rect): void {
      if (input) {
        positionEditorElement(input, rect);
      }
    },
    focus(): void {
      input?.focus();
    },
    getValue(): CellValue {
      // 写回的也是序列号——日期在这个模型里就是数字 + 日期格式（同 Excel）。
      if (!input?.value) {
        return "";
      }
      // `<input type="date">` 的 value 就是 `yyyy-mm-dd`，直接按墙上时钟换算，
      // 不绕 `Date`（那会按 UTC 解析，东八区往返会多出 0.333 天）。
      return serialFromISODate(input.value) ?? "";
    },
    destroy(): void {
      input?.remove();
      input = null;
    }
  };
}
