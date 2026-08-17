import type { CellRect } from "../layout";

export function positionEditorElement(
  element: HTMLElement,
  rect: CellRect
): void {
  element.style.left = `${rect.x}px`;
  element.style.top = `${rect.y}px`;
  element.style.width = `${rect.width}px`;
  element.style.height = `${rect.height}px`;
}
