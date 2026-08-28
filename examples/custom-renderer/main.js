import { CyExcelCore, rowsToWorkbookSnapshot } from "../../dist/core.js";

/** A CanvasCellRenderer: paints directly on the shared content canvas, and hitTest() turns a click into a CellAction the core translates into a Command. */
const clicksRenderer = {
  kind: "canvas",
  paint({ ctx, rect, value }) {
    ctx.fillStyle = "#0a7a3d";
    ctx.font = "14px -apple-system, BlinkMacSystemFont, sans-serif";
    ctx.textBaseline = "middle";
    ctx.textAlign = "left";
    ctx.fillText(`${value} clicks`, rect.x + 8, rect.y + rect.height / 2);
  },
  hitTest({ value }) {
    return { type: "set-value", value: Number(value) + 1 };
  }
};

/** A DOM renderer: mounts a real element per visible cell (pooled/recycled as you scroll), for anything a canvas paint call can't express — here, a native checkbox-style toggle button. */
const activeRendererFactory = {
  kind: "dom",
  create() {
    let button;
    let currentValue;
    return {
      mount(host, context) {
        button = document.createElement("button");
        button.type = "button";
        button.className = "toggle-button";
        currentValue = context.value;
        button.textContent = currentValue ? "Active" : "Inactive";
        button.addEventListener("click", () => {
          context.actions.commit(!currentValue);
        });
        host.append(button);
      },
      update(context) {
        currentValue = context.value;
        button.textContent = currentValue ? "Active" : "Inactive";
      },
      destroy() {
        button.remove();
      }
    };
  }
};

const columns = [
  { width: 200, header: "Name" },
  { width: 160, header: "Clicks", renderer: clicksRenderer },
  { width: 160, header: "Active", renderer: activeRendererFactory }
];

/** The i-th field is the i-th column — a parameter of the import, not of `columns`. */
const fields = ["name", "clicks", "active"];

const data = Array.from({ length: 100 }, (_unused, index) => ({
  name: `Item ${index}`,
  clicks: 0,
  active: index % 2 === 0
}));

new CyExcelCore(document.getElementById("sheet"), {
  columns,
  snapshot: rowsToWorkbookSnapshot(data, fields),
  rowHeight: 26,
  headerHeight: 28
});
