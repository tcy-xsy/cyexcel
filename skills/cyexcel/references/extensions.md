# Extensions and performance

## Styles and images

```ts
sheet.setRangeStyle(
  { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
  { bold: true, color: "#14532d", background: "#dcfce7" }
);
sheet.setCellImage(0, 0, "https://example.com/image.png");
sheet.setCellImage(0, 0, undefined); // Remove the image, retaining value and style.
```

Style patches replace only supplied keys; `null` clears a key. A per-cell image takes rendering priority without needing a column renderer. Use `createImageCellRenderer` from `cyexcel/core` for an entire column of image URLs. Complete URL/email cell values are recognized as links; embedded URLs in prose are not.

## Plugins

Register plugins at construction using `plugins`; runtime installation is not part of the public API.

```ts
import type { CyExcelPlugin } from "cyexcel/core";

const auditPlugin: CyExcelPlugin = {
  name: "audit",
  setup(context) {
    context.events.on("change", () => {
      console.log("Sheet changed");
    });
    context.registerShortcut("Mod+Shift+L", () => {
      console.log("Shortcut activated");
    });
  }
};
```

Pass the plugin through `createSheet(container, { plugins: [auditPlugin], ...options })`. Use the controlled `PluginContext` for events, hooks, clipboard, menu providers, and structural commands rather than accessing internal stores. Registered capabilities are cleaned up automatically. Return a disposer for additional resources created by the plugin. Built-in hooks are synchronous; do not use `tapAsync()` for them.

Menu providers register through `context.contextMenu.registerProvider(meta, provider)`. Inspect the installed `ContextMenuProvider`, `ContextMenuItem`, and `PluginContext` types before writing a custom provider. Action items support disabled/checked states, shortcuts, descriptions, and asynchronous execution; submenus use children. Optional `sheetMenu` exposes `getState()` and `run()` for supported spreadsheet commands. Do not call runtime capabilities during plugin setup before activation.

## Performance decisions

Rendering and hit testing are virtualized; snapshot import and storage are not. Budget import time and memory independently from scroll responsiveness. Use representative row counts, cell contents, styles, images, and formula dependencies when measuring. Avoid full rectangular scans of very large sparse sheets, rebuilding the instance for every framework update, or expensive custom validators during paint.

The documentation performance example uses 100,000 rows and five columns. It is a reproducible example, not a guaranteed capacity or frame-rate claim. Check behavior on the target browser and device, including import latency, scrolling, edits, and undo. Large clipboard operations can be bounded or refused; do not remove limits simply to make a demo pass.
