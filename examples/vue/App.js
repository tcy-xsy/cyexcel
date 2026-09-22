import { createApp, defineComponent, h, onMounted, onUnmounted, ref } from "vue";
import { CyExcelCore, matrixToWorkbookSnapshot } from "../../dist/core.js";

// No explicit headers means Excel-style A/B/C column letters.
const data = [
  ["A100", "Widget", 9.99],
  ["A101", "Gadget", 14.5],
  ["A102", "Doohickey", 3.25]
];

const CyExcelSheet = defineComponent({
  setup() {
    const host = ref(null);
    let sheet;

    onMounted(() => {
      // Imported once, on mount. The snapshot is a value handed to the sheet,
      // not a reactive binding — nothing flows back into `data` either way.
      sheet = new CyExcelCore(host.value, {
        snapshot: matrixToWorkbookSnapshot(data)
      });
    });
    onUnmounted(() => {
      sheet.dispose();
    });

    return () => h("div", { class: "sheet-host", ref: host });
  }
});

createApp(CyExcelSheet).mount("#app");
