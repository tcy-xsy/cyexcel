import { createApp, defineComponent, h, onMounted, onUnmounted, ref } from "vue";
import { CyExcelCore, rowsToWorkbookSnapshot } from "../../dist/core.js";

const columns = [
  { width: 140, header: "SKU" },
  { width: 240, header: "Name" },
  { width: 100, type: "number", header: "Price" }
];

/** The i-th field is the i-th column — a parameter of the import, not of `columns`. */
const fields = ["sku", "name", "price"];

const data = [
  { sku: "A100", name: "Widget", price: 9.99 },
  { sku: "A101", name: "Gadget", price: 14.5 },
  { sku: "A102", name: "Doohickey", price: 3.25 }
];

const CyExcelSheet = defineComponent({
  setup() {
    const host = ref(null);
    let sheet;

    onMounted(() => {
      // Imported once, on mount. The snapshot is a value handed to the sheet,
      // not a reactive binding — nothing flows back into `data` either way.
      sheet = new CyExcelCore(host.value, {
        columns,
        snapshot: rowsToWorkbookSnapshot(data, fields)
      });
    });
    onUnmounted(() => {
      sheet.dispose();
    });

    return () => h("div", { class: "sheet-host", ref: host });
  }
});

createApp(CyExcelSheet).mount("#app");
