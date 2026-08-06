import { createApp, defineComponent, h, onMounted, onUnmounted, ref } from "vue";
import { CyExcelCore } from "../../dist/core.js";

const columns = [
  { id: "sku", field: "sku", width: 140, header: "SKU" },
  { id: "name", field: "name", width: 240, header: "Name" },
  { id: "price", field: "price", width: 100, type: "number", header: "Price" }
];

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
      sheet = new CyExcelCore(host.value, { columns, data });
    });
    onUnmounted(() => {
      sheet.dispose();
    });

    return () => h("div", { class: "sheet-host", ref: host });
  }
});

createApp(CyExcelSheet).mount("#app");
