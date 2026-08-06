/**
 * Entry for the single UMD/CJS bundle (`dist/cyexcel.umd.cjs`), which
 * `package.json`'s `exports` map uses as the `require` condition for BOTH
 * `"."` and `"./register"` — it has to be a superset of `src/index.ts`
 * (full core + `defineCyExcel`) and `src/register.ts` (auto-register side
 * effect), which is why it's a separate entry rather than reusing either
 * one directly. `src/register.ts` itself stays untouched — it's still the
 * ESM entry for `dist/register.js`.
 */
import { DEFAULT_CYEXCEL_TAG_NAME, defineCyExcel } from "./element/defineCyExcel";

export * from "./core";
export { DEFAULT_CYEXCEL_TAG_NAME, defineCyExcel };

if (typeof globalThis.customElements !== "undefined") {
  defineCyExcel();
}
