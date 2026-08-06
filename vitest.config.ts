import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    // tests/build/**/*.test.ts depends on a prior `pnpm build` (it imports
    // from dist/), so it's excluded from plain `vitest run` and only run
    // via the separate `pnpm test:build` script (`vitest run tests/build`,
    // passed as an explicit path filter) — never as part of `test:unit`.
    include: ["tests/unit/**/*.test.ts", "tests/build/**/*.test.ts"],
    coverage: {
      provider: "v8",
      include: ["src/core/**/*.ts"],
      // Whole-file DOM glue with no branchy logic to unit-test in Node
      // (document.createElement/appendChild, requestAnimationFrame, no DOM
      // shim in this environment) — covered by tests/browser/full-loop.spec.ts,
      // smoke.spec.ts, and lifecycle.spec.ts instead. CyExcelCore.ts is the
      // integration layer itself (Step 7): it only wires together modules
      // that are each independently unit-tested, so this exclusion doesn't
      // hide untested logic, just untestable-in-Node DOM orchestration.
      exclude: ["src/core/editors/TextCellEditor.ts", "src/core/CyExcelCore.ts"],
      reporter: ["text", "html"],
      thresholds: {
        lines: 90,
        functions: 90,
        statements: 90,
        branches: 85
      }
    }
  }
});
