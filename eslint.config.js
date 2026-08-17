import js from "@eslint/js";
import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    ignores: [
      "coverage/**",
      "dist/**",
      "node_modules/**",
      "playwright-report/**",
      "test-results/**"
    ]
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ["**/*.ts"],
    rules: {
      "no-undef": "off",
      "@typescript-eslint/consistent-type-imports": "error",
      "@typescript-eslint/no-import-type-side-effects": "error",
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_", destructuredArrayIgnorePattern: "^_" }
      ]
    }
  },
  {
    files: ["src/core/commands/**/*.ts"],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          paths: [
            {
              name: "../data",
              message: "Command handlers consume narrow mutation capabilities; concrete owners belong in integration."
            },
            {
              name: "../data/index",
              message: "Command handlers consume narrow mutation capabilities; concrete owners belong in integration."
            },
            {
              name: "../data/TableStore",
              message: "Command handlers consume narrow mutation capabilities; concrete owners belong in integration."
            }
          ],
          patterns: [
            {
              group: [
                "../layout",
                "../layout/*",
                "../integration",
                "../integration/*",
                "../render",
                "../render/*",
                "../editors",
                "../editors/*"
              ],
              message: "Command handlers consume narrow mutation capabilities; concrete owners belong in integration."
            }
          ]
        }
      ]
    }
  },
  {
    // Migration plan §12: renderer / overlay / interaction consume the Scene
    // snapshot, never the mutable geometry owners. Pure value types (CellRect,
    // IndexRange) and pure index math (remapRowIndex) stay allowed — they carry
    // no coordinate authority. What is banned is anything that can re-derive a
    // second coordinate truth: the LayoutEngine itself, the scroll-subtraction
    // adapter, and the mutable ViewportState shape.
    files: [
      "src/core/render/**/*.ts",
      "src/core/interaction/**/*.ts",
      "src/core/cells/**/*.ts",
      "src/core/editors/**/*.ts",
      "src/core/headers/**/*.ts",
      "src/core/menu/**/*.ts",
      "src/core/scene/**/*.ts",
      "src/core/clipboard/**/*.ts",
      "src/core/measurement/**/*.ts"
    ],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          paths: [
            {
              name: "../layout",
              importNames: [
                "LayoutEngine",
                "toViewportRect",
                "getSelectionBottomRightRect",
                "ViewportState"
              ],
              message:
                "Consume the Scene snapshot for geometry. LayoutEngine/toViewportRect/ViewportState would re-create a second coordinate truth (ADR 0002)."
            }
          ],
          patterns: [
            {
              group: [
                "../layout/LayoutEngine",
                "../layout/viewportRect",
                "../layout/selectionGeometry",
                "../../layout/LayoutEngine",
                "../../layout/viewportRect",
                "../../layout/selectionGeometry"
              ],
              message:
                "Consume the Scene snapshot for geometry. LayoutEngine/toViewportRect/selectionGeometry would re-create a second coordinate truth (ADR 0002)."
            }
          ]
        }
      ]
    }
  },
  {
    // Migration plan §12: History stores receipts, not Command instances.
    files: ["src/core/history/**/*.ts"],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          patterns: [
            {
              group: ["../commands", "../commands/*"],
              message:
                "History admits CommitReceipt only. Importing a concrete Command class puts undo state back on the command (ADR 0002)."
            }
          ]
        }
      ]
    }
  },
  {
    // Migration plan §12: DOM scroll offsets and the devicePixelRatio global
    // each have exactly one owner. The whitelist below has now shrunk to those
    // owners alone — `CyExcelCore.ts` came off it once ViewportManager took
    // over the scroller and the spacer (Phase 7).
    files: ["src/**/*.ts"],
    ignores: [
      "src/core/scene/**/*.ts",
      "src/core/layout/**/*.ts",
      "src/core/render/dprCanvas.ts"
    ],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector: "MemberExpression[property.name=/^scroll(Top|Left)$/]",
          message:
            "Scroll offsets belong to the viewport owner. Read them from the Scene pane's scroll fact instead (ADR 0002)."
        },
        {
          selector:
            "MemberExpression[object.name=/^(window|globalThis|self)$/][property.name='devicePixelRatio']",
          message:
            "Read the dpr global through readDevicePixelRatio() in render/dprCanvas.ts; everything else takes dpr as a parameter."
        }
      ]
    }
  },
  {
    // Plain browser script (loaded via <script type="module">, not built or
    // type-checked), so DOM/browser globals need the same no-undef
    // exemption the .ts block gets from the TypeScript compiler instead.
    files: ["playground/**/*.js", "examples/**/*.js"],
    rules: {
      "no-undef": "off"
    }
  },
  {
    // Plain Node script, same story as the browser-globals block above but
    // for Node globals (process/console/URL) instead of DOM ones.
    files: ["scripts/**/*.mjs"],
    rules: {
      "no-undef": "off"
    }
  }
);
