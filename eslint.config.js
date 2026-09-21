import js from "@eslint/js";
import tseslint from "typescript-eslint";
import globals from "globals";

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
    //
    // `toViewportRect`/`getSelectionBottomRightRect` no longer exist — Phase 8
    // deleted them as dead compatibility adapters. They stay listed on purpose:
    // this list records what must never come back, not merely what exists.
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
    // Migration plan §11 (Phase 8): CyExcelCore must not interpret the model's
    // Patch vocabulary. The original condition was worded as "#invalidatePatches()
    // 已删除", which is not checkable — the method's remaining content is
    // post-commit composition, which §12 explicitly allows Core to hold. What is
    // actually required is that Core stops *reading* changes: dirty derivation
    // belongs to compileDirtySet(), axis facts to summarizeStructuralChanges(),
    // measurement to WrapMeasurementService. Importing `Patch` is the tripwire
    // for that, and unlike a method name it cannot be satisfied by renaming.
    files: ["src/core/CyExcelCore.ts"],
    rules: {
      "no-restricted-imports": [
        "error",
        {
          paths: [
            {
              name: "./types",
              importNames: ["Patch"],
              message:
                "CyExcelCore composes a commit receipt; it does not interpret changes. Consume ChangeSet/SemanticChange, and put any per-change logic in the owning service (ADR 0002, plan §11)."
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
    // Plain browser scripts (loaded via <script type="module">, not built or
    // type-checked). 这里**声明浏览器全局**而不是关掉 no-undef：关掉的话
    // `document` 不报了，但 `undoButton is not defined` 这类真错误也一起不报，
    // 只能等到运行时在页面上炸——Step 25 删 playground 旧控件时就是这么漏掉两处的。
    files: ["playground/**/*.js", "examples/**/*.js"],
    languageOptions: {
      globals: globals.browser
    }
  },
  {
    // Plain Node scripts. 同上，声明 Node 全局而不是关规则。
    files: ["scripts/**/*.mjs"],
    languageOptions: {
      globals: globals.node
    }
  }
);
