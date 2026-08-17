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
