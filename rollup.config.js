import typescript from "@rollup/plugin-typescript";

const typescriptPlugin = () =>
  typescript({
    tsconfig: "./tsconfig.json",
    compilerOptions: {
      declaration: false,
      declarationMap: false,
      noEmit: false
    }
  });

export default [
  {
    input: {
      index: "src/index.ts",
      core: "src/core/index.ts",
      register: "src/register.ts",
      "clipboard-worker": "src/core/clipboard/worker/serializeWorker.ts"
    },
    output: {
      dir: "dist",
      format: "es",
      sourcemap: true,
      entryFileNames: "[name].js"
    },
    plugins: [typescriptPlugin()]
  },
  {
    // A superset of both src/index.ts (full core + defineCyExcel) and
    // src/register.ts (auto-register side effect) — this one physical file
    // backs the `require` condition for both the "." and "./register"
    // package.json exports. See src/umd-entry.ts's own doc comment.
    input: "src/umd-entry.ts",
    output: {
      file: "dist/cyexcel.umd.cjs",
      format: "umd",
      name: "CyExcel",
      sourcemap: true
    },
    plugins: [typescriptPlugin()]
  }
];
