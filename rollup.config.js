import typescript from "@rollup/plugin-typescript";
import nodeResolve from "@rollup/plugin-node-resolve";

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
      toolbar: "src/toolbar.ts",
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
    // UMD 必须把运行时依赖（numfmt）打进产物：external 在 UMD 里意味着"使用者得
    // 先在全局挂一个 numfmt"，一个 <script> 标签就能用的产物不该有这种前置条件。
    // ESM 那一路相反，保持 external 交给使用者的打包器去解析（见上面那个配置项）。
    plugins: [nodeResolve(), typescriptPlugin()]
  }
];
