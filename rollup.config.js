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
      register: "src/register.ts"
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
    input: "src/register.ts",
    output: {
      file: "dist/cyexcel.umd.cjs",
      format: "umd",
      name: "CyExcel",
      sourcemap: true
    },
    plugins: [typescriptPlugin()]
  }
];
