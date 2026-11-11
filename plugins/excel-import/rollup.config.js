import typescript from "@rollup/plugin-typescript";

export default {
  input: "src/index.ts",
  output: { dir: "dist", format: "es", sourcemap: true },
  external: ["cyexcel", "cyexcel/toolbar", "exceljs"],
  plugins: [typescript({
    tsconfig: "tsconfig.json",
    compilerOptions: { emitDeclarationOnly: false, declaration: false, declarationMap: false, outDir: "dist" }
  })]
};
