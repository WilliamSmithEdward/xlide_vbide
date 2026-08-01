// esbuild resolves CSS side-effect imports and folds them into the emitted stylesheet;
// TypeScript needs to be told the module exists.
declare module "*.css";
