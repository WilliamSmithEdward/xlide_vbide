/**
 * The build's own timestamp, defined by esbuild at bundle time (engine/build.mjs) and answered
 * by `initialize`. Declared here so `tsc --noEmit` - which runs without esbuild's defines -
 * knows the name; only the bundle ever runs.
 */
declare const __ENGINE_BUILT__: string;
