import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const scratch = await mkdtemp(path.join(tmpdir(), 'xlide-editor-security-'));
try {
  const outfile = path.join(scratch, 'patterns.mjs');
  await build({
    stdin: {
      contents: `export { vbaLanguageConfiguration } from './src/vba.ts';
        import { registerMarkupLanguage } from './src/formmarkuplang.ts';
        registerMarkupLanguage();
        export { configurations } from 'monaco-editor/editor/editor.api.js';`,
      resolveDir: root,
    },
    outfile, bundle: true, platform: 'node', format: 'esm', logLevel: 'silent',
    alias: { 'xlide-spec': path.join(root, 'vendor/xlide-spec') },
    plugins: [{ name: 'capture-monaco', setup(plugin) {
      plugin.onResolve({ filter: /^monaco-editor\// }, args => ({ path: args.path, namespace: 'stub' }));
      plugin.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: `
        export const configurations = {};
        const stub = new Proxy(function(){}, {get:()=>stub, apply:()=>stub});
        export const languages = new Proxy({setLanguageConfiguration:(id,value)=>configurations[id]=value}, {get:(obj,key)=>obj[key]??stub});
      ` }));
    } }],
  });
  const { vbaLanguageConfiguration, configurations } = await import(pathToFileURL(outfile));
  const indent = vbaLanguageConfiguration.indentationRules.increaseIndentPattern;
  for (const name of ['i', 'élément', '変数', 'e\u0301']) {
    assert.ok(indent.test(`For ${name} = 1 To 10`), name);
    assert.ok(indent.test(`For Each ${name} In values`), name);
  }
  const form = configurations['xlide-form'].onEnterRules[0].beforeText;
  assert.ok(form.test('<Frame Name="Main" Caption="a > b">'));
  assert.ok(!form.test('<Frame Name="Main" />'));
  assert.ok(!form.test('</Frame>'));
  const start = performance.now();
  assert.ok(!form.test('<Form ' + '"'.repeat(10000) + 'x'));
  assert.ok(performance.now() - start < 1000, 'malformed markup must not stall the editor');
  console.log('Unicode indentation and malformed markup security regressions passed');
} finally {
  await rm(scratch, { recursive: true, force: true });
}
