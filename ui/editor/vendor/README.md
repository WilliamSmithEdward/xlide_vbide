# Vendored sources

Code in here is not ours. It is copied verbatim from another repository so that this one builds on
its own, and it should never be edited in place — an edit here is a change that the repo it came
from knows nothing about.

## xlide-spec

The smart-editing helpers the page bundles from the spec repo (`xlide_vscode`): Smart Enter, Smart
Tab, and the lexer they lean on. Typing in the VBE surface and typing in the VS Code extension run
the same code, and that is the point of copying it rather than reimplementing it.

`xlide-spec.json` records where the copy came from — the commit of the spec repo it was taken at,
the entry points the page imports, and a hash per file.

The arrangement has two halves, and they resolve the same `xlide-spec/*` specifier to different
places on purpose:

- **Behaviour** comes from here. `build.mjs` aliases the specifier to this directory, so what runs
  is the spec's real implementation.
- **Types** come from `src/spec/xlide-spec`, where hand-written declarations describe the same API.
  The spec compiles under its own compiler settings, which are looser than this project's, so
  type-checking its sources here would fail on rules the spec never agreed to.

### Keeping it honest

```bash
npm run spec:check
```

Compares the copy against a neighbouring `xlide_vscode` checkout and fails if the two have parted
ways, if the page has started importing something that was never vendored, or if a file here was
edited by hand. Without a neighbouring checkout — on CI, or in a clone of this repo alone — it
verifies the copy against its manifest and reports that it could not do more. `tools/verify.ps1`
runs it, so drift fails on a machine that has both repos rather than going unnoticed.

```bash
npm run spec:sync
```

Re-copies from the neighbouring checkout and rewrites the manifest. Run it after the spec repo
changes something the page depends on, and commit the result.

### What this does not check

The declarations in `src/spec/xlide-spec` are hand-written and are not compared against anything. If
the spec changes a signature, the copy here carries the new behaviour while the declarations still
describe the old shape, and the type-check will not notice. That is a pre-existing gap in the
arrangement, not something vendoring introduced, but it is worth knowing about when a sync pulls in
more than an implementation detail.
