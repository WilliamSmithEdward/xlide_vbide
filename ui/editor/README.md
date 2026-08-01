# xlide editor surface

Monaco editor bundled for the xlide WebView2 tool window. VBA syntax, the `xlide-dark` /
`xlide-light` themes, and a typed message bridge to the host.

## Build

```
npm install
node build.mjs      # or: npm run build
npm run typecheck   # tsc --noEmit, strict
npm test            # structural smoke check over dist/
```

`build.mjs` writes everything into `dist/`:

| file               | what it is                                                  |
| ------------------ | ----------------------------------------------------------- |
| `index.html`       | shell, references `./editor.css` and `./editor.js`           |
| `editor.js`        | page bundle, classic script (iife), minified, no sourcemap   |
| `editor.css`       | Monaco css plus the decoration classes, references `./codicon.ttf` |
| `editor.worker.js` | Monaco editor worker, module worker, loaded on demand        |
| `codicon.ttf`      | Monaco icon font                                             |

Every reference is relative, so `dist/` can be dropped behind a WebView2 virtual host mapping,
a static server, or any subdirectory without rewriting. Serve it over http(s) or a virtual
host mapping: the worker is a module worker and will not load from a `file:` origin. The page
itself still runs from `file:`; only the worker-backed extras (link detection, word based
completion) go quiet. `index.html` carries a `default-src 'none'` CSP that permits only
same-origin script, style, font, worker and `data:` images.

## Message contract

Transport is `window.chrome.webview`: `postMessage` out, the `message` event in. Every message
is a JSON object with a `type` discriminator. The types live in `src/bridge.ts`.

**All line and column numbers are 1-based in both directions.** Monaco is 1-based for lines and
1-based for columns (column 1 sits before the first character), so host coordinates map onto
`monaco.IRange` unchanged. An end position is exclusive; a zero-width range is an insertion
point.

`Range` below means `{ startLine, startColumn, endLine, endColumn }`.

### Host to page

| type             | payload                                        | effect                                                                             |
| ---------------- | ---------------------------------------------- | ---------------------------------------------------------------------------------- |
| `loadDocument`   | `moduleName: string`, `text: string`           | Swaps in a model at `xlide:/<moduleName>`, resets revision to 0. No echo.           |
| `applyEdit`      | `revision: number`, `changes: (Range & {text})[]` | Applies the changes in order, then adopts `revision`. No echo.                   |
| `setTheme`       | `theme: "xlide-dark" \| "xlide-light"`         | Sets the theme and pins it, so `prefers-color-scheme` no longer overrides it.       |
| `setDiagnostics` | `markers: (Range & {severity, message, code?})[]` | Replaces all markers for owner `xlide`.                                         |
| `setCurrentLine` | `line: number \| null`                         | Yellow whole-line highlight plus a gutter arrow. `null` clears it.                  |
| `setBreakpoints` | `lines: number[]`                              | Replaces the red dots in the glyph margin.                                          |
| `revealLine`     | `line: number`                                 | Scrolls the line into view if it is outside the viewport.                            |

`severity` is `"error" | "warning" | "info" | "hint"`.

### Page to host

| type                        | payload                                                | when                                                        |
| --------------------------- | ------------------------------------------------------ | ----------------------------------------------------------- |
| `ready`                     | none                                                   | Once, after the editor and the bridge are wired.             |
| `contentChanged`            | `revision: number`, `changes: (Range & {text})[]`, `fullText: string` | After any local edit.                     |
| `selectionChanged`          | `startLine, startColumn, endLine, endColumn`           | On every cursor or selection change.                        |
| `breakpointToggleRequested` | `line: number`                                         | Glyph margin click. The page does not toggle anything itself. |

### Revision and echo suppression

The page keeps a revision counter. It increments on every locally originated edit and the new
value rides along on `contentChanged`. While a host `applyEdit` or `loadDocument` is being
written into the model the counter is frozen and `contentChanged` is not emitted, so a host
edit never echoes back as a user edit. After `applyEdit` the page adopts the host's `revision`:
the host is the authority once it has written to the document. `loadDocument` resets it to 0.

`contentChanged.changes` is passed through in Monaco's order, which is bottom-up so that the
earlier ranges stay valid while the later ones are applied. Apply them in the order given.

## Demo mode

If `window.chrome.webview` is missing the page swaps in a loopback transport instead: it loads
a sample VBA module, sets a marker, two breakpoints and a current line, and logs every message
in both directions to the console as `[xlide demo] page -> host` and `[xlide demo] host -> page`.
Clicking the glyph margin round-trips through the fake host. Serve `dist/` and open it in a
browser to exercise the surface without the add-in.
