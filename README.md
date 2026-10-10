# Office File Viewer (docx · xlsx · pptx)

A VS Code extension that renders Microsoft Office Open XML files **inside the editor** — no Office
install, no server-side conversion, no export step.

Rendering is powered by [`@silurus/ooxml`](https://ooxml.silurus.dev/), an open-source
Rust/WebAssembly + Canvas 2D engine that parses and paints `.docx`, `.xlsx` and `.pptx` entirely
locally. Files never leave your machine.

## Features

- **Word (.docx / .docm)** — paginated layout, toolbar page navigation, zoom, fit-width and
  full-document find with highlight.
- **PowerPoint (.pptx / .pptm)** — slide navigation, zoom, fit-width, full-deck find.
- **Excel (.xlsx / .xlsm)** — the engine's full workbook UI: sheet tabs, frozen panes, range
  selection with TSV copy, zoom slider, cached formulas, number formats and charts.
- Hyperlinks inside documents open through VS Code (`http`, `https`, `mailto`, `tel` only).
- The view live-refreshes when the file changes on disk (re-run your generator, the viewer updates).
- Theme-aware UI (light/dark/high-contrast), keyboard shortcuts, works in Remote/WSR windows
  (files are read via the VS Code file system API).

| Keys (docx / pptx) | Action |
| --- | --- |
| `←` `→` `↑` `↓` `PgUp` `PgDn` | Previous / next page or slide |
| `Home` / `End` | First / last page or slide |
| `Ctrl`+= / `Ctrl`+- / `Ctrl`+0 | Zoom in / out / reset |
| `Ctrl`+F, `Enter` / `Shift`+`Enter`, `Esc` | Find, next / previous match, close |

The viewers are **read-only**; use *Open With… → Text Editor* as a fallback escape hatch.

## Settings

| Setting | Default | Description |
| --- | --- | --- |
| `officeFileViewer.useGoogleFonts` | `false` | Fetch metric-compatible webfonts and non-Latin script fallbacks (CJK, Arabic, Hebrew, …) from Google Fonts *while rendering*, so layout matches Office more closely. When off, only locally installed system fonts are used. |
| `officeFileViewer.enableTextSelection` | `true` | Overlay a transparent text layer on document pages and slides for native selection & copy. |
| `officeFileViewer.xlsxShowZoomSlider` | `true` | Show the built-in zoom slider in spreadsheet viewers. |

Settings are applied when an editor is (re)opened.

## Offline by default

With the default settings the extension makes **no network requests** while rendering:

- The WASM parsers, webview scripts and all engine code ship inside the extension and are served
  from its own resources; the webview Content-Security-Policy allows only the extension's local
  resource origin (`data:`/`blob:` workers for the sandboxed parser worker) — no internet host is
  permitted.
- Text is laid out with locally installed system fonts (no webfont CDN). Non-Latin text uses your
  system's font fallbacks.
- Optional chart engines (MathJax equations, 3-D charts, ChartEx, region maps) and the TIFF codec
  are deliberately **not bundled**; documents using them render placeholders instead of fetching
  renderer code.
- The only user-triggered network actions are: clicking a hyperlink *inside* a document (opens the
  browser via VS Code, `http`/`https`/`mailto`/`tel`), and enabling `officeFileViewer.useGoogleFonts`
  (then the renderer fetches fonts from `fonts.googleapis.com` / `fonts.gstatic.com` only).

## How it works

- Each format is registered as a VS Code **custom read-only editor** (modeled on the
  [custom-editor-sample](https://github.com/microsoft/vscode-extension-samples/tree/main/custom-editor-sample)).
- The extension host reads the file bytes (`vscode.workspace.fs`) and posts them to the webview;
  the bundled `@silurus/ooxml` engine parses (WASM) and renders (Canvas 2D) there.
- The parser WASM ships beside the webview bundle (`out/media/assets/*.wasm`) and is resolved via
  the webview `<base>`; the Content-Security-Policy only ever allows the extension's own resources
  (plus Google Fonts when you opt in).

## Development

```bash
npm install
npm run compile      # esbuild: extension + 3 webview bundles + wasm assets
npm run typecheck    # tsc for extension (node) and webview (dom) code
npm run watch
npm test             # integration test: renders sample.docx/xlsx/pptx in a real VS Code
```

`npm test` launches a disposable VS Code instance (`@vscode/test-electron`), opens each sample with
the matching custom editor and waits for the webview to report a successful render. Lifecycle events
are also visible at runtime in the **Office File Viewer** output channel (`resolve` / `rendered` /
`render FAILED` / hyperlinks).

Open the project in VS Code and press **F5** ("Run Office File Viewer") — a development host opens
the `samples/` workspace with `sample.docx`, `sample.xlsx` and `sample.pptx` to try.

Package a VSIX with `npx @vscode/vsce package` (requires `npm run package` first, which is wired as
`vscode:prepublish`).

## Limitations

- Rendering fidelity is defined by the engine; complex documents may differ from Office. See the
  [feature support matrix](https://github.com/yukiyokotani/office-open-xml-viewer#feature-support).
- Large files are loaded fully into memory (bytes are transferred to the webview).
- Rendering runs on the webview's main thread (`mode: 'main'`); the engine's worker render mode is
  not used.

## Credits

- [`@silurus/ooxml`](https://ooxml.silurus.dev/) by Yuki Yokotani — the rendering engine (MIT).
- [`microsoft/vscode-extension-samples`](https://github.com/microsoft/vscode-extension-samples) —
  the custom editor template this extension is modeled on.
