import * as vscode from 'vscode';
import { getViewerConfig, type ViewerConfig } from './config';
import { getNonce } from './util';

export type OfficeFormat = 'docx' | 'xlsx' | 'pptx';

/** Read-only document model: the file's raw bytes plus its format. */export class OfficeDocument implements vscode.CustomDocument {
	constructor(
		readonly uri: vscode.Uri,
		readonly format: OfficeFormat,
		/** Full file content; transferred to the webview as an ArrayBuffer. */
		public data: Uint8Array,
	) {}

	private readonly _onDidDispose = new vscode.EventEmitter<void>();
	readonly onDidDispose = this._onDidDispose.event;

	dispose(): void {
		this._onDidDispose.fire();
		this._onDidDispose.dispose();
	}
}

/**
 * Read-only custom editor that renders OOXML files (docx/xlsx/pptx) in a
 * webview using the @silurus/ooxml WASM/Canvas engine.
 */
export class OfficeEditorProvider implements vscode.CustomReadonlyEditorProvider<OfficeDocument> {
	public static register(
		context: vscode.ExtensionContext,
		format: OfficeFormat,
	): vscode.Disposable {
		const provider = new OfficeEditorProvider(context, format);
		return vscode.window.registerCustomEditorProvider(
			`${OfficeEditorProvider.VIEW_TYPE_PREFIX}.${format}`,
			provider,
			{
				supportsMultipleEditorsPerDocument: false,
			},
		);
	}

	private static readonly VIEW_TYPE_PREFIX = 'officeFileViewer';

	/** Shared diagnostics channel; also mirrored to the extension host log. */
	private static readonly output = vscode.window.createOutputChannel('Office File Viewer');

	/** Ring buffer of lifecycle events, queryable via the diagnostics command. */
	private static readonly events: string[] = [];

	private static log(message: string): void {
		const line = `[${new Date().toISOString()}] ${message}`;
		OfficeEditorProvider.output.appendLine(line);
		OfficeEditorProvider.events.push(line);
		if (OfficeEditorProvider.events.length > 200) OfficeEditorProvider.events.shift();
		console.log(`[office-file-viewer] ${message}`);
	}

	/** Most recent lifecycle events (newest last); used by tests & debugging. */
	public static getDiagnostics(): string[] {
		return [...OfficeEditorProvider.events];
	}

	constructor(
		private readonly context: vscode.ExtensionContext,
		private readonly format: OfficeFormat,
	) {}

	async openCustomDocument(
		uri: vscode.Uri,
		_openContext: vscode.CustomDocumentOpenContext,
		_token: vscode.CancellationToken,
	): Promise<OfficeDocument> {
		const data = await vscode.workspace.fs.readFile(uri);
		return new OfficeDocument(uri, this.format, data);
	}

	async resolveCustomEditor(
		document: OfficeDocument,
		webviewPanel: vscode.WebviewPanel,
		_token: vscode.CancellationToken,
	): Promise<void> {
		OfficeEditorProvider.log(`resolve ${document.format}: ${document.uri.fsPath}`);
		webviewPanel.webview.options = {
			enableScripts: true,
			localResourceRoots: [this.context.extensionUri],
		};
		webviewPanel.webview.html = this.getHtmlForWebview(webviewPanel.webview, document.format);

		// Push the file bytes whenever the webview asks for them, so the same
		// path also serves reloads after the file changes on disk.
		const sendInit = async () => {
			let data: Uint8Array;
			try {
				data = await vscode.workspace.fs.readFile(document.uri);
				document.data = data;
			} catch {
				data = document.data;
			}
			// postMessage structured-clones the payload; slice the exact buffer.
			const bytes = data.buffer.slice(
				data.byteOffset,
				data.byteOffset + data.byteLength,
			) as ArrayBuffer;
			await webviewPanel.webview.postMessage({
				type: 'init',
				format: document.format,
				bytes,
				config: { ...getViewerConfig(), wasmUrl: this.getWasmUri(webviewPanel.webview, document.format) },
				fileName: vscode.workspace.asRelativePath(document.uri, false),
			});
		};

		const messageSubscription = webviewPanel.webview.onDidReceiveMessage((message) => {
			switch (message?.type) {
				case 'ready':
					void sendInit();
					return;
				case 'loaded':
					OfficeEditorProvider.log(`rendered ${message.fileName ?? document.uri.fsPath}`);
					return;
				case 'load-error':
					OfficeEditorProvider.log(
						`render FAILED for ${message.fileName ?? document.uri.fsPath}: ${message.message ?? 'unknown error'}`,
					);
					return;
				case 'link':
					OfficeEditorProvider.log(`hyperlink requested: ${message.url}`);
					void this.openExternalLink(message.url);
					return;
			}
		});

		// Refresh the view when the underlying file changes (e.g. regenerating
		// a report rewrites the workbook while the viewer is open).
		const watcher = vscode.workspace.createFileSystemWatcher(
			new vscode.RelativePattern(document.uri, '*'),
		);
		const reloadSubscription = watcher.onDidChange(() => {
			void sendInit();
		});

		webviewPanel.onDidDispose(() => {
			messageSubscription.dispose();
			reloadSubscription.dispose();
			watcher.dispose();
		});
	}

	async backupCustomDocument(
		document: OfficeDocument,
		_context: vscode.CustomDocumentBackupContext,
		_token: vscode.CancellationToken,
	): Promise<vscode.CustomDocumentBackup> {
		// The document is read-only; there is nothing to restore on hot exit.
		return {
			id: document.uri.toString(),
			delete(): void {
				// Nothing to clean up.
			},
		};
	}

	/**
	 * Open an external hyperlink from the document. Only well-known safe
	 * schemes are honored; anything else is ignored.
	 */
	private async openExternalLink(rawUrl: unknown): Promise<void> {
		if (typeof rawUrl !== 'string') return;
		let uri: vscode.Uri;
		try {
			uri = vscode.Uri.parse(rawUrl);
		} catch {
			return;
		}
		const allowed = new Set(['http', 'https', 'mailto', 'tel']);
		if (!allowed.has(uri.scheme.toLowerCase())) return;
		await vscode.env.openExternal(uri);
	}

	/**
	 * Absolute webview URI of the format's parser WASM asset. The engine
	 * resolves its bundled relative asset URL against the webview document
	 * location (which ignores <base>), so we pass the URL explicitly.
	 */
	private getWasmUri(webview: vscode.Webview, format: OfficeFormat): string {
		return webview
			.asWebviewUri(
				vscode.Uri.joinPath(
					this.context.extensionUri,
					'out',
					'media',
					'assets',
					`${format}_parser_bg.wasm`,
				),
			)
			.toString();
	}

	private getHtmlForWebview(webview: vscode.Webview, format: OfficeFormat): string {
		const scriptUri = webview.asWebviewUri(
			vscode.Uri.joinPath(this.context.extensionUri, 'out', 'media', `webview-${format}.js`),
		);
		// esbuild emits the parser WASM as a relative asset URL next to the
		// bundle. The webview document has its own origin, so anchor the
		// document base at the script's directory to resolve those URLs.
		const assetBaseUri = new URL('.', scriptUri.toString()).toString();

		const config = getViewerConfig();
		const nonce = getNonce();
		const csp = buildContentSecurityPolicy(webview.cspSource, nonce, config.useGoogleFonts);
		const paged = format !== 'xlsx';

		return /* html */ `<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<meta http-equiv="Content-Security-Policy" content="${csp}" />
	<base href="${assetBaseUri}" />
	<title>Office Viewer</title>
	<style>
		* { box-sizing: border-box; margin: 0; padding: 0; }
		html, body {
			width: 100%;
			height: 100%;
			overflow: hidden;
			background: var(--vscode-editor-background);
			color: var(--vscode-foreground);
			font-family: var(--vscode-font-family, sans-serif);
		}
		#viewer-root {
			position: relative;
			width: 100%;
			height: 100%;
			min-height: 0;
			display: flex;
			flex-direction: column;
		}
		#viewer-toolbar {
			height: 35px;
			flex: 0 0 35px;
			display: flex;
			align-items: center;
			justify-content: flex-end;
			gap: 4px;
			padding: 0 8px;
			border-bottom: 1px solid var(--vscode-panel-border, rgba(127, 127, 127, 0.35));
			background: var(--vscode-editorGroupHeader-tabsBackground, var(--vscode-editor-background));
			user-select: none;
		}
		body[data-toolbar="none"] #viewer-toolbar { display: none; }
		#viewer-toolbar button {
			min-width: 26px;
			height: 24px;
			padding: 0 6px;
			border: 1px solid transparent;
			border-radius: 3px;
			color: var(--vscode-foreground);
			background: transparent;
			font: 14px/1 var(--vscode-font-family, sans-serif);
			cursor: pointer;
		}
		#viewer-toolbar button:hover:not(:disabled) {
			background: var(--vscode-toolbar-hoverBackground, rgba(127, 127, 127, 0.2));
		}
		#viewer-toolbar button:focus-visible {
			outline: 1px solid var(--vscode-focusBorder);
			outline-offset: -1px;
		}
		#viewer-toolbar button:disabled {
			cursor: default;
			opacity: 0.4;
		}
		.toolbar-separator {
			width: 1px;
			height: 18px;
			margin: 0 4px;
			background: var(--vscode-panel-border, rgba(127, 127, 127, 0.35));
		}
		#position-label {
			min-width: 90px;
			text-align: center;
			font-size: 12px;
			font-variant-numeric: tabular-nums;
			white-space: nowrap;
		}
		#zoom-label {
			min-width: 46px;
			text-align: center;
			font-size: 12px;
			font-variant-numeric: tabular-nums;
		}
		#viewer-container {
			position: relative;
			width: 100%;
			min-width: 0;
			min-height: 0;
			flex: 1 1 auto;
			overflow: hidden;
		}
		#canvas-host {
			position: absolute;
			inset: 0;
			overflow: auto;
			display: flex;
			padding: 16px;
		}
		#canvas-host > canvas {
			margin: auto;
			box-shadow: 0 1px 6px rgba(0, 0, 0, 0.35);
			background: #ffffff;
		}
		body[data-toolbar="none"] #canvas-host { padding: 0; }
		#find-popup {
			position: absolute;
			z-index: 20;
			top: 43px;
			right: 12px;
			display: flex;
			align-items: center;
			gap: 4px;
			padding: 6px;
			border: 1px solid var(--vscode-widget-border, var(--vscode-panel-border));
			border-radius: 4px;
			background: var(--vscode-editorWidget-background, var(--vscode-editor-background));
			color: var(--vscode-editorWidget-foreground, var(--vscode-foreground));
			box-shadow: 0 2px 8px var(--vscode-widget-shadow, rgba(0, 0, 0, 0.35));
		}
		#find-popup[hidden] { display: none; }
		#find-input {
			width: min(240px, 42vw);
			height: 24px;
			padding: 2px 6px;
			border: 1px solid var(--vscode-input-border, transparent);
			outline: none;
			background: var(--vscode-input-background);
			color: var(--vscode-input-foreground);
			font: 12px/1.4 var(--vscode-font-family, sans-serif);
		}
		#find-input:focus { border-color: var(--vscode-focusBorder); }
		#find-status {
			min-width: 70px;
			text-align: center;
			white-space: nowrap;
			font-size: 12px;
			font-variant-numeric: tabular-nums;
		}
		#find-popup button {
			width: 24px;
			height: 24px;
			border: 1px solid transparent;
			border-radius: 3px;
			color: inherit;
			background: transparent;
			font: 14px/1 var(--vscode-font-family, sans-serif);
			cursor: pointer;
		}
		#find-popup button:hover:not(:disabled) {
			background: var(--vscode-toolbar-hoverBackground, rgba(127, 127, 127, 0.2));
		}
		#find-popup button:disabled {
			cursor: default;
			opacity: 0.4;
		}
		#status {
			position: absolute;
			inset: 0;
			display: flex;
			align-items: center;
			justify-content: center;
			pointer-events: none;
			z-index: 10;
		}
		#status[data-state="idle"] { display: none; }
		#status[data-state="error"] {
			pointer-events: auto;
			color: var(--vscode-errorForeground, #f44747);
			font-size: 13px;
			padding: 16px;
			text-align: center;
			white-space: pre-wrap;
		}
		.spinner {
			width: 28px;
			height: 28px;
			border: 3px solid color-mix(in srgb, var(--vscode-foreground) 20%, transparent);
			border-top-color: var(--vscode-progressBar-background, var(--vscode-foreground));
			border-radius: 50%;
			animation: spin 0.9s linear infinite;
		}
		@keyframes spin { to { transform: rotate(360deg); } }
	</style>
</head>
<body data-toolbar="${paged ? 'paged' : 'none'}">
	<div id="viewer-root">
		<div id="viewer-toolbar" role="toolbar" aria-label="Viewer controls">
			<button id="nav-prev" type="button" aria-label="Previous" title="Previous (←)" disabled>◀</button>
			<span id="position-label" aria-live="polite"></span>
			<button id="nav-next" type="button" aria-label="Next" title="Next (→)" disabled>▶</button>
			<div class="toolbar-separator"></div>
			<button id="fit-width" type="button" aria-label="Fit width" title="Fit width">↔</button>
			<button id="zoom-out" type="button" aria-label="Zoom out" title="Zoom out (Ctrl+-)" disabled>−</button>
			<span id="zoom-label" aria-live="polite">100%</span>
			<button id="zoom-in" type="button" aria-label="Zoom in" title="Zoom in (Ctrl+=)" disabled>+</button>
			<div class="toolbar-separator"></div>
			<button id="find-open" type="button" aria-label="Find" title="Find (Ctrl+F)">⌕</button>
		</div>
		<div id="find-popup" role="dialog" aria-label="Find" hidden>
			<input id="find-input" type="text" aria-label="Find" placeholder="Find" autocomplete="off" spellcheck="false" />
			<span id="find-status" aria-live="polite"></span>
			<button id="find-previous" type="button" aria-label="Previous match" title="Previous match (Shift+Enter)" disabled>↑</button>
			<button id="find-next" type="button" aria-label="Next match" title="Next match (Enter)" disabled>↓</button>
			<button id="find-close" type="button" aria-label="Close find" title="Close (Escape)">×</button>
		</div>
		<div id="viewer-container">
			<div id="canvas-host"></div>
			<div id="status" data-state="busy"><div class="spinner"></div></div>
		</div>
	</div>
	<script nonce="${nonce}">window.__OFFICE_FORMAT__ = ${JSON.stringify(format)};</script>
	<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
	}
}

/** Origin that serves the Google Fonts *stylesheets* (`<link rel="stylesheet">`). */
const GOOGLE_FONTS_CSS_ORIGIN = 'https://fonts.googleapis.com';
/** Origin that serves the actual `.woff2` font binaries referenced by the CSS. */
const GOOGLE_FONTS_FILES_ORIGIN = 'https://fonts.gstatic.com';

/**
 * Build the webview Content-Security-Policy string.
 *
 * The policy is fully offline unless the user opts into webfonts: the only
 * allowed origin is the extension's own `cspSource` (the locally served
 * webview resources). When `useGoogleFonts` is true exactly two directives
 * are widened (fonts stylesheet + font binaries) so the engine can load its
 * metric-compatible web fonts; nothing else can leave the machine.
 * `worker-src` keeps the data:/blob: forms the engine uses to bootstrap its
 * sandboxed parser worker.
 */
export function buildContentSecurityPolicy(
	cspSource: string,
	nonce: string,
	useGoogleFonts: boolean,
): string {
	const fontSrc = useGoogleFonts
		? `font-src ${cspSource} ${GOOGLE_FONTS_FILES_ORIGIN};`
		: `font-src ${cspSource};`;
	const styleSrc = useGoogleFonts
		? `style-src 'unsafe-inline' ${GOOGLE_FONTS_CSS_ORIGIN};`
		: `style-src 'unsafe-inline';`;

	return [
		`default-src 'none';`,
		`base-uri ${cspSource};`,
		`img-src ${cspSource} data: blob:;`,
		`media-src ${cspSource} blob:;`,
		fontSrc,
		`script-src 'nonce-${nonce}' 'wasm-unsafe-eval';`,
		`worker-src data: blob:;`,
		styleSrc,
		`connect-src ${cspSource} data: blob:;`,
	].join(' ');
}
