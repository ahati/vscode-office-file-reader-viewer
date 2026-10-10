/**
 * Shared webview glue: receives the file bytes from the extension host,
 * drives the toolbar / status / find UI, and delegates actual rendering to a
 * format-specific viewer created by one of the entry-*.ts files.
 */

export interface ViewerConfigMessage {
	/** Fetch metric-compatible webfonts from Google Fonts (user opt-in). */
	useGoogleFonts: boolean;
	enableTextSelection: boolean;
	xlsxShowZoomSlider: boolean;
	/** Absolute webview URI of this format's parser WASM asset. */
	wasmUrl: string;
}

export interface InitMessage {
	type: 'init';
	format: 'docx' | 'xlsx' | 'pptx';
	bytes: ArrayBuffer;
	config: ViewerConfigMessage;
	fileName: string;
}

/** Hyperlink target as reported by the viewer engines. */
export interface HyperlinkTarget {
	kind: 'external' | 'internal';
	url?: string;
	ref?: string;
	slideIndex?: number;
}

export interface NavControls {
	prev(): Promise<void> | void;
	next(): Promise<void> | void;
	goTo(index: number): Promise<void> | void;
	count(): number;
}

export interface ZoomControls {
	get(): number;
	set(scale: number): Promise<void> | void;
	fitWidth?(): Promise<void> | void;
}

export interface FindControls {
	run(query: string): Promise<readonly { matchIndex: number }[]>;
	next(): Promise<{ matchIndex: number } | null>;
	prev(): Promise<{ matchIndex: number } | null>;
	clear(): void;
}

/**
 * Handed to the format-specific entry so it can expose its viewer's
 * capabilities to the shared toolbar without knowing about the DOM.
 */
export interface ViewerContext {
	bytes: ArrayBuffer;
	config: ViewerConfigMessage;
	fileName: string;
	/** Scrolled, padded host for the viewer's own <canvas> (docx/pptx). */
	readonly canvasHost: HTMLElement;
	/** The full content area (xlsx viewer owns its complete UI inside this). */
	readonly viewerContainer: HTMLElement;
	readonly openHyperlink: (target: HyperlinkTarget) => void;
	/** Set by the entry to route internal links (e.g. pptx slide jumps). */
	onInternalLink?: (target: HyperlinkTarget) => void;
	setNav(nav: NavControls): void;
	setZoom(zoom: ZoomControls): void;
	setFind(find: FindControls): void;
	/** Report the current page/slide; total <= 0 means "not known yet". */
	setProgress(current: number, total: number, layoutComplete: boolean): void;
	/** Report a scale change; pass 0 to reset the label to the default. */
	setScale(scale: number): void;
}

export interface ViewerHandle {
	destroy(): void;
}

const ZOOM_STEP = 1.25;
const ZOOM_MIN = 0.1;
const ZOOM_MAX = 4;

/** Encode WASM bytes as a fetchable data: URL for the opaque-origin worker. */
function toWasmDataUrl(bytes: Uint8Array): string {
	let binary = '';
	const chunkSize = 0x8000;
	for (let i = 0; i < bytes.length; i += chunkSize) {
		binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
	}
	return `data:application/wasm;base64,${btoa(binary)}`;
}

function debounce(fn: () => void, ms: number): () => void {
	let timer: number | undefined;
	return () => {
		if (timer !== undefined) window.clearTimeout(timer);
		timer = window.setTimeout(() => {
			timer = undefined;
			fn();
		}, ms);
	};
}

export function bootstrap(create: (ctx: ViewerContext) => Promise<ViewerHandle>): void {
	const vscode = acquireVsCodeApi();
	const viewerContainer = document.getElementById('viewer-container') as HTMLElement;
	const canvasHost = document.getElementById('canvas-host') as HTMLElement;
	const status = document.getElementById('status') as HTMLElement;
	const paged = document.body.dataset.toolbar === 'paged';

	const positionLabel = document.getElementById('position-label') as HTMLSpanElement;
	const zoomLabel = document.getElementById('zoom-label') as HTMLSpanElement;
	const navPrev = document.getElementById('nav-prev') as HTMLButtonElement;
	const navNext = document.getElementById('nav-next') as HTMLButtonElement;
	const fitWidthBtn = document.getElementById('fit-width') as HTMLButtonElement;
	const zoomInBtn = document.getElementById('zoom-in') as HTMLButtonElement;
	const zoomOutBtn = document.getElementById('zoom-out') as HTMLButtonElement;
	const findOpenBtn = document.getElementById('find-open') as HTMLButtonElement;
	const findPopup = document.getElementById('find-popup') as HTMLDivElement;
	const findInput = document.getElementById('find-input') as HTMLInputElement;
	const findStatus = document.getElementById('find-status') as HTMLSpanElement;
	const findNextBtn = document.getElementById('find-next') as HTMLButtonElement;
	const findPrevBtn = document.getElementById('find-previous') as HTMLButtonElement;
	const findCloseBtn = document.getElementById('find-close') as HTMLButtonElement;

	let currentViewer: ViewerHandle | undefined;
	let generation = 0;
	let nav: NavControls | undefined;
	let zoom: ZoomControls | undefined;
	let find: FindControls | undefined;
	let findTotal = 0;
	let lastScale = 0;

	// Capabilities appear as the entry registers them.
	fitWidthBtn.style.display = 'none';
	findOpenBtn.style.display = 'none';

	const setStatus = (state: 'idle' | 'busy' | 'error', message?: string) => {
		status.dataset.state = state;
		status.textContent =
			state === 'busy' ? '' : state === 'error' ? (message ?? 'Failed to open the file.') : '';
		if (state === 'busy') {
			const spinner = document.createElement('div');
			spinner.className = 'spinner';
			status.appendChild(spinner);
		}
	};

	const refreshNavButtons = () => {
		navPrev.disabled = !nav;
		navNext.disabled = !nav;
	};

	const updateFindButtons = () => {
		findNextBtn.disabled = findTotal === 0;
		findPrevBtn.disabled = findTotal === 0;
	};

	const teardown = () => {
		nav = undefined;
		zoom = undefined;
		find = undefined;
		findTotal = 0;
		lastScale = 0;
		positionLabel.textContent = '';
		zoomLabel.textContent = '100%';
		findStatus.textContent = '';
		findPopup.hidden = true;
		refreshNavButtons();
		updateFindButtons();
		try {
			currentViewer?.destroy();
		} catch {
			// destroy() may throw when a previous load never completed; ignore.
		}
		currentViewer = undefined;
		canvasHost.replaceChildren();
	};

	const ctx: ViewerContext = {
		bytes: new ArrayBuffer(0),
		config: {
			useGoogleFonts: false,
			enableTextSelection: true,
			xlsxShowZoomSlider: true,
			wasmUrl: '',
		},
		fileName: '',
		canvasHost,
		viewerContainer,
		onInternalLink: undefined,
		openHyperlink(target) {
			if (target.kind === 'external' && typeof target.url === 'string') {
				vscode.postMessage({ type: 'link', url: target.url });
				return;
			}
			// Internal link: prefer the entry-provided handler (e.g. slide jumps).
			if (ctx.onInternalLink) {
				ctx.onInternalLink(target);
			} else {
				console.info('[office-viewer] internal link ignored:', target.ref);
			}
		},
		setNav(next) {
			nav = next;
			refreshNavButtons();
		},
		setZoom(next) {
			zoom = next;
			zoomInBtn.disabled = false;
			zoomOutBtn.disabled = false;
			fitWidthBtn.style.display = zoom.fitWidth ? '' : 'none';
		},
		setFind(next) {
			find = next;
			findOpenBtn.style.display = '';
		},
		setProgress(current, total, layoutComplete) {
			if (!paged) return;
			const noun = document.body.dataset.format === 'pptx' ? 'Slide' : 'Page';
			positionLabel.textContent =
				total > 0
					? `${noun} ${current + 1} of ${total}${layoutComplete ? '' : '…'}`
					: '';
			navPrev.disabled = !nav || current <= 0;
			navNext.disabled = !nav || (total > 0 && current >= total - 1);
		},
		setScale(scale) {
			lastScale = scale > 0 ? scale : 0;
			zoomLabel.textContent = scale > 0 ? `${Math.round(scale * 100)}%` : '100%';
		},
	};

	// ---- Toolbar wiring ------------------------------------------------------

	navPrev.addEventListener('click', () => void nav?.prev());
	navNext.addEventListener('click', () => void nav?.next());
	fitWidthBtn.addEventListener('click', () => void zoom?.fitWidth?.());

	const applyZoom = (factor: number) => {
		if (!zoom) return;
		const next = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoom.get() * factor));
		void zoom.set(next);
	};
	zoomInBtn.addEventListener('click', () => applyZoom(ZOOM_STEP));
	zoomOutBtn.addEventListener('click', () => applyZoom(1 / ZOOM_STEP));

	const closeFind = () => {
		findPopup.hidden = true;
		find?.clear();
		findStatus.textContent = '';
		findTotal = 0;
		updateFindButtons();
	};

	findOpenBtn.addEventListener('click', () => {
		findPopup.hidden = false;
		findInput.focus();
		findInput.select();
	});
	findCloseBtn.addEventListener('click', closeFind);
	findNextBtn.addEventListener('click', () => void stepFind(1));
	findPrevBtn.addEventListener('click', () => void stepFind(-1));

	async function stepFind(direction: 1 | -1): Promise<void> {
		if (!find || findTotal === 0) return;
		const active = direction > 0 ? await find.next() : await find.prev();
		if (active) {
			findStatus.textContent = `${active.matchIndex + 1} / ${findTotal}`;
		}
	}

	let findRunTimer: number | undefined;
	findInput.addEventListener('input', () => {
		if (findRunTimer !== undefined) window.clearTimeout(findRunTimer);
		findRunTimer = window.setTimeout(() => void runFind(), 250);
	});

	async function runFind(): Promise<void> {
		if (!find) return;
		const query = findInput.value;
		if (!query) {
			find?.clear();
			findTotal = 0;
			findStatus.textContent = '';
			updateFindButtons();
			return;
		}
		try {
			const matches = await find.run(query);
			findTotal = matches.length;
			findStatus.textContent = findTotal > 0 ? `1 / ${findTotal}` : 'No results';
			updateFindButtons();
		} catch (err) {
			console.error('[office-viewer] find failed', err);
		}
	}

	findInput.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			void stepFind(event.shiftKey ? -1 : 1);
		} else if (event.key === 'Escape') {
			event.preventDefault();
			closeFind();
		}
	});

	// ---- Keyboard shortcuts --------------------------------------------------

	window.addEventListener('keydown', (event) => {
		if (!paged) return;
		const mod = event.ctrlKey || event.metaKey;
		const inInput =
			event.target instanceof HTMLElement &&
			(event.target.tagName === 'INPUT' || event.target.tagName === 'TEXTAREA');
		if (event.key === 'Escape' && !findPopup.hidden) {
			closeFind();
			return;
		}
		if (mod && event.key.toLowerCase() === 'f') {
			event.preventDefault();
			if (find) {
				findPopup.hidden = false;
				findInput.focus();
				findInput.select();
			}
			return;
		}
		if (inInput) return;
		if (mod && (event.key === '=' || event.key === '+')) {
			event.preventDefault();
			applyZoom(ZOOM_STEP);
			return;
		}
		if (mod && event.key === '-') {
			event.preventDefault();
			applyZoom(1 / ZOOM_STEP);
			return;
		}
		if (mod && event.key === '0') {
			event.preventDefault();
			void zoom?.set(1);
			return;
		}
		if (!nav) return;
		switch (event.key) {
			case 'ArrowRight':
			case 'ArrowDown':
			case 'PageDown':
				event.preventDefault();
				void nav.next();
				break;
			case 'ArrowLeft':
			case 'ArrowUp':
			case 'PageUp':
				event.preventDefault();
				void nav.prev();
				break;
			case 'Home':
				event.preventDefault();
				void nav.goTo(0);
				break;
			case 'End': {
				event.preventDefault();
				const count = nav.count();
				if (count > 0) void nav.goTo(count - 1);
				break;
			}
		}
	});

	// ---- Load / reload lifecycle --------------------------------------------

	const refitOnResize = debounce(() => {
		// Only auto-refit while the user has not chosen an explicit zoom.
		if (zoom?.fitWidth && lastScale === 0) void zoom.fitWidth();
	}, 150);
	window.addEventListener('resize', refitOnResize);

	window.addEventListener('message', (event: MessageEvent<InitMessage>) => {
		const message = event.data;
		if (message?.type !== 'init') return;

		const thisGeneration = ++generation;
		teardown();
		setStatus('busy');
		document.body.dataset.format = message.format;

		ctx.bytes = message.bytes;
		ctx.config = message.config;
		ctx.fileName = message.fileName;

		Promise.resolve()
			.then(async () => {
				// The parser worker bootstraps from an inline `data:` URL and
				// therefore has an opaque origin: it cannot go through the
				// webview resource service worker, so fetching the WASM asset
				// URL there fails. Fetch on the main thread (which can) and
				// hand the worker a self-contained data: URL instead.
				if (message.config.wasmUrl) {
					const resp = await fetch(message.config.wasmUrl);
					if (!resp.ok) throw new Error(`Fetching parser WASM failed with HTTP ${resp.status}`);
					const bytes = new Uint8Array(await resp.arrayBuffer());
					message.config.wasmUrl = toWasmDataUrl(bytes);
				}
				return create(ctx);
			})
			.then((viewer) => {
				if (thisGeneration !== generation) {
					// A newer init arrived while this one was loading.
					try {
						viewer.destroy();
					} catch {
						/* ignore */
					}
					return;
				}
				currentViewer = viewer;
				setStatus('idle');
				// Expose for debugging in the webview devtools console.
				(window as unknown as Record<string, unknown>).__officeViewer = viewer;
				vscode.postMessage({ type: 'loaded', fileName: message.fileName });
			})
			.catch((err: unknown) => {
				if (thisGeneration !== generation) return;
				console.error('[office-viewer] failed to render document', err);
				const detail = err instanceof Error ? `${err.name}: ${err.message}` : String(err);
				vscode.postMessage({ type: 'load-error', fileName: message.fileName, message: detail });
				setStatus(
					'error',
					`Could not render ${message.fileName}\n\n${detail}\n\nThe file may be corrupt, encrypted, or use unsupported features.`,
				);
			});
	});

	vscode.postMessage({ type: 'ready' });
}
