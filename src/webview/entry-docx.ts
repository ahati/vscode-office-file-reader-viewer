import { DocxViewer } from '@silurus/ooxml/docx';
import { bootstrap } from './shared';

bootstrap(async (ctx) => {
	const canvas = document.createElement('canvas');
	ctx.canvasHost.appendChild(canvas);

	const viewer = new DocxViewer(canvas, {
		container: ctx.canvasHost,
		enableTextSelection: ctx.config.enableTextSelection,
		useGoogleFonts: ctx.config.useGoogleFonts,
		wasmUrl: ctx.config.wasmUrl,
		onPageChange: (index, total, layoutComplete) => ctx.setProgress(index, total, layoutComplete),
		onScaleChange: (scale) => ctx.setScale(scale),
		onHyperlinkClick: (target) => ctx.openHyperlink(target),
		onError: (err) => console.error('[office-viewer] docx viewer error', err),
	});

	await viewer.load(ctx.bytes);

	ctx.setNav({
		prev: () => viewer.prevPage(),
		next: () => viewer.nextPage(),
		goTo: (index) => viewer.goToPage(index),
		count: () => viewer.pageCount,
	});
	ctx.setZoom({
		get: () => viewer.getScale(),
		set: (scale) => viewer.setScale(scale),
		fitWidth: () => viewer.fitWidth(),
	});
	ctx.setFind({
		run: (query) => viewer.findText(query),
		next: () => viewer.findNext(),
		prev: () => viewer.findPrev(),
		clear: () => viewer.clearFind(),
	});

	// Size the page to the available width once layout is known.
	await viewer.fitWidth();

	return viewer;
});
