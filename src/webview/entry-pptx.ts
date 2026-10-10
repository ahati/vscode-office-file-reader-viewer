import { PptxViewer } from '@silurus/ooxml/pptx';
import { bootstrap, type HyperlinkTarget, type ViewerHandle } from './shared';

bootstrap(async (ctx) => {
	const canvas = document.createElement('canvas');
	ctx.canvasHost.appendChild(canvas);

		const viewer = new PptxViewer(canvas, {
			enableTextSelection: ctx.config.enableTextSelection,
			useGoogleFonts: ctx.config.useGoogleFonts,
			wasmUrl: ctx.config.wasmUrl,
			onSlideChange: (index, total, layoutComplete) => ctx.setProgress(index, total, layoutComplete),
			onScaleChange: (scale) => ctx.setScale(scale),
			onHyperlinkClick: (target) => ctx.openHyperlink(target),
			onError: (err) => console.error('[office-viewer] pptx viewer error', err),
		});

		await viewer.load(ctx.bytes);

		// Internal slide jumps from action buttons / hyperlinks.
		ctx.onInternalLink = (target: HyperlinkTarget) => {
			if (typeof target.slideIndex === 'number') {
				void viewer.goToSlide(target.slideIndex);
			} else {
				console.info('[office-viewer] internal link ignored:', target.ref);
			}
		};

		ctx.setNav({
			prev: () => viewer.prevSlide(),
			next: () => viewer.nextSlide(),
			goTo: (index) => viewer.goToSlide(index),
			count: () => viewer.slideCount,
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

	ctx.setProgress(0, viewer.slideCount, true);
	await viewer.fitWidth();

	return viewer;
});
