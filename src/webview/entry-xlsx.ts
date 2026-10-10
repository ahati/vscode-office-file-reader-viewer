import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { bootstrap, type ViewerHandle } from './shared';

bootstrap(async (ctx) => {
	// XlsxViewer owns its whole UI (canvas, sheet-tab bar, zoom slider), so
	// it gets the raw container instead of the padded canvas host.
		const viewer = new XlsxViewer(ctx.viewerContainer, {
			showZoomSlider: ctx.config.xlsxShowZoomSlider,
			useGoogleFonts: ctx.config.useGoogleFonts,
			wasmUrl: ctx.config.wasmUrl,
			onHyperlinkClick: (target) => ctx.openHyperlink(target),
			onError: (err) => console.error('[office-viewer] xlsx viewer error', err),
		});

	await viewer.load(ctx.bytes);

	return viewer;
});
