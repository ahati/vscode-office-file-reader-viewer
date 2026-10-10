import * as vscode from 'vscode';

/** Viewer settings, snapshotted once per editor resolution. */
export interface ViewerConfig {
	/**
	 * Fetch metric-compatible webfonts from Google Fonts while rendering.
	 * When false (default) only locally installed system fonts are used and
	 * rendering is fully offline.
	 */
	useGoogleFonts: boolean;
	enableTextSelection: boolean;
	xlsxShowZoomSlider: boolean;
}

export function getViewerConfig(): ViewerConfig {
	const config = vscode.workspace.getConfiguration('officeFileViewer');
	return {
		useGoogleFonts: config.get<boolean>('useGoogleFonts', false),
		enableTextSelection: config.get<boolean>('enableTextSelection', true),
		xlsxShowZoomSlider: config.get<boolean>('xlsxShowZoomSlider', true),
	};
}
