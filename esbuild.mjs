import * as esbuild from 'esbuild';
import { readFile } from 'node:fs/promises';

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

/** @type {esbuild.BuildOptions} */
const extensionConfig = {
	entryPoints: ['src/extension.ts'],
	bundle: true,
	format: 'cjs',
	platform: 'node',
	target: 'node18',
	external: ['vscode'],
	outfile: 'out/extension.js',
	sourcemap: !production,
	minify: production,
};

/**
 * The published @silurus/ooxml bundles reference their parser WASM with Vite's
 * `new URL("*.wasm", import.meta.url).href` pattern, which cannot survive
 * esbuild's IIFE output. Rewrite those to static file imports so esbuild owns
 * the bytes and emits them as assets beside the webview bundle. The webview
 * HTML sets its <base> to the script directory so the emitted relative URLs
 * resolve against the webview resource origin.
 */
const bundledAssetSidecars = {
	name: 'bundled-asset-sidecars',
	setup(build) {
		build.onLoad(
			{ filter: /@silurus[/\\]ooxml[/\\]dist[/\\][^/\\]+\.(?:mjs|js)$/ },
			async ({ path }) => {
				const source = await readFile(path, 'utf8');
				const imports = new Map();
				const contents = source.replace(
					/new URL\((["'])([^"']+\.(?:wasm|ttf))\1, import\.meta\.url\)\.href/g,
					(_match, _quote, file) => {
						let identifier = imports.get(file);
						if (!identifier) {
							identifier = `__ooxmlAsset${imports.size}`;
							imports.set(file, identifier);
						}
						return identifier;
					},
				);
				if (imports.size === 0) return { contents: source, loader: 'js' };
				const prefix = [...imports]
					.map(([file, id]) => `import ${id} from ${JSON.stringify(`./${file}`)};`)
					.join('\n');
				return { contents: `${prefix}\n${contents}`, loader: 'js' };
			},
		);
	},
};

/** @type {esbuild.BuildOptions} */
const webviewConfig = {
	entryPoints: [
		{ in: 'src/webview/entry-docx.ts', out: 'webview-docx' },
		{ in: 'src/webview/entry-pptx.ts', out: 'webview-pptx' },
		{ in: 'src/webview/entry-xlsx.ts', out: 'webview-xlsx' },
	],
	bundle: true,
	format: 'iife',
	platform: 'browser',
	target: 'es2020',
	outdir: 'out/media',
	entryNames: '[name]',
	// Vite's plain ?url imports (notably parser WASM) are real files in this
	// bundle. The viewer does not bundle font binaries.
	// Parser WASM sidecars get deterministic names so the extension host can
	// compute a webview URI for each and pass it via `wasmUrl`.
	assetNames: 'assets/[name]',
	sourcemap: !production,
	minify: production,
	loader: {
		'.wasm': 'file',
		'.ttf': 'file',
	},
	plugins: [bundledAssetSidecars],
};

/** @type {esbuild.BuildOptions} */
const integrationTestConfig = {
	entryPoints: ['src/test-integration/index.ts'],
	bundle: true,
	format: 'cjs',
	platform: 'node',
	target: 'node18',
	external: ['vscode'],
	outfile: 'out/test/integration.js',
	sourcemap: !production,
};

async function build() {
	if (watch) {
		const [extCtx, wvCtx] = await Promise.all([
			esbuild.context(extensionConfig),
			esbuild.context(webviewConfig),
		]);
		await Promise.all([extCtx.watch(), wvCtx.watch()]);
		console.log('[esbuild] watching...');
	} else {
		await Promise.all([
			esbuild.build(extensionConfig),
			esbuild.build(webviewConfig),
			esbuild.build(integrationTestConfig),
		]);
		console.log('[esbuild] build complete');
	}
}

build().catch((err) => {
	console.error(err);
	process.exit(1);
});
