import * as vscode from 'vscode';

/**
 * Integration tests: open each sample with the matching custom editor and
 * wait for the provider's diagnostics to report a successful render.
 * Executed inside the extension host via @vscode/test-electron.
 */

const VIEW_TYPES: Record<string, string> = {
	docx: 'officeFileViewer.docx',
	xlsx: 'officeFileViewer.xlsx',
	pptx: 'officeFileViewer.pptx',
};

function sleep(ms: number): Promise<void> {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getDiagnostics(): Promise<string[]> {
	return (await vscode.commands.executeCommand('officeFileViewer.getDiagnostics')) as string[];
}

async function waitFor(predicate: () => boolean, timeoutMs: number, what: string): Promise<void> {
	const deadline = Date.now() + timeoutMs;
	while (Date.now() < deadline) {
		if (predicate()) return;
		await sleep(250);
	}
	throw new Error(`Timed out after ${timeoutMs}ms waiting for ${what}`);
}

function customTabFor(viewType: string): boolean {
	return vscode.window.tabGroups.all.some((group) =>
		group.tabs.some(
			(tab) =>
				tab.input instanceof vscode.TabInputCustom && tab.input.viewType === viewType,
		),
	);
}

async function testFormat(samplesDir: string, format: keyof typeof VIEW_TYPES): Promise<void> {
	const uri = vscode.Uri.file(`${samplesDir}/sample.${format}`);
	const viewType = VIEW_TYPES[format];
	const marker = `rendered sample.${format}`;

	await vscode.commands.executeCommand('vscode.openWith', uri, viewType);
	await waitFor(() => customTabFor(viewType), 30_000, `custom tab for ${viewType}`);

	const deadline = Date.now() + 60_000;
	let tail = '';
	while (Date.now() < deadline) {
		const events = await getDiagnostics();
		tail = events.slice(-8).join('\n');
		if (events.some((line) => line.includes(marker))) return;
		if (events.some((line) => line.includes(`render FAILED for sample.${format}`))) {
			throw new Error(`Renderer failed for ${format}:\n${tail}`);
		}
		await sleep(250);
	}
	throw new Error(`No render confirmation for ${format} within 60s. Last events:\n${tail}`);
}

export async function run(): Promise<void> {
	// The development extension's own path IS the workspace root.
	const ext = vscode.extensions.getExtension('silurus.office-file-viewer');
	const workspaceRoot =
		ext?.extensionPath ?? vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '';
	const samplesDir = `${workspaceRoot}/samples`;

	const results: string[] = [];
	for (const format of ['docx', 'xlsx', 'pptx'] as const) {
		await testFormat(samplesDir, format);
		results.push(`OK sample.${format}`);
	}
	console.log(`\n=== office-file-viewer integration results ===\n${results.join('\n')}\n`);
}
