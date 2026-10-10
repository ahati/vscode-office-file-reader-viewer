import * as vscode from 'vscode';
import { OfficeEditorProvider, type OfficeFormat } from './officeEditorProvider';

export function activate(context: vscode.ExtensionContext): void {
	for (const format of ['docx', 'xlsx', 'pptx'] as const) {
		context.subscriptions.push(OfficeEditorProvider.register(context, format));
	}
	// Programmatic diagnostics access (tests / "Developer: Run Command"):
	context.subscriptions.push(
		vscode.commands.registerCommand('officeFileViewer.getDiagnostics', () =>
			OfficeEditorProvider.getDiagnostics(),
		),
	);
}

export type { OfficeFormat };

export function deactivate(): void {
	// Nothing to clean up; providers are disposed via context.subscriptions.
}
