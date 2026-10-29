import * as vscode from 'vscode';
import { OfficeEditorProvider, FormatHandler } from './provider';
import { csvHandler } from './formats/csv';
import { xlsxHandler } from './formats/xlsx';
import { docxHandler } from './formats/docx';
import { pptxHandler } from './formats/pptx';
import { pdfHandler } from './formats/pdf';

export function activate(context: vscode.ExtensionContext) {
  const register = (viewType: string, pick: (uri: vscode.Uri) => FormatHandler) =>
    context.subscriptions.push(
      vscode.window.registerCustomEditorProvider(viewType, new OfficeEditorProvider(context, pick), {
        webviewOptions: { retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: false,
      }),
    );

  register('officeEditor.table', (uri) => (uri.path.toLowerCase().endsWith('.xlsx') ? xlsxHandler : csvHandler));
  register('officeEditor.docx', () => docxHandler);
  register('officeEditor.pptx', () => pptxHandler);
  register('officeEditor.pdf', () => pdfHandler);
}

export function deactivate() {}
