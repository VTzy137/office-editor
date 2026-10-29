import * as vscode from 'vscode';

/**
 * A format handler turns file bytes into a model the webview can render, and turns the
 * webview's edit payload back into file bytes. Payloads are always cumulative relative to
 * the bytes the document was opened with, so handlers can be stateless.
 */
export interface FormatHandler {
  /** Name of the webview bundle in media/dist (without .js). */
  readonly script: string;
  /** Use VS Code's undo stack (the webview implements undo/redo) instead of native undo. */
  readonly customUndo?: boolean;
  load(bytes: Uint8Array, uri: vscode.Uri): Promise<unknown>;
  save(baseBytes: Uint8Array, payload: any, uri: vscode.Uri): Promise<Uint8Array>;
}

class OfficeDocument implements vscode.CustomDocument {
  webview?: vscode.WebviewPanel;
  private pending = new Map<number, (payload: unknown) => void>();
  private nextRequest = 1;

  constructor(readonly uri: vscode.Uri, public baseBytes: Uint8Array) {}

  /** Ask the webview for its current edit payload. */
  requestPayload(): Promise<unknown> {
    const panel = this.webview;
    if (!panel) return Promise.resolve(null);
    const requestId = this.nextRequest++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId);
        reject(new Error('Timed out waiting for editor content'));
      }, 30000);
      this.pending.set(requestId, (p) => {
        clearTimeout(timer);
        resolve(p);
      });
      panel.webview.postMessage({ type: 'getContent', requestId });
    });
  }

  resolvePayload(requestId: number, payload: unknown) {
    this.pending.get(requestId)?.(payload);
    this.pending.delete(requestId);
  }

  dispose() {
    this.pending.clear();
  }
}

export class OfficeEditorProvider implements vscode.CustomEditorProvider<OfficeDocument> {
  private readonly changeEmitter = new vscode.EventEmitter<any>();
  readonly onDidChangeCustomDocument = this.changeEmitter.event;

  constructor(private readonly context: vscode.ExtensionContext, private readonly handlers: (uri: vscode.Uri) => FormatHandler) {}

  async openCustomDocument(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext): Promise<OfficeDocument> {
    const source = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const bytes = await vscode.workspace.fs.readFile(source);
    return new OfficeDocument(uri, bytes);
  }

  async resolveCustomEditor(document: OfficeDocument, panel: vscode.WebviewPanel): Promise<void> {
    const handler = this.handlers(document.uri);
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    document.webview = panel;
    panel.onDidDispose(() => (document.webview = undefined));
    panel.webview.options = { enableScripts: true, localResourceRoots: [mediaRoot] };
    panel.webview.html = this.html(panel.webview, handler.script);

    panel.webview.onDidReceiveMessage(async (msg) => {
      switch (msg.type) {
        case 'ready':
          await this.sendModel(document, handler);
          break;
        case 'edit':
          if (handler.customUndo) {
            this.changeEmitter.fire({
              document,
              label: msg.label ?? 'Edit',
              undo: () => panel.webview.postMessage({ type: 'undo' }),
              redo: () => panel.webview.postMessage({ type: 'redo' }),
            });
          } else {
            this.changeEmitter.fire({ document });
          }
          break;
        case 'content':
          document.resolvePayload(msg.requestId, msg.payload);
          break;
        case 'error':
          vscode.window.showErrorMessage(`Office Editor: ${msg.message}`);
          break;
      }
    });
  }

  private async sendModel(document: OfficeDocument, handler: FormatHandler) {
    try {
      const model = await handler.load(document.baseBytes, document.uri);
      document.webview?.webview.postMessage({ type: 'init', model, fileName: document.uri.path.split('/').pop() });
    } catch (err: any) {
      document.webview?.webview.postMessage({ type: 'fatal', message: String(err?.message ?? err) });
    }
  }

  private async render(document: OfficeDocument): Promise<Uint8Array> {
    const handler = this.handlers(document.uri);
    const payload = await document.requestPayload();
    if (payload == null) return document.baseBytes;
    return handler.save(document.baseBytes, payload, document.uri);
  }

  async saveCustomDocument(document: OfficeDocument): Promise<void> {
    await vscode.workspace.fs.writeFile(document.uri, await this.render(document));
  }

  async saveCustomDocumentAs(document: OfficeDocument, destination: vscode.Uri): Promise<void> {
    await vscode.workspace.fs.writeFile(destination, await this.render(document));
  }

  async revertCustomDocument(document: OfficeDocument): Promise<void> {
    document.baseBytes = await vscode.workspace.fs.readFile(document.uri);
    await this.sendModel(document, this.handlers(document.uri));
  }

  async backupCustomDocument(document: OfficeDocument, context: vscode.CustomDocumentBackupContext): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(context.destination, await this.render(document));
    return {
      id: context.destination.toString(),
      delete: () => vscode.workspace.fs.delete(context.destination).then(undefined, () => undefined),
    };
  }

  private html(webview: vscode.Webview, script: string): string {
    const media = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(media, 'dist', `${script}.js`));
    const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(media, 'style.css'));
    const pdfjsUri = webview.asWebviewUri(vscode.Uri.joinPath(media, 'dist', 'pdfjs'));
    const nonce = Array.from({ length: 32 }, () => Math.floor(Math.random() * 36).toString(36)).join('');
    const csp = [
      `default-src 'none'`,
      `img-src ${webview.cspSource} data: blob:`,
      `style-src ${webview.cspSource} 'unsafe-inline'`,
      `font-src ${webview.cspSource} data: blob:`,
      `script-src 'nonce-${nonce}' ${webview.cspSource}`,
      `worker-src ${webview.cspSource} blob:`,
      `connect-src ${webview.cspSource}`,
    ].join('; ');
    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="${csp}">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link href="${styleUri}" rel="stylesheet">
</head>
<body data-pdfjs="${pdfjsUri}">
  <div id="app"><div class="loading">Loading…</div></div>
  <script type="module" nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}
