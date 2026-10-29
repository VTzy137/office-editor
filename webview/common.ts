declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };

export const vscode = acquireVsCodeApi();

export interface Editor<M> {
  init(model: M, fileName: string): void | Promise<void>;
  getContent(): unknown | Promise<unknown>;
  undo?(): void;
  redo?(): void;
}

export function start<M>(editor: Editor<M>) {
  window.addEventListener('message', async (e) => {
    const msg = e.data;
    try {
      switch (msg.type) {
        case 'init':
          await editor.init(msg.model, msg.fileName);
          break;
        case 'fatal':
          showFatal(msg.message);
          break;
        case 'getContent':
          vscode.postMessage({ type: 'content', requestId: msg.requestId, payload: await editor.getContent() });
          break;
        case 'undo':
          editor.undo?.();
          break;
        case 'redo':
          editor.redo?.();
          break;
      }
    } catch (err: any) {
      console.error(err);
      if (msg.type === 'init') showFatal(String(err?.message ?? err));
      else vscode.postMessage({ type: 'error', message: String(err?.message ?? err) });
    }
  });
  vscode.postMessage({ type: 'ready' });
}

export const markDirty = (label?: string) => vscode.postMessage({ type: 'edit', label });

function showFatal(message: string) {
  const app = document.getElementById('app')!;
  app.replaceChildren(h('div', { class: 'fatal' }, h('h3', {}, 'Could not open this file'), h('pre', {}, message)));
}

type Attrs = Record<string, string | number | boolean | undefined | null>;

/** Tiny element builder: h('div', { class: 'x' }, child, 'text'). */
export function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Attrs = {}, ...children: (Node | string | null | undefined | false)[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    el.setAttribute(k, v === true ? '' : String(v));
  }
  for (const c of children) if (c !== null && c !== undefined && c !== false) el.append(c);
  return el;
}

/** Make pasting into a contenteditable insert plain text only. */
export function plainTextPaste(el: HTMLElement) {
  el.addEventListener('paste', (e) => {
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') ?? '';
    document.execCommand('insertText', false, text);
  });
}

export function button(label: string, title: string, onClick: () => void, cls = ''): HTMLButtonElement {
  const b = h('button', { class: 'tb-btn ' + cls, title, type: 'button' }, label);
  // Keep the text selection in the document when clicking toolbar buttons.
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', onClick);
  return b;
}
