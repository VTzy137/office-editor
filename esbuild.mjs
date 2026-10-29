import * as esbuild from 'esbuild';
import { cpSync, mkdirSync } from 'node:fs';

const watch = process.argv.includes('--watch');

// Copy pdf.js runtime assets (worker, CMaps, standard fonts) next to the webview bundles.
mkdirSync('media/dist/pdfjs', { recursive: true });
cpSync('node_modules/pdfjs-dist/build/pdf.worker.min.mjs', 'media/dist/pdfjs/pdf.worker.min.mjs');
cpSync('node_modules/pdfjs-dist/cmaps', 'media/dist/pdfjs/cmaps', { recursive: true });
cpSync('node_modules/pdfjs-dist/standard_fonts', 'media/dist/pdfjs/standard_fonts', { recursive: true });

const extension = {
  entryPoints: ['src/extension.ts'],
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node18',
  outfile: 'dist/extension.js',
  external: ['vscode'],
  sourcemap: true,
  minify: !watch,
};

const webviews = {
  entryPoints: ['webview/table.ts', 'webview/docx.ts', 'webview/pptx.ts', 'webview/pdf.ts'],
  bundle: true,
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  outdir: 'media/dist',
  sourcemap: true,
  minify: !watch,
};

if (watch) {
  const contexts = await Promise.all([esbuild.context(extension), esbuild.context(webviews)]);
  await Promise.all(contexts.map((c) => c.watch()));
  console.log('watching...');
} else {
  await Promise.all([esbuild.build(extension), esbuild.build(webviews)]);
  console.log('build complete');
}
