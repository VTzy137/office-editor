# Office Editor for VS Code

Open and edit **CSV / TSV, XLSX, DOCX, PPTX and PDF** files directly in VS Code.

| Format | What you see | What you can edit |
| --- | --- | --- |
| CSV / TSV | Spreadsheet grid | Cells, insert/delete rows & columns, paste blocks from Excel |
| XLSX | Grid with sheet tabs and a formula bar | Cell values and formulas (`=SUM(A1:A3)`), rows & columns. Styles, other sheets and formatting are preserved |
| DOCX | Page view with headings, font sizes, bold/italic/underline, lists, tables, images | Text, **B**/*I*/U, font size, paragraph style (Heading 1, …), new paragraphs |
| PPTX | Slides rendered at their real layout (theme colors, placeholders, images, tables) | Text inside any text box, placeholder or table cell |
| PDF | Rendered pages (pdf.js) | Click any line of text and retype it |

Save with `Ctrl+S` as usual. Undo/redo works (`Ctrl+Z` / `Ctrl+Y`).

## Running it

```bash
npm install
npm run build
```

Then press **F5** in VS Code ("Run Extension") and open any supported file in the new window.
To install it permanently: `npm run package` and install the generated `.vsix`
(Extensions view → `…` → *Install from VSIX…*).

If you prefer VS Code's built-in editor for a file (e.g. CSV as text), right-click it → *Open With…*.

### Spreadsheet keys
Arrow keys / Tab move · type or `Enter`/`F2` to edit · `Enter` commits and moves down · `Esc` cancels ·
`Delete` clears · `Ctrl+Arrow` jumps to the edge · `Ctrl+C` / `Ctrl+V` copy and paste (multi-cell paste from Excel works).

## How saving works (and its limits)

The editors change the original file as little as possible:

- **DOCX**: paragraphs you didn't touch are written back exactly as they were. An edited paragraph is rebuilt from its text runs. Paragraphs containing images, fields (TOC, page numbers) or equations are read-only, shown with a light highlight on hover. Hyperlinks inside an *edited* paragraph become plain text.
- **PPTX**: only the text of the runs you changed is replaced, so formatting stays. You can't add new paragraphs (Enter is disabled); charts and SmartArt show as placeholders.
- **XLSX**: changes are replayed onto the original workbook with ExcelJS. Formulas get recalculated when Excel opens the file. Inserting or deleting rows **does not rewrite formulas that refer to shifted cells**, which is an ExcelJS limitation.
- **PDF**: PDFs have no editable "text" model, so an edit **covers the original line with a white box and draws the new text** in a similar font (Arial/Times/Courier from your system, so Vietnamese and other Unicode text works). The original text is still in the file underneath, so text extraction or search may find both. Edits look wrong on colored backgrounds. Rotated text can't be edited.
- **CSV**: delimiter (`,` `;` tab `|`), line endings, BOM and trailing newline are kept. Fields are re-quoted only where needed.

## Development

```
src/extension.ts        registers one custom editor per format
src/provider.ts         generic CustomEditorProvider (save, backup, revert, undo routing)
src/formats/*.ts        parse file → model for the webview; apply webview edits → file bytes
webview/*.ts            the editors' UI (bundled to media/dist)
```

- `npm run watch` rebuilds on change
- `npm run test:smoke` loads, edits and saves every file in `samples/` through the handlers
- `npm run test:webview` drives the real editor UIs in Edge (playwright-core, no browser download) and saves the result. Set `BROWSER_CHANNEL=chrome` to use Chrome instead.
