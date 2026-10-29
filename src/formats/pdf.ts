import { existsSync, readFileSync } from 'node:fs';
import { PDFDocument, PDFFont, StandardFonts, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import type { FormatHandler } from '../provider';

export interface PdfEdit {
  page: number;
  /** Baseline origin and size in PDF user space. */
  x: number;
  y: number;
  /** Width of the original text, which gets covered. */
  width: number;
  size: number;
  text: string;
  bold: boolean;
  italic: boolean;
  family: 'sans' | 'serif' | 'mono';
}

type Family = PdfEdit['family'];
type Variant = 'regular' | 'bold' | 'italic' | 'boldItalic';

const WIN = 'C:\\Windows\\Fonts\\';
const MAC = '/System/Library/Fonts/Supplemental/';
const SYSTEM_FONTS: Record<Family, Record<Variant, string[]>> = {
  sans: {
    regular: [WIN + 'arial.ttf', MAC + 'Arial.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'],
    bold: [WIN + 'arialbd.ttf', MAC + 'Arial Bold.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'],
    italic: [WIN + 'ariali.ttf', MAC + 'Arial Italic.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf', '/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf'],
    boldItalic: [WIN + 'arialbi.ttf', MAC + 'Arial Bold Italic.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf', '/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf'],
  },
  serif: {
    regular: [WIN + 'times.ttf', MAC + 'Times New Roman.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf', '/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf'],
    bold: [WIN + 'timesbd.ttf', MAC + 'Times New Roman Bold.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf', '/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf'],
    italic: [WIN + 'timesi.ttf', MAC + 'Times New Roman Italic.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf', '/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf'],
    boldItalic: [WIN + 'timesbi.ttf', MAC + 'Times New Roman Bold Italic.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSerif-BoldItalic.ttf', '/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf'],
  },
  mono: {
    regular: [WIN + 'cour.ttf', MAC + 'Courier New.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf', '/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf'],
    bold: [WIN + 'courbd.ttf', MAC + 'Courier New Bold.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf', '/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf'],
    italic: [WIN + 'couri.ttf', MAC + 'Courier New Italic.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Oblique.ttf', '/usr/share/fonts/truetype/liberation/LiberationMono-Italic.ttf'],
    boldItalic: [WIN + 'courbi.ttf', MAC + 'Courier New Bold Italic.ttf', '/usr/share/fonts/truetype/dejavu/DejaVuSansMono-BoldOblique.ttf', '/usr/share/fonts/truetype/liberation/LiberationMono-BoldItalic.ttf'],
  },
};

const STANDARD: Record<Family, Record<Variant, StandardFonts>> = {
  sans: { regular: StandardFonts.Helvetica, bold: StandardFonts.HelveticaBold, italic: StandardFonts.HelveticaOblique, boldItalic: StandardFonts.HelveticaBoldOblique },
  serif: { regular: StandardFonts.TimesRoman, bold: StandardFonts.TimesRomanBold, italic: StandardFonts.TimesRomanItalic, boldItalic: StandardFonts.TimesRomanBoldItalic },
  mono: { regular: StandardFonts.Courier, bold: StandardFonts.CourierBold, italic: StandardFonts.CourierOblique, boldItalic: StandardFonts.CourierBoldOblique },
};

class FontCache {
  private cache = new Map<string, { font: PDFFont; unicode: boolean }>();
  constructor(private doc: PDFDocument) {}

  async get(family: Family, variant: Variant) {
    const key = family + variant;
    let entry = this.cache.get(key);
    if (!entry) {
      const file = SYSTEM_FONTS[family][variant].find((f) => existsSync(f));
      // Prefer a TrueType system font so non-Latin text (e.g. Vietnamese) can be written.
      entry = file
        ? { font: await this.doc.embedFont(readFileSync(file), { subset: true }), unicode: true }
        : { font: await this.doc.embedFont(STANDARD[family][variant]), unicode: false };
      this.cache.set(key, entry);
    }
    return entry;
  }
}

function encodable(font: PDFFont, text: string): string {
  return Array.from(text)
    .map((ch) => {
      try {
        font.encodeText(ch);
        return ch;
      } catch {
        return '?';
      }
    })
    .join('');
}

export const pdfHandler: FormatHandler = {
  script: 'pdf',
  async load(bytes) {
    // workspace.fs returns a Node Buffer, whose toJSON() would turn it into {type,data} in postMessage;
    // a plain Uint8Array copy is transferred as binary.
    return { data: new Uint8Array(bytes) };
  },
  async save(base, payload: { edits: PdfEdit[] }) {
    if (!payload.edits.length) return base;
    const doc = await PDFDocument.load(base, { ignoreEncryption: true });
    doc.registerFontkit(fontkit);
    const fonts = new FontCache(doc);
    const pages = doc.getPages();
    for (const e of payload.edits) {
      const page = pages[e.page];
      if (!page) continue;
      // Cover the original glyphs, then draw the replacement on the same baseline.
      page.drawRectangle({
        x: e.x - 1,
        y: e.y - e.size * 0.25,
        width: e.width + 2,
        height: e.size * 1.2,
        color: rgb(1, 1, 1),
      });
      if (!e.text) continue;
      const variant: Variant = e.bold && e.italic ? 'boldItalic' : e.bold ? 'bold' : e.italic ? 'italic' : 'regular';
      const { font, unicode } = await fonts.get(e.family, variant);
      const text = unicode ? e.text : encodable(font, e.text);
      page.drawText(text, { x: e.x, y: e.y, size: e.size, font, color: rgb(0, 0, 0) });
    }
    return doc.save();
  },
};
