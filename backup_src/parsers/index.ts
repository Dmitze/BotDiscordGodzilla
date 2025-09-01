import { ParserRouter, type ParserRouterOptions } from './ParserRouter';
import type { IParser, ParseInput, ParseResult } from './IParser';
import { PlainTextParser } from './PlainTextParser';
import { GoogleDocsExportParser } from './GoogleDocsExportParser';
import { SheetsParser } from './SheetsParser';
import { PdfParser } from './PdfParser';
import { DocxParser } from './DocxParser';
import { ImageOcrParser } from './ImageOcrParser';

export type { IParser, ParseInput, ParseResult };
export { ParserRouter };

/**
 * Create default ParserRouter with sane ordering and options.
 * Order:
 *  - Plain text / JSON / CSV
 *  - Google Docs (export text/plain)
 *  - Google Sheets (export csv)
 *  - PDF (pdf-parse with OCR fallback)
 *  - DOC/DOCX (mammoth with raw fallback)
 *  - Images (OCR)
 */
export function createDefaultParserRouter(options: ParserRouterOptions = {}): ParserRouter {
  const parsers = [
    new PlainTextParser(),
    new GoogleDocsExportParser(),
    new SheetsParser(),
    new PdfParser(),
    new DocxParser(),
    new ImageOcrParser(),
  ];
  return new ParserRouter(parsers, options);
}
