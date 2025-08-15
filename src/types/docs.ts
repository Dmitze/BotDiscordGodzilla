/*
 * Узкие интерфейсы для Google Docs + type-guards
 */

export interface DocTextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface DocTextRun {
  content: string; // Всегда строка, может заканчиваться "\n"
  textStyle?: DocTextStyle;
}

export interface DocParagraphElement {
  textRun?: DocTextRun;
}

export interface DocParagraph {
  elements?: DocParagraphElement[];
  paragraphStyle?: { namedStyleType?: string };
  bullet?: { listId?: string };
}

export interface DocContentElement {
  paragraph?: DocParagraph;
  table?: {
    tableRows?: Array<{
      tableCells?: Array<{
        content?: Array<DocContentElement>;
      }>;
    }>;
  };
}

export interface DocBody {
  content?: DocContentElement[];
}

export interface DocDocument {
  body?: DocBody;
  footnotes?: Record<string, { content?: Array<DocContentElement> }>;
}

// ----- Type guards -----
export function isParagraphElement(el: unknown): el is DocParagraphElement {
  if (!el || typeof el !== 'object') return false;
  const e = el as Record<string, unknown>;
  if (!('textRun' in e)) return false;
  const tr = (e as { textRun?: unknown }).textRun;
  if (tr == null) return true; // допускаем наличие пустого textRun
  if (typeof tr !== 'object') return false;
  const t = tr as { content?: unknown };
  return typeof t.content === 'string' || typeof t.content === 'undefined';
}

export function isParagraph(el: unknown): el is DocParagraph {
  if (!el || typeof el !== 'object') return false;
  const e = el as { elements?: unknown };
  if (e.elements == null) return true;
  return Array.isArray(e.elements) && e.elements.every(isParagraphElement);
}

export function isContentElement(el: unknown): el is DocContentElement {
  if (!el || typeof el !== 'object') return false;
  const e = el as { paragraph?: unknown; table?: unknown };
  if (e.paragraph == null) return true;
  if (!isParagraph(e.paragraph)) return false;
  if (e.table != null && typeof e.table !== 'object') return false;
  return true;
}

export function isDocBody(b: unknown): b is DocBody {
  if (!b || typeof b !== 'object') return false;
  const body = b as { content?: unknown };
  if (body.content == null) return true;
  return Array.isArray(body.content) && body.content.every(isContentElement);
}

export function isDocDocument(d: unknown): d is DocDocument {
  if (!d || typeof d !== 'object') return false;
  const doc = d as { body?: unknown; footnotes?: unknown };
  if (doc.body == null) return true;
  if (!isDocBody(doc.body)) return false;
  if (doc.footnotes != null && typeof doc.footnotes !== 'object') return false;
  return true;
}

// ===== DTO для структурированной выгрузки =====
export type DocBlock =
  | { kind: 'paragraph'; text: string; style?: DocTextStyle }
  | { kind: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; text: string; style?: DocTextStyle }
  | { kind: 'listItem'; text: string; listId?: string; style?: DocTextStyle }
  | { kind: 'table'; rows: TableRow[] }
  | { kind: 'footnote'; id: string; text: string };

export interface TableCell {
  text: string;
}

export interface TableRow {
  cells: TableCell[];
}

