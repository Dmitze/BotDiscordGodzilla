import logger from '@/utils/logger';
import type { MetricsService } from '../MetricsService';
import { type DocBody, type DocParagraph, type DocBlock, type DocContentElement, isDocDocument, isContentElement, isParagraph, isParagraphElement } from '@/types/docs';

/**
 * DocsService: узкая логика для извлечения текста из Google Docs
 * Без сетевых вызовов. Сервис-парсер структуры документа.
 */
export class DocsService {
  constructor(private readonly metrics?: MetricsService) {}

  /**
   * Извлекает обычный текст из структуры Google Docs (Schema$Document)
   * Принимает unknown, валидирует через guards и проходит по параграфам/элементам.
   */
  public extractTextFromDoc(doc: unknown): string {
    const start = Date.now();
    try {
      if (!isDocDocument(doc)) {
        logger.warn('DocsService: некорректная структура документа', { type: typeof doc });
        return '';
      }
      const body: DocBody | undefined = doc.body;
      if (!body || !Array.isArray(body.content)) return '';

      const parts: string[] = [];
      for (const ce of body.content) {
        if (!isContentElement(ce)) continue;
        if (!ce.paragraph) continue;
        parts.push(this.extractFromParagraph(ce.paragraph));
      }
      const text = parts.filter(Boolean).join('\n');
      return text;
    } finally {
      const dur = Date.now() - start;
      try { this.metrics?.updateGoogleApiMetrics('docs', 'parse', 'ok', dur); } catch { /* noop: метрики не критичны */ }
    }
  }

  private extractFromParagraph(p: DocParagraph): string {
    if (!isParagraph(p) || !Array.isArray(p.elements)) return '';
    const segs: string[] = [];
    for (const el of p.elements) {
      if (!isParagraphElement(el)) continue;
      const tr = el.textRun;
      const content = tr?.content ?? '';
      if (typeof content === 'string' && content.length > 0) segs.push(content);
    }
    const joined = segs.join('');
    // Убираем лишние CR, нормализуем перевод строки
    return joined.replace(/\r\n?/g, '\n');
  }

  /**
   * Структурированная выгрузка: заголовки, списки, таблицы, сноски.
   */
  public extractBlocksFromDoc(doc: unknown): DocBlock[] {
    const start = Date.now();
    try {
      if (!isDocDocument(doc)) return [];
      const out: DocBlock[] = [];
      const body: DocBody | undefined = doc.body;
      const footnotes = (doc as { footnotes?: Record<string, { content?: Array<DocContentElement> }> }).footnotes;

      if (body?.content && Array.isArray(body.content)) {
        for (const ce of body.content) {
          if (!isContentElement(ce)) continue;
          if (ce.paragraph) {
            const p = ce.paragraph;
            const text = this.extractFromParagraph(p).trimEnd();
            const style = p.paragraphStyle?.namedStyleType || '';
            const isList = !!p.bullet?.listId;
            const headingLevel = this.getHeadingLevel(style);
            if (headingLevel) {
              out.push({ kind: 'heading', level: headingLevel, text });
            } else if (isList) {
              const base = { kind: 'listItem' as const, text };
              const withList = p.bullet?.listId ? { listId: p.bullet.listId } : {};
              out.push({ ...base, ...withList });
            } else if (text.length > 0) {
              out.push({ kind: 'paragraph', text });
            }
          } else if (ce.table) {
            const rows = (ce.table.tableRows || []).map(r => ({
              cells: (r.tableCells || []).map(c => ({
                text: this.extractTextFromContent(c.content || []),
              })),
            }));
            out.push({ kind: 'table', rows });
          }
        }
      }

      // Footnotes
      if (footnotes && typeof footnotes === 'object') {
        for (const [fid, f] of Object.entries(footnotes)) {
          const text = this.extractTextFromContent(f.content || []);
          if (text) out.push({ kind: 'footnote', id: fid, text });
        }
      }

      return out;
    } finally {
      const dur = Date.now() - start;
      try { this.metrics?.updateGoogleApiMetrics('docs', 'parse_structured', 'ok', dur); } catch { /* noop: метрики не критичны */ }
    }
  }

  private getHeadingLevel(style: string | undefined): 1 | 2 | 3 | 4 | 5 | 6 | undefined {
    if (!style) return undefined;
    const m = /^HEADING_(\d)$/.exec(style);
    if (!m) return undefined;
    const n = Number(m[1]);
    if (n >= 1 && n <= 6) return n as 1 | 2 | 3 | 4 | 5 | 6;
    return undefined;
  }

  private extractTextFromContent(content: Array<DocContentElement>): string {
    const parts: string[] = [];
    for (const ce of content) {
      if (!isContentElement(ce)) continue;
      if (ce.paragraph) parts.push(this.extractFromParagraph(ce.paragraph));
      if (ce.table) {
        for (const row of ce.table.tableRows || []) {
          for (const cell of row.tableCells || []) {
            parts.push(this.extractTextFromContent(cell.content || []));
          }
        }
      }
    }
    return parts.filter(Boolean).join('\n');
  }
}
