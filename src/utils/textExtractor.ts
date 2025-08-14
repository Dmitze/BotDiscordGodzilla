import type { DriveFile } from '@/types/drive';
import { basename } from 'path';
import logger from '@/utils/logger';
import { sanitizeInput } from '@/utils/aiEnhanced';
import type { GoogleService } from '@/services/GoogleService';
import pdfParse from 'pdf-parse';
import * as mammoth from 'mammoth';

export interface ExtractResult {
  text: string;
  warnings: string[];
  mimeType: string;
  fileName: string;
}

// Поддерживаемые MIME для локального парсинга из буфера
const MIME = {
  TEXT: ['text/plain', 'text/markdown'],
  JSON: ['application/json'],
  CSV: ['text/csv'],
  PDF: ['application/pdf'],
  DOCX: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  DOC: ['application/msword'],
  GOOGLE_DOC: ['application/vnd.google-apps.document'],
  GOOGLE_SHEET: ['application/vnd.google-apps.spreadsheet'],
  GOOGLE_SLIDES: ['application/vnd.google-apps.presentation'],
} as const;

/**
 * Унифицированный роутер извлечения текста для файлов Google Drive
 */
export async function extractTextFromDriveFile(
  googleService: GoogleService,
  file: DriveFile
): Promise<ExtractResult> {
  const warnings: string[] = [];
  const name = file.name || 'file';
  const mime = file.mimeType || '';

  try {
    // Google native docs → export
    if (MIME.GOOGLE_DOC.includes(mime)) {
      const buf = await googleService.exportFile(file.id, 'text/plain');
      const text = sanitizeAndTrim(buf.toString('utf8'));
      return { text, warnings, mimeType: 'text/plain', fileName: name };
    }
    if (MIME.GOOGLE_SHEET.includes(mime)) {
      const buf = await googleService.exportFile(file.id, 'text/csv');
      const text = sanitizeAndTrim(buf.toString('utf8'));
      return { text, warnings, mimeType: 'text/csv', fileName: name };
    }
    if (MIME.GOOGLE_SLIDES.includes(mime)) {
      // Экспорт в txt зачастую недоступен; пробуем PDF → не извлекаем тут текст, отдадим как предупреждение
      const buf = await googleService.exportFile(file.id, 'application/pdf');
      const text = await tryParsePdf(buf, warnings);
      return { text, warnings, mimeType: 'application/pdf', fileName: name };
    }

    // Binary/download path
    const buf = await googleService.downloadFile(file.id);

    // Simple routes
    if (MIME.TEXT.includes(mime)) {
      return { text: sanitizeAndTrim(buf.toString('utf8')), warnings, mimeType: mime, fileName: name };
    }
    if (MIME.JSON.includes(mime)) {
      const raw = buf.toString('utf8');
      try {
        const obj = JSON.parse(raw);
        const text = sanitizeAndTrim(JSON.stringify(obj, null, 2));
        return { text, warnings, mimeType: mime, fileName: name };
      } catch (e) {
        warnings.push('JSON parse failed, returning raw text');
        return { text: sanitizeAndTrim(raw), warnings, mimeType: mime, fileName: name };
      }
    }
    if (MIME.CSV.includes(mime)) {
      return { text: sanitizeAndTrim(buf.toString('utf8')), warnings, mimeType: mime, fileName: name };
    }
    if (MIME.PDF.includes(mime)) {
      const text = await tryParsePdf(buf, warnings);
      return { text, warnings, mimeType: mime, fileName: name };
    }
    if (MIME.DOCX.includes(mime)) {
      const text = await tryParseDocx(buf, warnings);
      return { text, warnings, mimeType: mime, fileName: name };
    }
    if (MIME.DOC.includes(mime)) {
      // .doc старый формат — прямой парсер отсутствует. Предложим конверсию через экспорт, если это Google Doc (не здесь), иначе вернём предупреждение
      warnings.push('Формат .doc не поддерживается для прямого извлечения текста');
      return { text: '', warnings, mimeType: mime, fileName: name };
    }

    warnings.push(`Неизвестный MIME: ${mime}`);
    return { text: '', warnings, mimeType: mime, fileName: name };
  } catch (error) {
    const err = error instanceof Error ? error.message : String(error);
    logger.error('❌ Ошибка извлечения текста', {
      type: 'system',
      event: 'extract_text_failed',
      component: 'textExtractor',
      fileId: file.id,
      name,
      mime,
      error: err,
    });
    return { text: '', warnings: warnings.concat(err), mimeType: mime, fileName: name };
  }
}

export function summarizeText(text: string, maxChars = 1200): string {
  const trimmed = text.trim();
  if (trimmed.length <= maxChars) return trimmed;
  // Простой эвристический саммари: первые N символов до конца абзаца
  const slice = trimmed.slice(0, maxChars);
  const lastBreak = Math.max(slice.lastIndexOf('\n\n'), slice.lastIndexOf('\n'), slice.lastIndexOf('. '));
  return slice.slice(0, lastBreak > 300 ? lastBreak + 1 : maxChars) + '\n…';
}

export function toDiscordAttachment(fileName: string, data: Buffer) {
  // Минимальный контракт для последующей обёртки AttachmentBuilder
  // Здесь намеренно без импорта discord.js, чтобы не тянуть типов и зависимостей
  const safeName = basename(fileName) || 'file.txt';
  return { name: safeName, data } as const;
}

async function tryParsePdf(buf: Buffer, warnings: string[]): Promise<string> {
  try {
    const res = await pdfParse(buf);
    return sanitizeAndTrim(res.text || '');
  } catch (e) {
    warnings.push('pdf-parse: не удалось извлечь текст');
    return '';
  }
}

async function tryParseDocx(buf: Buffer, warnings: string[]): Promise<string> {
  try {
    const res = await mammoth.extractRawText({ buffer: buf });
    return sanitizeAndTrim(res.value || '');
  } catch (e) {
    warnings.push('mammoth: не удалось извлечь текст из DOCX');
    return '';
  }
}

function sanitizeAndTrim(text: string): string {
  const cleaned = sanitizeInput(text ?? '', { inputType: 'text' }).sanitized || '';
  return cleaned.trim();
}
