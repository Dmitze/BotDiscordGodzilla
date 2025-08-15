import type { DriveFile } from '@/types/drive';
import { basename } from 'path';
import logger from '@/utils/logger';
import { sanitizeInput } from '@/utils/security';
import type { GoogleService } from '@/services/GoogleService';
import pdfParse from 'pdf-parse';
import * as mammoth from 'mammoth';
import * as xlsx from 'xlsx';

export interface ExtractResult {
  text: string;
  warnings: string[];
  mimeType: string;
  fileName: string;
}

// Поддерживаемые MIME для локального парсинга из буфера
const MIME: {
  TEXT: string[];
  JSON: string[];
  CSV: string[];
  PDF: string[];
  DOCX: string[];
  DOC: string[];
  GOOGLE_DOC: string[];
  GOOGLE_SHEET: string[];
  GOOGLE_SLIDES: string[];
  EXCEL: string[];
  FOLDER: string[];
} = {
  TEXT: ['text/plain', 'text/markdown'],
  JSON: ['application/json'],
  CSV: ['text/csv'],
  PDF: ['application/pdf'],
  DOCX: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  DOC: ['application/msword'],
  GOOGLE_DOC: ['application/vnd.google-apps.document'],
  GOOGLE_SHEET: ['application/vnd.google-apps.spreadsheet'],
  GOOGLE_SLIDES: ['application/vnd.google-apps.presentation'],
  EXCEL: [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
  ],
  FOLDER: ['application/vnd.google-apps.folder'],
};

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
    // Ранний выход для папок — скачивание невозможно
    if (MIME.FOLDER.includes(mime)) {
      warnings.push('Папка: бинарное содержимое отсутствует, пропуск без запроса к Drive');
      return { text: '', warnings, mimeType: mime, fileName: name };
    }

    // Google native (Docs/Sheets/Slides)
    const native = await handleGoogleNative(googleService, file, warnings);
    if (native) return native;

    // Остальные — скачиваем бинарно и маршрутизируем по MIME
    const buf = await googleService.downloadFile(file.id);
    return await handleBinaryByMime(buf, mime, name, warnings);
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

async function handleGoogleNative(
  googleService: GoogleService,
  file: DriveFile,
  warnings: string[]
): Promise<ExtractResult | null> {
  const name = file.name || 'file';
  const mime = file.mimeType || '';
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
    const buf = await googleService.exportFile(file.id, 'application/pdf');
    const text = await tryParsePdf(buf, warnings);
    return { text, warnings, mimeType: 'application/pdf', fileName: name };
  }
  return null;
}

async function handleBinaryByMime(
  buf: Buffer,
  mime: string,
  name: string,
  warnings: string[]
): Promise<ExtractResult> {
  if (MIME.TEXT.includes(mime)) {
    return { text: sanitizeAndTrim(buf.toString('utf8')), warnings, mimeType: mime, fileName: name };
  }
  if (MIME.JSON.includes(mime)) {
    const raw = buf.toString('utf8');
    try {
      const obj: unknown = JSON.parse(raw);
      const text = sanitizeAndTrim(JSON.stringify(obj, null, 2));
      return { text, warnings, mimeType: mime, fileName: name };
    } catch {
      warnings.push('JSON parse failed, returning raw text');
      return { text: sanitizeAndTrim(raw), warnings, mimeType: mime, fileName: name };
    }
  }
  if (MIME.CSV.includes(mime)) {
    return { text: sanitizeAndTrim(buf.toString('utf8')), warnings, mimeType: mime, fileName: name };
  }
  if (MIME.EXCEL.includes(mime)) {
    const text = await tryParseXlsx(buf, warnings);
    return { text, warnings, mimeType: 'text/csv', fileName: name };
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
    warnings.push('Формат .doc не поддерживается для прямого извлечения текста');
    return { text: '', warnings, mimeType: mime, fileName: name };
  }
  warnings.push(`Неизвестный MIME: ${mime}`);
  return { text: '', warnings, mimeType: mime, fileName: name };
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

async function tryParseXlsx(buf: Buffer, warnings: string[]): Promise<string> {
  try {
    const wb = xlsx.read(buf, { type: 'buffer' });
    const sheetNames: string[] = Array.isArray(wb.SheetNames) ? wb.SheetNames : [];
    if (sheetNames.length === 0) return '';
    // Берём первую страницу и конвертируем в CSV для универсального текстового представления
    const first = sheetNames[0];
    if (!first) return '';
    const ws = wb.Sheets[first as keyof typeof wb.Sheets] as xlsx.WorkSheet | undefined;
    if (!ws) return '';
    const csv = xlsx.utils.sheet_to_csv(ws, { FS: ',', RS: '\n' });
    return sanitizeAndTrim(csv);
  } catch (e) {
    warnings.push('xlsx: не удалось извлечь текст из XLSX/XLS');
    return '';
  }
}

function sanitizeAndTrim(text: string): string {
  const cleaned = sanitizeInput(text ?? '');
  return (cleaned || '').trim();
}
