import 'dotenv/config';
import logger from '@/utils/logger';
import { Config } from '@/config/Config';
import { GoogleService } from '@/services/GoogleService';
// Типы tesseract.js могут отличаться между минорными версиями, поэтому используем узкие приведения ниже

// Допустимые MIME для изображений
const IMAGE_MIME = [
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp',
  'image/bmp',
  'image/tiff',
];

function extractIdFromUrl(input: string): { kind: 'file' | 'folder'; id: string } | null {
  const s = input.trim();
  // Если похоже на ID
  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return { kind: 'file', id: s };
  try {
    const url = new URL(s);
    // Файл: https://drive.google.com/file/d/<ID>/view
    const mFile = url.pathname.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
    if (mFile && mFile[1]) return { kind: 'file', id: mFile[1] };
    // Папка: https://drive.google.com/drive/folders/<ID>
    const mFolder = url.pathname.match(/\/drive\/folders\/([a-zA-Z0-9-_]+)/);
    if (mFolder && mFolder[1]) return { kind: 'folder', id: mFolder[1] };
    // Альтернативный формат: open?id=
    const openId = url.searchParams.get('id');
    if (openId && /^[a-zA-Z0-9-_]{20,}$/.test(openId)) return { kind: 'file', id: openId };
  } catch {
    // not an URL
  }
  return null;
}

async function main(): Promise<void> {
  const started = Date.now();
  try {
    const args = process.argv.slice(2);
    if (args.length === 0) {
      console.error('Usage: npm run ocr:smoke:drive -- <drive-file-or-folder-url-or-id>');
      process.exitCode = 2;
      return;
    }
    const arg0 = args[0];
    if (typeof arg0 !== 'string' || !arg0) {
      console.error('Первый аргумент должен быть строкой ссылки или ID Google Drive');
      process.exitCode = 2;
      return;
    }
    const input: string = arg0;

    const cfg = Config.load();
    const gs = new GoogleService(cfg);
    await gs.initialize();

    const langs = cfg.google.tesseractLangs || 'eng';
    const langPath = cfg.google.tesseractLangPath || undefined;

    const parsed = extractIdFromUrl(input);
    if (!parsed) {
      console.error('Не удалось распознать ссылку/ID Google Drive');
      process.exitCode = 2;
      return;
    }

    let fileId: string | null = null;
    if (parsed.kind === 'file') {
      fileId = parsed.id;
    } else {
      // Папка: найдём 1-ю картинку
      const page = await gs.listDriveFiles({
        folderId: parsed.id,
        mimeIncludes: IMAGE_MIME,
        pageSize: 10,
      });
      const img = page.files.find(f => IMAGE_MIME.includes(f.mimeType || '')) || page.files[0];
      if (!img) {
        console.error('В папке не найдено изображений');
        process.exitCode = 3;
        return;
      }
      fileId = img.id;
      logger.info('Выбран файл для OCR', { id: img.id, name: img.name, mime: img.mimeType });
    }

    if (!fileId) {
      console.error('Не удалось определить fileId');
      process.exitCode = 2;
      return;
    }

    const buf = await gs.downloadFile(fileId);

    // Динамический импорт tesseract.js (офлайн OCR)
    const { createWorker } = await import('tesseract.js');
    /* eslint-disable @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment */
    const worker = await (createWorker as unknown as (...args: any[]) => Promise<any>)({ langPath });
    await worker.loadLanguage(langs as any);
    await worker.initialize(langs as any);

    const { data } = await worker.recognize(buf as any);
    const text: string = (data && typeof data.text === 'string') ? data.text : '';
    /* eslint-enable @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment */

    // eslint-disable-next-line no-console
    console.log('--- OCR RESULT START ---');
    // eslint-disable-next-line no-console
    console.log(text || '[empty]');
    // eslint-disable-next-line no-console
    console.log('--- OCR RESULT END ---');

    /* eslint-disable @typescript-eslint/no-unsafe-call */
    await worker.terminate();
    /* eslint-enable @typescript-eslint/no-unsafe-call */
  } catch (error) {
    logger.error('OCR Drive smoke failed', {
      type: 'script',
      component: 'ocrSmokeDrive',
      error: error instanceof Error ? error.message : String(error),
    });
    process.exitCode = 1;
  } finally {
    logger.info('🏁 Завершено', { durationMs: Date.now() - started });
  }
}

void main();
