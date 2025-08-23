import { readFileSync, existsSync } from 'fs';
import { resolve } from 'path';
import { config as dotenv } from 'dotenv';
import logger from '@/utils/logger';
import { Config } from '@/config/Config';

(async () => {
  try {
    dotenv();
    const args = process.argv.slice(2);
    if (!args.length) {
      console.error('Usage: npm run ocr:smoke -- <path-to-image>');
      process.exit(2);
    }
    const [arg0] = args as [string, ...string[]];
    const filePath = resolve(process.cwd(), arg0);
    if (!existsSync(filePath)) {
      console.error(`File not found: ${filePath}`);
      process.exit(2);
    }

    const cfg = Config.load();
    const langs = cfg.google.tesseractLangs || 'eng';
    const langPath = cfg.google.tesseractLangPath || undefined;

    // Динамический импорт tesseract.js, офлайн обработка
    const { createWorker } = await import('tesseract.js');
    const worker = await createWorker({ langPath } as any);
    await (worker as any).loadLanguage(langs);
    await (worker as any).initialize(langs);

    const buf = readFileSync(filePath);
    const { data } = (await (worker as any).recognize(buf));
    const text: string = data?.text ?? '';

    console.log('--- OCR RESULT START ---');
    console.log(text || '[empty]');
    console.log('--- OCR RESULT END ---');

    await (worker as any).terminate();
  } catch (err) {
    logger.error('OCR smoke test failed', {
      type: 'script',
      component: 'ocrSmoke',
      error: err instanceof Error ? err.message : String(err),
    });
    process.exit(1);
  }
})();
