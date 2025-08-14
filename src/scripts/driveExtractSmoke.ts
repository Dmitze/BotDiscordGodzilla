import 'dotenv/config';
import logger from '@/utils/logger';
import { Config } from '@/config/Config';
import { GoogleService } from '@/services/GoogleService';
import type { DriveFile } from '@/types/drive';
import { extractTextFromDriveFile, summarizeText } from '@/utils/textExtractor';

async function main(): Promise<void> {
  const started = Date.now();
  try {
    logger.info('🧪 Drive Extract Smoke Test старт');

    const cfg = Config.load();
    const folderId = cfg.drive.folderId;
    if (!folderId) {
      logger.error('❌ GOOGLE_DRIVE_FOLDER_ID не задан');
      process.exitCode = 2;
      return;
    }

    const gs = new GoogleService(cfg);
    await gs.initialize();

    // 1) Листинг 1–3 файлов
    const page = await gs.listDriveFiles({ folderId, pageSize: Math.min(3, cfg.drive.pageSize) });
    const files = page.files.slice(0, 3);

    if (files.length === 0) {
      logger.warn('⚠️ В папке нет файлов для проверки');
      return;
    }

    logger.info('📄 Файлы для проверки:', {
      count: files.length,
      names: files.map(f => `${f.name} (${f.mimeType})`),
    });

    for (const f of files) {
      await runOne(gs, f);
    }
  } catch (error) {
    logger.error('❌ Ошибка smoke-теста', {
      error: error instanceof Error ? error.message : String(error),
    });
    process.exitCode = 1;
  } finally {
    logger.info('🏁 Завершено', { durationMs: Date.now() - started });
  }
}

async function runOne(gs: GoogleService, file: DriveFile): Promise<void> {
  logger.info('▶️ Обработка файла', {
    id: file.id,
    name: file.name,
    mime: file.mimeType,
    size: file.size ?? 0,
  });

  try {
    const meta = await gs.getDriveFile(file.id);
    logger.debug('ℹ️ Метаданные', {
      id: meta.id,
      owners: meta.owners,
      shortcut: meta.isShortcut ? 'yes' : 'no',
    });

    const res = await extractTextFromDriveFile(gs, meta);
    const preview = summarizeText(res.text, 600);

    logger.info('✅ Извлечение завершено', {
      name: meta.name,
      mime: res.mimeType,
      textBytes: Buffer.byteLength(res.text, 'utf8'),
      warnings: res.warnings,
      previewSample: preview.slice(0, 200).replace(/\n/g, ' ⏎ '),
    });
  } catch (e) {
    logger.error('❌ Ошибка при обработке файла', {
      id: file.id,
      name: file.name,
      error: e instanceof Error ? e.message : String(e),
    });
  }
}

// Execute
void main().catch(err => {
  logger.error('❌ Неперехваченная ошибка в smoke-тесте', {
    error: err instanceof Error ? err.message : String(err),
  });
  process.exitCode = 1;
});
