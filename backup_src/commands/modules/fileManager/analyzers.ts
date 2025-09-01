import type { ChatInputCommandInteraction } from 'discord.js';
import { t } from '@/i18n';
import type { AIService } from '@/services/AIService';
import type { SheetsContextService } from '@/services/SheetsContextService';

export interface AnalyzeDeps {
  config: any;
  getGoogleService: (interaction: ChatInputCommandInteraction) => any | undefined;
  isMimeAllowed: (mime: string, allowed: string[]) => boolean;
  isOwnerAllowed: (owners: string[], allowlist: string[]) => boolean;
  isTooLarge: (bytes: number, limitMb: number) => boolean;
  getAnalysisTypeName: (t: 'summary' | 'detailed' | 'key_points' | string) => string;
  resolve: <T = unknown>(interaction: ChatInputCommandInteraction, name: string) => T | undefined;
}

export async function handleAnalyze(
  interaction: ChatInputCommandInteraction,
  options: { fileId: string; analysisType: 'summary' | 'detailed' | 'key_points' },
  deps: AnalyzeDeps
): Promise<{ success: boolean; message: string }> {
  const { config, getGoogleService, isMimeAllowed, isOwnerAllowed, isTooLarge, getAnalysisTypeName, resolve } = deps;

  const analysisTypeName = getAnalysisTypeName(options.analysisType);

  const ai = resolve<AIService>(interaction, 'ai');
  const googleSvc = getGoogleService(interaction);
  const sheetsContext = resolve<SheetsContextService>(interaction, 'sheetsContext');

  if (!googleSvc) {
    return { success: false, message: t('files.error.serviceUnavailable') };
  }

  const meta = await googleSvc.getDriveFileMetadata(options.fileId);
  const driveCfg = config.drive;
  const mime = String(meta.mimeType || '');
  if (driveCfg?.allowedMime && !isMimeAllowed(mime, driveCfg.allowedMime)) {
    return { success: false, message: t('files.policy.disallowedMime') };
  }
  if (driveCfg?.ownerAllowlist?.length) {
    const owners = (meta.owners as any[])?.map((o: any) => o?.emailAddress || o?.displayName).filter(Boolean) || [];
    if (!isOwnerAllowed(owners, driveCfg.ownerAllowlist)) {
      return { success: false, message: t('files.policy.deniedOwner') };
    }
  }

  const sizeBytes = Number(meta.size || 0) || 0;
  const tooLarge = isTooLarge(sizeBytes, (driveCfg?.fileMaxSizeMb ?? 0));

  if (tooLarge && !mime.startsWith('application/vnd.google-apps')) {
    const linkAllowed = !(driveCfg?.hideWebLink);
    const link = linkAllowed ? String(meta.webViewLink || '') : '';
    const sizeMb = (sizeBytes / (1024 * 1024)).toFixed(1);
    const summary = t('files.summary.largeFile', {
      name: String(meta.name || ''),
      mimeType: String(meta.mimeType || ''),
      size: sizeMb,
    });
    const linkText = linkAllowed && link ? `\n${t('files.summary.link')}: ${link}` : '';
    return { success: true, message: `${summary}${linkText}` };
  }

  // Отримуємо контекст
  let contextText = '';
  try {
    if (googleSvc) {
      if (meta.mimeType === 'application/vnd.google-apps.document') {
        const buf = await googleSvc.exportDriveFile(options.fileId, 'text/plain');
        contextText = buf.toString('utf8').slice(0, 4000);
      } else if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
        const buf = await googleSvc.exportDriveFile(options.fileId, 'text/csv');
        contextText = buf.toString('utf8').slice(0, 4000);
      } else {
        contextText = `File: ${meta.name} (${meta.mimeType})`;
      }
    }
  } catch {}

  let sheetCtxNote = '';
  try {
    if (sheetsContext) {
      const ctx = await (sheetsContext as any).get?.('current');
      if (ctx) sheetCtxNote = `\nContext: ${JSON.stringify(ctx).slice(0, 500)}`;
    }
  } catch {}

  let analysis = `Тип аналізу: ${analysisTypeName}\n${sheetCtxNote}`;
  if (ai) {
    try {
      const res = await (ai as any).generate?.(
        `Проаналізуй наступний вміст та надай ${analysisTypeName}:\n\n${contextText}`,
        { maxTokens: 512 }
      );
      if (res && typeof res.content === 'string') {
        analysis = res.content;
      }
    } catch {
      analysis = `${analysis}\n\nЗведення (локальне): ${contextText.slice(0, 800)}`;
    }
  } else {
    analysis = `${analysis}\n\nЗведення (локальне): ${contextText.slice(0, 800)}`;
  }

  const allowLink = !(driveCfg?.hideWebLink);
  const viewLink = allowLink ? String(meta.webViewLink || '') : '';
  const linkNote = allowLink && viewLink ? `\n${t('files.summary.link') || 'Посилання'}: ${viewLink}` : '';

  return { success: true, message: `🤖 **AI-аналіз файлу**\n\n${analysis}${linkNote}` };
}
