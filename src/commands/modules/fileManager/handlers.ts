import type { ButtonInteraction } from 'discord.js';
import { t } from '@/i18n';
import logger from '@/utils/logger';
import { handleAnalyze as analyzeModule } from '@/commands/modules/fileManager/analyzers';

export type DriveAction = 'open' | 'download' | 'summary' | 'question';

export interface DriveDeps {
  config: unknown;
  getGoogleService: (interaction: any) => any;
  isMimeAllowed: (...args: any[]) => boolean;
  isOwnerAllowed: (...args: any[]) => boolean;
  isTooLarge: (...args: any[]) => boolean;
  getAnalysisTypeName: (x: any) => string;
  resolve: <T = unknown>(interaction: any, name: string) => T | undefined;
}

export async function handleDriveAction(
  interaction: ButtonInteraction,
  action: DriveAction,
  id: string,
  deps: DriveDeps
): Promise<void> {
  if (!id) return;
  switch (action) {
    case 'open':
      return handleDriveOpen(interaction, id);
    case 'download':
      return handleDriveDownload(interaction, id);
    case 'summary':
      return handleDriveSummary(interaction, id, deps);
    case 'question':
      return handleDriveQuestion(interaction, id, deps);
  }
}

// --- Internal helpers ---
function isDeferred(i: ButtonInteraction): boolean {
  return Boolean(i.deferred || i.replied);
}

async function ensureDeferred(i: ButtonInteraction): Promise<void> {
  if (!isDeferred(i)) await i.deferReply({ ephemeral: true });
}

async function replyOrFollow(i: ButtonInteraction, content: string): Promise<void> {
  if (!isDeferred(i)) {
    await i.reply({ content, ephemeral: true });
  } else {
    await i.followUp({ content, ephemeral: true });
  }
}

async function handleDriveOpen(interaction: ButtonInteraction, id: string): Promise<void> {
  const viewLink = `https://drive.google.com/file/d/${id}/view`;
  return replyOrFollow(interaction, `🔗 Відкрити файл: ${viewLink}`);
}

async function handleDriveDownload(interaction: ButtonInteraction, id: string): Promise<void> {
  const dlLink = `https://drive.google.com/uc?export=download&id=${id}`;
  return replyOrFollow(interaction, `📥 Завантажити файл: ${dlLink}`);
}

async function handleDriveSummary(
  interaction: ButtonInteraction,
  id: string,
  deps: DriveDeps
): Promise<void> {
  await ensureDeferred(interaction);
  try {
    const result = await analyzeModule(
      interaction as any,
      { fileId: id, analysisType: 'summary' } as any,
      {
        config: deps.config,
        getGoogleService: deps.getGoogleService,
        isMimeAllowed: deps.isMimeAllowed,
        isOwnerAllowed: deps.isOwnerAllowed,
        isTooLarge: deps.isTooLarge,
        getAnalysisTypeName: deps.getAnalysisTypeName,
        resolve: deps.resolve,
      }
    );
    const msg = (result as any)?.message || t('files.error.process');
    await interaction.editReply({ content: msg });
  } catch {
    await interaction.editReply({ content: t('files.error.process') });
  }
}

function getLimits() {
  return {
    ctxChars: Number(process.env['DRIVE_QA_MAX_CONTEXT_CHARS'] ?? '4000'),
    qaTokens: Number(process.env['DRIVE_QA_MAX_TOKENS'] ?? process.env['AI_MAX_TOKENS'] ?? '512'),
    ctxTokens: Number(process.env['RAG_MAX_CONTEXT_TOKENS'] ?? '1200'),
  };
}

async function tryRagAnswer(interaction: ButtonInteraction, id: string, deps: DriveDeps) {
  const { ctxTokens, qaTokens } = getLimits();
  const rag = deps.resolve<any>(interaction as any, 'rag');
  if (rag && typeof rag.answer === 'function') {
    const q = 'Коротко відповідай на основні питання, які може мати користувач щодо цього файлу.';
    const ans = await rag.answer(
      `${q}\nID: ${id}`,
      { filters: { fileId: [id] } },
      { maxTokens: ctxTokens },
      { maxTokens: qaTokens }
    );
    const text = (ans && (ans.text || ans.content || ans.answer)) || t('files.error.process');
    return `💬 ${text}`;
  }
  return undefined;
}

async function buildContextFromIndex(interaction: ButtonInteraction, id: string, deps: DriveDeps): Promise<string> {
  const { ctxChars } = getLimits();
  const searchIndex = deps.resolve<any>(interaction as any, 'searchIndex');
  if (searchIndex && typeof searchIndex.search === 'function') {
    try {
      const { hits } = await searchIndex.search({ text: '*', limit: 6, filters: { fileId: [id] } });
      return (hits || [])
        .map((h: any, i: number) => `(${i + 1}) ${h.name || ''} [${h.fileId}]\n${h.snippet || ''}`)
        .join('\n\n')
        .slice(0, ctxChars);
    } catch {}
  }
  return '';
}

async function buildContextFromExport(interaction: ButtonInteraction, id: string, deps: DriveDeps): Promise<string> {
  const { ctxChars } = getLimits();
  const googleSvc = deps.getGoogleService(interaction as any);
  if (!googleSvc) return '';
  try {
    const meta = await googleSvc.getDriveFileMetadata(id);
    if (meta?.mimeType === 'application/vnd.google-apps.document') {
      const buf = await googleSvc.exportDriveFile(id, 'text/plain');
      return String(buf?.toString?.('utf8') || '').slice(0, ctxChars);
    } else if (meta?.mimeType === 'application/vnd.google-apps.spreadsheet') {
      const buf = await googleSvc.exportDriveFile(id, 'text/csv');
      return String(buf?.toString?.('utf8') || '').slice(0, ctxChars);
    }
  } catch {}
  return '';
}

async function generateFromAI(interaction: ButtonInteraction, id: string, contextText: string, deps: DriveDeps): Promise<string | undefined> {
  const { qaTokens } = getLimits();
  const ai = deps.resolve<any>(interaction as any, 'ai');
  if (ai && typeof ai.generateResponse === 'function') {
    const prompt = `Відповідай на ключові питання по файлу (ID: ${id}). Використай наданий контекст, якщо він є.\n\n${contextText}`;
    const res = await ai.generateResponse(prompt, { maxTokens: qaTokens, useCache: false });
    const text = (res && (res.content || res.text)) || t('files.error.process');
    return `💬 ${text}`;
  }
  return undefined;
}

async function handleDriveQuestion(
  interaction: ButtonInteraction,
  id: string,
  deps: DriveDeps
): Promise<void> {
  await ensureDeferred(interaction);
  try {
    // 1) Try RAG
    const ragText = await tryRagAnswer(interaction, id, deps);
    if (ragText) {
      await interaction.editReply({ content: ragText });
      return;
    }

    // 2) Try index context
    let contextText = await buildContextFromIndex(interaction, id, deps);

    // 3) Fallback export if empty
    if (!contextText) contextText = await buildContextFromExport(interaction, id, deps);

    // 4) Generate via AI
    const aiText = await generateFromAI(interaction, id, contextText, deps);
    if (aiText) {
      await interaction.editReply({ content: aiText });
      return;
    }

    await interaction.editReply({ content: t('files.error.serviceUnavailable') });
  } catch (e) {
    logger.error('drive:question error', { error: String(e) });
    await interaction.editReply({ content: t('files.error.process') });
  }
}

export interface TextDeps {
  sessions: Map<string, { fileName: string; chunks: string[]; link?: string }>;
  buildTextPage: (args: { sid: string; page: number; fileName: string; chunks: string[]; link?: string }) => {
    embed: any;
    components: Array<any>;
  };
}

export async function handleTextAction(
  interaction: ButtonInteraction,
  txtParsed: { sid: string; page: number; action?: 'close' },
  deps: TextDeps
): Promise<void> {
  const { sid, page, action } = txtParsed;
  const session = deps.sessions.get(sid);
  if (!session) {
    await interaction.reply({ content: t('doc.sessionExpired'), ephemeral: true });
    return;
  }
  if (action === 'close') {
    deps.sessions.delete(sid);
    if (interaction.deferred || interaction.replied) {
      await interaction.editReply({ components: [] });
    } else {
      await interaction.update({ components: [] });
    }
    return;
  }
  const args: { sid: string; page: number; fileName: string; chunks: string[]; link?: string } = { sid, page, fileName: session.fileName, chunks: session.chunks };
  if (session.link) args.link = session.link;
  const { embed, components } = deps.buildTextPage(args);
  if (interaction.deferred || interaction.replied) {
    await interaction.editReply({ embeds: [embed], components });
  } else {
    await interaction.update({ embeds: [embed], components });
  }
}

export interface SearchDeps {
  sessions: Map<string, { changesOnly: boolean; baseline: number }>;
  buildSearchPage: (args: { interaction: any; sid: string; page: number }) => Promise<{
    embed: any;
    components: Array<any>;
  }>;
}

export async function handleSearchAction(
  interaction: ButtonInteraction,
  parsed: { sid: string; page: number; action?: 'toggle' | 'reset' | 'close' },
  deps: SearchDeps
): Promise<void> {
  const { sid, page, action } = parsed;
  const session = deps.sessions.get(sid);
  if (!session) {
    await interaction.reply({ content: t('doc.sessionExpired'), ephemeral: true });
    return;
  }
  if (action === 'close') {
    deps.sessions.delete(sid);
    if (interaction.deferred || interaction.replied) {
      await interaction.editReply({ components: [] });
    } else {
      await interaction.update({ components: [] });
    }
    return;
  }
  if (action === 'toggle') {
    session.changesOnly = !session.changesOnly;
  } else if (action === 'reset') {
    session.baseline = Math.floor(Date.now() / 1000);
  }
  const { embed, components } = await deps.buildSearchPage({ interaction: interaction as any, sid, page });
  if (interaction.deferred || interaction.replied) {
    await interaction.editReply({ embeds: [embed], components });
  } else {
    await interaction.update({ embeds: [embed], components });
  }
}
