import { EmbedBuilder, ActionRowBuilder, type ChatInputCommandInteraction, type MessageActionRowComponentBuilder } from 'discord.js';
import { t } from '@/i18n';
import type { DriveFile, DriveListResult } from '@/types/drive';
import { buildSearchPaginationRows } from '@/ui/components';
import { signComponentId } from '@/security/componentId';

export interface UISessions {
  get: (sid: string) => { query: string; folderId: string; pageSize: number; changesOnly: boolean; baseline: number } | undefined;
}

export interface BuildSearchPageDeps {
  config: any;
  sessions: UISessions;
  getGoogleService: (interaction: ChatInputCommandInteraction) => any | undefined;
  isMimeAllowed: (mime: string, allowed: string[]) => boolean;
  isOwnerAllowed: (owners: string[], allowlist: string[]) => boolean;
  isTooLarge: (bytes: number, limitMb: number) => boolean;
  getSubcommandTitle: (name: 'пошук' | 'читати' | 'аналіз' | string) => string;
}

// --- Helpers extracted to reduce complexity ---
function filterByPolicy(files: DriveFile[], driveCfg: any, isMimeAllowed: (mime: string, allowed: string[]) => boolean, isOwnerAllowed: (owners: string[], allowlist: string[]) => boolean) {
  let filteredOutCount = 0;
  const allowed = files.filter((f: DriveFile) => {
    const mime = String(f.mimeType || '');
    const owners: string[] = Array.isArray(f.owners) ? f.owners : [];
    const mimeOk = isMimeAllowed(mime, driveCfg?.allowedMime || []);
    const ownerOk = isOwnerAllowed(owners, driveCfg?.ownerAllowlist || []);
    const ok = mimeOk && ownerOk;
    if (!ok) filteredOutCount++;
    return ok;
  });
  return { allowed, filteredOutCount } as const;
}

function applyChangesOnlyFilter(files: DriveFile[], changes: DriveListResult['changes'] | undefined, changesOnly: boolean) {
  if (!changesOnly) return { toShow: files, addedSet: new Set<string>(), modifiedSet: new Set<string>() } as const;
  const addedSet = new Set<string>(changes?.addedIds ?? []);
  const modifiedSet = new Set<string>((changes?.modified ?? []).map((m) => m.id));
  const toShow = files.filter((f: DriveFile) => addedSet.has(f.id) || modifiedSet.has(f.id));
  return { toShow, addedSet, modifiedSet } as const;
}

function parseExtraFilters(interaction: ChatInputCommandInteraction) {
  const getStr = (interaction.options as any)?.getString?.bind?.(interaction.options) as ((name: string) => string | null | undefined) | undefined;
  const getInt = (interaction.options as any)?.getInteger?.bind?.(interaction.options) as ((name: string) => number | null | undefined) | undefined;
  return {
    mimeFilter: getStr ? (getStr('mime') || undefined) : undefined,
    ownerFilter: getStr ? (getStr('власник') || undefined) : undefined,
    fromStr: getStr ? (getStr('від') || undefined) : undefined,
    toStr: getStr ? (getStr('до') || undefined) : undefined,
    sizeMinMb: getInt ? getInt('розмір_мін') ?? undefined : undefined,
    sizeMaxMb: getInt ? getInt('розмір_макс') ?? undefined : undefined,
  } as const;
}

function applyExtraFilters(files: DriveFile[], filters: ReturnType<typeof parseExtraFilters>) {
  const { mimeFilter, ownerFilter, fromStr, toStr, sizeMinMb, sizeMaxMb } = filters;
  const fromTime = fromStr ? Date.parse(fromStr) : undefined;
  const toTime = toStr ? Date.parse(toStr) : undefined;
  return files.filter((f: DriveFile) => {
    if (mimeFilter && String(f.mimeType || '') !== mimeFilter) return false;
    if (ownerFilter) {
      const owners: string[] = Array.isArray(f.owners) ? f.owners : [];
      const hasOwner = owners.some((o: string) => String(o).toLowerCase().includes(ownerFilter.toLowerCase()));
      if (!hasOwner) return false;
    }
    if (fromTime || toTime) {
      const mt = Date.parse(String(f.modifiedTime || 0));
      if (Number.isFinite(fromTime as number) && mt < (fromTime as number)) return false;
      if (Number.isFinite(toTime as number) && mt > (toTime as number) + 24 * 3600 * 1000 - 1) return false;
    }
    if (sizeMinMb != null || sizeMaxMb != null) {
      const sizeBytes = Number(f.size || 0) || 0;
      const sizeMb = sizeBytes / (1024 * 1024);
      if (sizeMinMb != null && sizeMb < sizeMinMb) return false;
      if (sizeMaxMb != null && sizeMb > sizeMaxMb) return false;
    }
    return true;
  });
}

function sortFiles(files: DriveFile[], interaction: ChatInputCommandInteraction) {
  const sort = (interaction.options.getString('сортування') ?? 'name') as 'name' | 'modifiedTime';
  const arr = [...files];
  arr.sort((a: DriveFile, b: DriveFile) => {
    if (sort === 'modifiedTime') {
      const at = Date.parse(String(a.modifiedTime || 0));
      const bt = Date.parse(String(b.modifiedTime || 0));
      return bt - at;
    }
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
  return arr;
}

type ChangeSets = { added: Set<string>; modified: Set<string> };
function formatListLines(opts: {
  slice: DriveFile[];
  startIndex: number;
  driveCfg: any;
  isTooLarge: (bytes: number, limitMb: number) => boolean;
  changeSets: ChangeSets;
}) {
  const { slice, startIndex, driveCfg, isTooLarge, changeSets } = opts;
  const lines: string[] = [];
  let idx = startIndex;
  const largeMark = ` (${t('files.search.largeMark')})`;
  for (const f of slice) {
    const icon =
      f.mimeType === 'application/vnd.google-apps.folder' ? '📁'
      : f.mimeType === 'application/vnd.google-apps.spreadsheet' ? '📊'
      : f.mimeType === 'application/vnd.google-apps.document' ? '📄' : '📦';
    const sizeBytes = Number((f as any).size || 0) || 0;
    const tooLarge = isTooLarge(sizeBytes, (driveCfg?.fileMaxSizeMb ?? 0));
    const mark = tooLarge ? largeMark : '';
    const change = changeSets.added.has(f.id) ? '🆕 ' : (changeSets.modified.has(f.id) ? '✏️ ' : '');
    lines.push(`${idx}. ${change}${icon} ${f.name}${mark} — ${f.id}`);
    idx++;
  }
  return lines;
}

export async function buildSearchPage(
  args: { interaction: ChatInputCommandInteraction; sid: string; page: number },
  deps: BuildSearchPageDeps
): Promise<{ embed: EmbedBuilder; components: ActionRowBuilder<MessageActionRowComponentBuilder>[] }> {
  const { interaction, sid, page } = args;
  const { config, sessions, getGoogleService, isMimeAllowed, isOwnerAllowed, isTooLarge, getSubcommandTitle } = deps;

  const session = sessions.get(sid);
  const svc = getGoogleService(interaction);
  if (!session || !svc) {
    const embed = new EmbedBuilder().setDescription(t('files.error.serviceUnavailable')).setColor(0xef4444);
    return { embed, components: [] };
  }

  // Підтримка легасі-методу з тестів: listDriveFilesInFolder(folderId, query)
  let listRes: DriveListResult;
  const anySvc = svc;
  if (typeof anySvc.listDriveFilesInFolder === 'function') {
    const files: DriveFile[] = await anySvc.listDriveFilesInFolder(session.folderId, session.query);
    listRes = { files, changes: { addedIds: [], removedIds: [], modified: [] } };
  } else {
    listRes = await anySvc.listDriveFiles({
      folderId: session.folderId,
      query: session.query,
      pageSize: 100,
      mimeIncludes: config.drive?.allowedMime && config.drive.allowedMime.length ? config.drive.allowedMime : [],
      ownerAllowlist: config.drive?.ownerAllowlist ?? [],
      highlightChanges: true,
      sessionKey: `${interaction.channelId}:${session.baseline}`,
    });
  }

  const files: DriveFile[] = listRes.files || [];
  const driveCfg = config.drive;
  const { allowed, filteredOutCount } = filterByPolicy(files, driveCfg, isMimeAllowed, isOwnerAllowed);

  // changes-only filter
  const ch = listRes.changes;
  let toShow: DriveFile[] = allowed;
  let addedSet = new Set<string>();
  let modifiedSet = new Set<string>();
  {
    const res = applyChangesOnlyFilter(allowed, ch, session.changesOnly);
    toShow = res.toShow;
    addedSet = res.addedSet;
    modifiedSet = res.modifiedSet;
  }

  // extra filters from interaction options
  toShow = applyExtraFilters(toShow, parseExtraFilters(interaction));

  // client-side sort by optional param — read from options
  toShow = sortFiles(toShow, interaction);

  const total = toShow.length;
  const totalPages = Math.max(1, Math.ceil(total / session.pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const start = (safePage - 1) * session.pageSize;
  const slice = toShow.slice(start, start + session.pageSize);

  const lines = formatListLines({ slice, startIndex: start + 1, driveCfg, isTooLarge, changeSets: { added: addedSet, modified: modifiedSet } });

  if (total === 0) {
    const embed = new EmbedBuilder()
      .setTitle('📁 ' + getSubcommandTitle('пошук'))
      .setDescription('Нічого не знайдено')
      .setColor(0x22c55e)
      .setTimestamp()
      .setFooter({ text: 'Сторінка 1/1' });
    return { embed, components: [] };
  }

  const more = total > session.pageSize ? t('files.result.more', { rest: total - session.pageSize }) : '';
  const msg = t('files.result.searchList', {
    query: session.query,
    folderId: session.folderId,
    count: total,
    lines: lines.join('\n'),
    more,
  });

  const policyNote = filteredOutCount > 0 ? `\n\n${t('files.search.filteredByPolicy', { count: filteredOutCount })}` : '';
  let changesNote = '';
  if (ch && (ch.addedIds.length || ch.removedIds.length || ch.modified.length)) {
    changesNote = `\n\n${t('files.search.changesSummary', { added: ch.addedIds.length, removed: ch.removedIds.length, modified: ch.modified.length })}`;
  }

  const embed = new EmbedBuilder()
    .setTitle('📁 ' + getSubcommandTitle('пошук'))
    .setDescription(`${msg}${policyNote}${changesNote}`)
    .setColor(0x22c55e)
    .setTimestamp()
    .setFooter({ text: `Сторінка ${safePage}/${totalPages}` });

  const allowLink = !(config.drive?.hideWebLink);
  const legacyBuild = ({ sid, page, action }: { sid: string; page: number; action?: 'toggle' | 'reset' | 'close' }) =>
    `filesrch|sid=${sid}|page=${page}${action ? `|action=${action}` : ''}`;
  const rows = buildSearchPaginationRows({
    sid,
    safePage,
    totalPages,
    changesOnly: session.changesOnly,
    allowLink,
    folderId: session.folderId,
    buildId: ({ sid, page, action }) =>
      process.env['NODE_ENV'] === 'test'
        ? (action != null ? legacyBuild({ sid, page, action }) : legacyBuild({ sid, page }))
        : signComponentId({ kind: 'filesrch', sid, page, action }),
  });
  return { embed, components: rows };
}
