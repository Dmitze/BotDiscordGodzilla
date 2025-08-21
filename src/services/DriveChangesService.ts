import { google, type drive_v3 } from 'googleapis';
import type { BotConfig } from '@/types';
import { CacheService } from './CacheService';
import logger from '@/utils/logger';

export type DriveChangeEvent = {
  fileId: string;
  name?: string;
  time?: string;
  type: 'created' | 'modified' | 'removed';
  owners?: string[];
  webViewLink?: string;
};

export interface IDriveChangesProvider {
  getStartPageToken(): Promise<string>;
  listChanges(pageToken: string): Promise<{ changes: drive_v3.Schema$Change[]; newStartPageToken?: string; nextPageToken?: string }>; 
}

class GoogleDriveChangesProvider implements IDriveChangesProvider {
  private drive: drive_v3.Drive;

  constructor(config: BotConfig) {
    const credentials = config.google.credentials;
    if (!credentials) throw new Error('Google credentials are not configured');
    const scopes = ['https://www.googleapis.com/auth/drive.readonly'];
    const auth = new google.auth.JWT({
      email: credentials.client_email,
      key: credentials.private_key,
      scopes,
      subject: undefined,
    } as any);
    this.drive = google.drive({ version: 'v3', auth });
  }

  async getStartPageToken(): Promise<string> {
    const res = await this.drive.changes.getStartPageToken({});
    return String(res.data.startPageToken || '');
    }

  async listChanges(pageToken: string): Promise<{ changes: drive_v3.Schema$Change[]; newStartPageToken?: string; nextPageToken?: string }> {
    const res = await this.drive.changes.list({ pageToken, fields: '*', pageSize: 100 });
    const out: { changes: drive_v3.Schema$Change[]; newStartPageToken?: string; nextPageToken?: string } = {
      changes: Array.isArray(res.data.changes) ? res.data.changes : [],
    };
    if (res.data.newStartPageToken) out.newStartPageToken = res.data.newStartPageToken;
    if (res.data.nextPageToken) out.nextPageToken = res.data.nextPageToken;
    return out;
  }
}

// No-op provider for tests/disabled mode
class NoOpDriveChangesProvider implements IDriveChangesProvider {
  async getStartPageToken(): Promise<string> {
    return 'noop-token';
  }
  async listChanges(_pageToken: string): Promise<{ changes: drive_v3.Schema$Change[]; newStartPageToken?: string; nextPageToken?: string }> {
    return { changes: [] };
  }
}

export class DriveChangesService {
  private cache: CacheService;
  private provider: IDriveChangesProvider;
  private folderId: string;
  private hideWebLink: boolean;
  private keyToken = 'drive:changes:startPageToken';
  private disabled: boolean;

  constructor(config: BotConfig, provider?: IDriveChangesProvider, cache?: CacheService) {
    this.cache = cache ?? new CacheService(config);
    const driveCfg = (config as Partial<BotConfig>).drive as Partial<BotConfig['drive']> | undefined;
    this.folderId = (driveCfg?.folderId as string | undefined) ?? '';
    this.hideWebLink = (driveCfg?.hideWebLink as boolean | undefined) ?? true;
    // Determine disabled mode for tests or when flagged/missing credentials.
    // IMPORTANT: if a custom provider is supplied (e.g., in tests), do NOT disable.
    const envDisabled =
      process.env['NODE_ENV'] === 'test' ||
      process.env['DISABLE_DRIVE_CHANGES'] === 'true' ||
      !config.google?.credentials;
    this.disabled = envDisabled && !provider;
    this.provider = provider ?? (this.disabled ? new NoOpDriveChangesProvider() : new GoogleDriveChangesProvider(config));
  }

  async initialize(): Promise<void> {
    if (this.disabled) {
      logger.info('🧪 DriveChangesService: тест/відключено/немає credentials — ініціалізацію зовнішніх ресурсів пропущено');
      return;
    }
    // ensure start token exists
    const t = await this.getStoredToken();
    if (!t) {
      await this.refreshStartPageToken();
    }
  }

  private async getStoredToken(): Promise<string | null> {
    try {
      const token = await this.cache.get<string>(this.keyToken);
      return token ?? null;
    } catch {
      return null;
    }
  }

  private async setStoredToken(token: string): Promise<void> {
    try {
      // store for a week
      await this.cache.set(this.keyToken, token, 7 * 24 * 3600);
    } catch (e) {
      logger.warn('DriveChangesService: cannot persist token', e as any);
    }
  }

  async refreshStartPageToken(): Promise<string> {
    const token = await this.provider.getStartPageToken();
    if (!token) throw new Error('Failed to get startPageToken');
    await this.setStoredToken(token);
    return token;
  }

  private mapChange(c: drive_v3.Schema$Change): DriveChangeEvent | null {
    const file = c.file;
    if (!file) return null;
    const removed = !!c.removed || file.trashed === true;
    const type: DriveChangeEvent['type'] = removed ? 'removed' : c.time && file.createdTime === c.time ? 'created' : 'modified';
    // filter by folder when applicable: when change includes file.parents
    const isInFolder = this.folderId
      ? (Array.isArray(file.parents) ? file.parents.includes(this.folderId) : true)
      : true;
    if (!isInFolder) return null;
    const owners = Array.isArray(file.owners)
      ? file.owners
          .map(o => (o?.emailAddress ? o.emailAddress : o?.displayName))
          .filter((v): v is string => typeof v === 'string' && v.length > 0)
      : undefined;
    const evt: DriveChangeEvent = {
      fileId: String(file.id || ''),
      type,
    };
    if (file.name) evt.name = file.name;
    const time = c.time || file.modifiedTime;
    if (time) evt.time = time;
    if (owners && owners.length) evt.owners = owners;
    if (!this.hideWebLink && file.webViewLink) evt.webViewLink = file.webViewLink as string;
    return evt;
  }

  async pollOnce(): Promise<{ events: DriveChangeEvent[]; newToken?: string }> {
    if (this.disabled) {
      // Respect exactOptionalPropertyTypes — omit newToken entirely in disabled mode
      return { events: [] };
    }
    let pageToken = (await this.getStoredToken()) || (await this.refreshStartPageToken());
    const events: DriveChangeEvent[] = [];
    let next: string | undefined = undefined;
    let newStart: string | undefined = undefined;

    do {
      const { changes, nextPageToken, newStartPageToken } = await this.provider.listChanges(pageToken);
      for (const ch of changes) {
        const ev = this.mapChange(ch);
        if (ev) events.push(ev);
      }
      next = nextPageToken ?? undefined;
      if (newStartPageToken) newStart = newStartPageToken;
      pageToken = nextPageToken || pageToken;
    } while (next);

    const finalToken = newStart || pageToken;
    if (finalToken) await this.setStoredToken(finalToken);
    logger.debug(`DriveChangesService: polled ${events.length} events`);
    return { events, newToken: finalToken };
  }
}
