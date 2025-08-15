import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import DriveIndexerService from '@/services/DriveIndexerService';

export interface GoogleMock {
  listDriveFiles: jest.Mock<Promise<{ files: DriveFile[]; nextPageToken?: string }>, any>;
  extractTextFromFile: jest.Mock<Promise<string>, any>;
  getDriveFile: jest.Mock<Promise<DriveFile>, any>;
}

export interface CacheMock {
  store: Map<string, any>;
  get: <T = any>(key: string) => Promise<T | null>;
  set: (key: string, value: any, _ttl?: number) => Promise<void>;
}

export interface MetricsMock {
  incCounter: jest.Mock<void, any>;
  observeHistogram: jest.Mock<void, any>;
}

export function createCacheMock(): CacheMock {
  const store = new Map<string, any>();
  return {
    store,
    async get<T>(key: string): Promise<T | null> {
      return store.has(key) ? (store.get(key) as T) : null;
    },
    async set(key: string, value: any): Promise<void> {
      store.set(key, value);
    },
  };
}

export function createGoogleMock(): GoogleMock {
  return {
    listDriveFiles: jest.fn(),
    extractTextFromFile: jest.fn(),
    getDriveFile: jest.fn(),
  };
}

export function createMetricsMock(): MetricsMock {
  return {
    incCounter: jest.fn(),
    observeHistogram: jest.fn(),
  };
}

export function createBotMock(
  google: GoogleMock,
  cache: CacheMock,
  options?: { config?: Partial<BotConfig>; services?: Record<string, any> }
) {
  const baseConfig: BotConfig = {
    env: 'test',
    discord: { token: 'x', clientId: 'x', guildId: 'x', enableSlash: false, enableChat: false, enableMessageContentIntent: false },
    google: { apiKey: 'x', driveFolderId: 'root' },
    drive: { enableTextIndex: true, folderId: 'root', ttlTextSec: 3600, indexCron: '*/30 * * * *' },
    ...(options?.config as any),
  } as BotConfig;

  const services: Record<string, any> = {
    google,
    cache,
    scheduler: { scheduleJob: jest.fn() },
    ...(options?.services || {}),
  };

  return {
    config: baseConfig,
    getService(name: string) {
      return services[name];
    },
  };
}

export async function initIndexer(bot: any) {
  const indexer = new DriveIndexerService(bot);
  await indexer.initialize();
  return indexer;
}
