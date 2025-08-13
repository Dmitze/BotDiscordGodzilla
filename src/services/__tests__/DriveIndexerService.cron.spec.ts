import { describe, it, expect, beforeEach, jest, afterEach } from '@jest/globals';
import DriveIndexerService from '../DriveIndexerService';

function makeBot(overrides: Partial<any> = {}) {
  const config = {
    drive: {
      enableTextIndex: true,
      indexCron: '*/5 * * * *',
    },
    google: {},
  } as any;
  const services: Record<string, any> = {
    google: { extractTextFromFile: jest.fn(), listDriveFiles: jest.fn() },
    cache: { get: jest.fn(), set: jest.fn() },
    searchIndex: { upsert: jest.fn(), search: jest.fn() },
    metrics: { incCounter: jest.fn(), observeHistogram: jest.fn() },
    scheduler: { scheduleJob: jest.fn() },
    ...(overrides.services || {}),
  };
  return {
    config: { ...config, ...(overrides.config || {}) },
    getService: (name: string) => services[name],
  } as any;
}

const env = process.env;

describe('DriveIndexerService cron registration', () => {
  beforeEach(() => {
    jest.resetModules();
    process.env = { ...env };
  });
  afterEach(() => {
    process.env = env;
  });

  it('does not register cron when NODE_ENV=test', async () => {
    process.env.NODE_ENV = 'test';
    const bot = makeBot();
    const svc = new DriveIndexerService(bot);
    await svc.initialize();
    const scheduler = bot.getService('scheduler');
    expect(scheduler.scheduleJob).not.toHaveBeenCalled();
  });

  it('does not register cron when DISABLE_CRON=true', async () => {
    process.env.NODE_ENV = 'production';
    process.env.DISABLE_CRON = 'true';
    const bot = makeBot();
    const svc = new DriveIndexerService(bot);
    await svc.initialize();
    const scheduler = bot.getService('scheduler');
    expect(scheduler.scheduleJob).not.toHaveBeenCalled();
  });

  it('registers cron in non-test when enabled', async () => {
    process.env.NODE_ENV = 'production';
    delete process.env.DISABLE_CRON;
    const bot = makeBot();
    const svc = new DriveIndexerService(bot);
    await svc.initialize();
    const scheduler = bot.getService('scheduler');
    expect(scheduler.scheduleJob).toHaveBeenCalledWith(
      'drive-index',
      expect.any(String),
      expect.any(Function)
    );
  });
});
