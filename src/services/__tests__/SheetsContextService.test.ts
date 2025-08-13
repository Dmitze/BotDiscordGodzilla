import { jest } from '@jest/globals';
import { SheetsContextService } from '@/services/SheetsContextService';
import type { BotConfig } from '@/types';

const makeConfig = (ttl = 300): BotConfig => ({
  // минимальная тестовая конфигурация
  performance: { cacheTTL: ttl } as any,
} as unknown as BotConfig);

describe('SheetsContextService', () => {
  let service: SheetsContextService;

  beforeAll(async () => {
    service = new SheetsContextService(makeConfig(120));
    await service.initialize();
  });

  afterAll(async () => {
    await service.shutdown();
  });

  test('set/get by guild/channel/user with priority user > channel > guild', async () => {
    // базовый контекст на уровне guild
    await service.setContext({ guildId: 'g1' }, { spreadsheetId: 's-g', sheetName: 'G' });
    let ctx = await service.getContext({ guildId: 'g1' });
    expect(ctx).not.toBeNull();
    expect(ctx?.spreadsheetId).toBe('s-g');

    // перекрываем на уровне channel
    await service.setContext({ channelId: 'c1' }, { spreadsheetId: 's-c', sheetName: 'C' });
    ctx = await service.getContext({ guildId: 'g1', channelId: 'c1' });
    expect(ctx?.spreadsheetId).toBe('s-c');

    // перекрываем на уровне user
    await service.setContext({ userId: 'u1' }, { spreadsheetId: 's-u', sheetName: 'U' });
    ctx = await service.getContext({ guildId: 'g1', channelId: 'c1', userId: 'u1' });
    expect(ctx?.spreadsheetId).toBe('s-u');
  });

  test('clearContext removes entries for provided scope keys', async () => {
    await service.setContext({ guildId: 'g2' }, { spreadsheetId: 'S1' });
    await service.setContext({ guildId: 'g2', channelId: 'c2' } as any, { spreadsheetId: 'S2' });

    let ctx = await service.getContext({ guildId: 'g2', channelId: 'c2' });
    expect(ctx?.spreadsheetId).toBe('S2');

    // Удаляем только по guild — канал останется
    const removedGuild = await service.clearContext({ guildId: 'g2' });
    expect(removedGuild).toBe(true);
    ctx = await service.getContext({ guildId: 'g2', channelId: 'c2' });
    expect(ctx?.spreadsheetId).toBe('S2');

    // Теперь очищаем по channel — запись исчезнет
    const removedChannel = await service.clearContext({ channelId: 'c2' });
    expect(removedChannel).toBe(true);
    ctx = await service.getContext({ guildId: 'g2', channelId: 'c2' });
    expect(ctx).toBeNull();
  });

  test('TTL expiration returns null after expiry', async () => {
    const realNow = Date.now;
    const base = realNow();
    jest.spyOn(Date, 'now').mockImplementation(() => base);

    await service.setContext({ channelId: 'ttl' }, { spreadsheetId: 'TTL1' }, 1); // 1s TTL

    // До истечения
    let ctx = await service.getContext({ channelId: 'ttl' });
    expect(ctx).not.toBeNull();

    // Сервис применяет минимальный TTL 30 секунд — перемещаемся дальше
    ;(Date.now as jest.Mock).mockImplementation(() => base + 31_000);

    ctx = await service.getContext({ channelId: 'ttl' });
    expect(ctx).toBeNull();

    // Восстанавливаем Date.now
    ;(Date.now as jest.Mock).mockRestore();
  });
});
