/* eslint-disable @typescript-eslint/no-unsafe-return */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { jest, beforeAll, describe, it, expect } from '@jest/globals';
import type * as Winston from 'winston';
import type { Logger as LoggerType } from '@/utils/logger';

let logger: LoggerType;

beforeAll(async () => {
  jest.unstable_mockModule('winston', () => {
    const mockedLogger = {
      log: jest.fn(),
      info: jest.fn(),
      error: jest.fn(),
      warn: jest.fn(),
      debug: jest.fn(),
      close: jest.fn(),
    };
    const emptyFormat = {} as unknown as Winston.Logform.Format;
    const emptyTransport = {} as unknown as Winston.transport;
    const mocked = {
      createLogger: jest.fn<() => Winston.Logger>(() => mockedLogger as unknown as Winston.Logger),
      format: {
        combine: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        timestamp: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        errors: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        json: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        colorize: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        simple: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
        printf: jest.fn<() => Winston.Logform.Format>(() => emptyFormat),
      },
      transports: {
        Console: jest.fn<() => Winston.transport>(() => emptyTransport),
        File: jest.fn<() => Winston.transport>(() => emptyTransport),
      },
    };
    return { __esModule: true, default: mocked, ...mocked };
  });

  const mod = await import('@/utils/logger');
  logger = mod.default;
});
<<<<<<< HEAD

describe('logger redact', () => {
  it('redacts sensitive fields in meta', () => {
    logger.info('Test with secrets', {
      password: 'secret',
      token: 'abc',
      apiKey: 'key',
      nested: { secret: 'shh' },
    });
    const buf = logger.getLogBuffer();
    const last = buf[buf.length - 1] as { meta: Record<string, unknown> };
    const redacted = '[REDACTED]';
    // flat fields
    const meta = last.meta;
    expect(meta['password']).toBe(redacted);
    expect(meta['token']).toBe(redacted);
    expect(meta['apiKey']).toBe(redacted);
    // nested
    const nested = meta['nested'] as Record<string, unknown>;
    expect(nested['secret']).toBe(redacted);
  });
});
=======
>>>>>>> 9c806657 (test(unit): перевірка редагування логів)
