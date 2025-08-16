import logger from '@/utils/logger';

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
    expect((last.meta as any).password).toBe(redacted);
    expect((last.meta as any).token).toBe(redacted);
    expect((last.meta as any).apiKey).toBe(redacted);
    // nested
    expect(((last.meta as any).nested as any).secret).toBe(redacted);
  });
});
