import logger from '@/utils/logger';

describe('logger redact', () => {
  it('redacts sensitive keys in meta before buffering', () => {
    logger.info('test redact', {
      token: 'abc',
      apiKey: 'sk-secret',
      nested: { password: 'p@ss', note: 'ok' },
    });

    const buf = logger.getLogBuffer();
    const last = buf[buf.length - 1]!;
    expect(last.meta['token']).toBe('***');
    expect(last.meta['apiKey']).toBe('***');
    expect(last.meta['nested']['password']).toBe('***');
    expect(last.meta['nested']['note']).toBe('ok');
  });
});

