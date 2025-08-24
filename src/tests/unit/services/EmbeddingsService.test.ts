import { EmbeddingsService } from '@/services/EmbeddingsService';

describe('EmbeddingsService (mock provider)', () => {
  const old = { ...process.env };
  afterAll(() => { process.env = old; });

  it('returns deterministic unit-length vectors', async () => {
    process.env['EMBEDDINGS_PROVIDER'] = 'mock';
    const svc = new EmbeddingsService({} as any);
    const v1 = await svc.embed('hello world');
    const v2 = await svc.embed('hello world');
    const v3 = await svc.embed('another');

    expect(v1).toHaveLength(384);
    // same input -> same output
    expect(v1).toEqual(v2);
    // different input -> different output (very likely)
    expect(v1).not.toEqual(v3);

    // Check unit length approximately
    const norm = Math.sqrt(v1.reduce((s, x) => s + x * x, 0));
    expect(norm).toBeGreaterThan(0.99);
    expect(norm).toBeLessThan(1.01);
  });
});
