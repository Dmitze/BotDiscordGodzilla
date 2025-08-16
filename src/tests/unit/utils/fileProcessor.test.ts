import { readFile, writeFile, getFileProcessorStats, cleanupFileProcessor } from '@/utils/fileProcessor';
import { existsSync, mkdirSync } from 'fs';

const TMP_DIR = 'data/tmp/jest';

describe('FileProcessor basic read/write and mime detection', () => {
  beforeAll(() => {
    if (!existsSync(TMP_DIR)) mkdirSync(TMP_DIR, { recursive: true });
  });

  afterAll(() => {
    cleanupFileProcessor();
  });

  it('writes and reads a small text file', async () => {
    const path = TMP_DIR + '/sample.txt';
    const content = 'hello world';
    const writeRes = await writeFile(path, content);
    expect(writeRes.success).toBe(true);
    const readRes = await readFile(path);
    expect(readRes.success).toBe(true);
    expect(readRes.fileInfo?.mimeType).toBe('text/plain');
    expect(typeof readRes.content).toBe('string');
    expect(String(readRes.content)).toContain('hello world');
  });

  it('detects mime for .pdf', async () => {
    const path = TMP_DIR + '/empty.pdf';
    const writeRes = await writeFile(path, Buffer.from('PDF'));
    expect(writeRes.success).toBe(true);
    const readRes = await readFile(path);
    expect(readRes.success).toBe(true);
    expect(readRes.fileInfo?.mimeType).toBe('application/pdf');
  });

  it('warns for disallowed extension but remains valid', async () => {
    const path = TMP_DIR + '/blob.bin';
    const writeRes = await writeFile(path, Buffer.from([1,2,3]));
    expect(writeRes.success).toBe(true);
    const readRes = await readFile(path);
    expect(readRes.success).toBe(true);
    expect(readRes.fileInfo?.warnings.some(w => w.includes('Недозволене розширення'))).toBe(true);
    expect(readRes.fileInfo?.isValid).toBe(true);
  });

  it('updates stats after operations', async () => {
    const stats = getFileProcessorStats();
    expect(stats.totalOperations).toBeGreaterThan(0);
    expect(stats.filesProcessed).toBeGreaterThan(0);
  });
});
