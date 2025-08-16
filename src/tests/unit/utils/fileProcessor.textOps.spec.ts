import { normalizeText, sanitizeTextForChat, chunkTextForDiscord, buildPaginatedChunks, summarizeTlDr, cleanupFileProcessor } from '@/utils/fileProcessor';

describe('fileProcessor text operations', () => {
  afterAll(() => {
    // зупиняємо інтервал очищення, який стартує в singleton при імпорті
    cleanupFileProcessor();
  });
  it('normalizeText removes zero-width, normalizes newlines and tabs', () => {
    const input = 'A\r\nB\u200B\tC\n\n\nD';
    const out = normalizeText(input);
    expect(out).toBe('A\nB  C\n\nD');
  });

  it('sanitizeTextForChat trims and respects maxLen with ellipsis', () => {
    const input = 'x'.repeat(2000);
    const out = sanitizeTextForChat(input, 180);
    expect(out.length).toBeLessThanOrEqual(181); // may end with ellipsis
    expect(out.endsWith('…')).toBe(true);
  });

  it('chunkTextForDiscord splits large text within limit and on boundaries', () => {
    const para = Array.from({ length: 50 }, (_, i) => `Sentence ${i + 1}.`).join(' ');
    const text = `${para}\n\n${para}\n\n${para}`;
    const chunks = chunkTextForDiscord(text, { maxChunkLen: 300 });
    expect(chunks.length).toBeGreaterThan(1);
    expect(chunks.every(c => c.length <= 300)).toBe(true);
  });

  it('buildPaginatedChunks appends footer i/N', () => {
    const text = 'Para1. '.repeat(1200);
    const parts = buildPaginatedChunks(text, { maxChunkLen: 400 });
    if (parts.length > 1) {
      const last = parts[parts.length - 1];
      expect(last).toMatch(new RegExp(`_${parts.length}/${parts.length}_$`));
    }
  });

  it('summarizeTlDr returns compact summary within budget', () => {
    const long = 'Intro. ' + Array.from({ length: 200 }, (_, i) => `Sentence number ${i + 1} is informative.`).join(' ');
    const budget = 500;
    const summary = summarizeTlDr(long, { budget });
    expect(summary.length).toBeLessThanOrEqual(budget + 1); // +ellipsis possible
    expect(summary).toMatch(/[A-Za-zА-Яа-яІіЇїЄє0-9]/);
  });
});
