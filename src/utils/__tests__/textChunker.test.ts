import { chunkTextByTokens, chunkTextBySentences } from '../textChunker';
import { countTokens } from '../token';

// Mock the tiktoken library
jest.mock('../token', () => ({
  countTokens: jest.fn((text: string) => Math.ceil(text.length / 4))
}));

describe('textChunker', () => {
  beforeEach(() => {
    (countTokens as jest.Mock).mockClear();
  });

  describe('chunkTextByTokens', () => {
    it('should return empty array for empty text', () => {
      const result = chunkTextByTokens('');
      expect(result).toEqual([]);
    });

    it('should chunk text into multiple parts', () => {
      // Mock token counting to return specific values
      (countTokens as jest.Mock)
        .mockImplementationOnce(() => 500) // First chunk
        .mockImplementationOnce(() => 600); // Second chunk

      const text = 'a'.repeat(4000); // 1000 tokens approximately
      const result = chunkTextByTokens(text, 500, 400, 600, 50);

      expect(result.length).toBeGreaterThan(1);
      expect(result[0]).toBeDefined();
      if (result[0]) {
        expect(result[0].tokenCount).toBe(500);
      }
      expect(result[1]).toBeDefined();
      if (result[1]) {
        expect(result[1].tokenCount).toBe(600);
      }
    });

    it('should respect token limits', () => {
      // Mock token counting
      (countTokens as jest.Mock)
        .mockImplementationOnce(() => 800) // Within limits
        .mockImplementationOnce(() => 1200); // At max limit

      const text = 'test '.repeat(1000);
      const result = chunkTextByTokens(text, 1000, 800, 1200, 100);

      expect(result.length).toBeGreaterThan(1);
      expect(result[0]).toBeDefined();
      if (result[0]) {
        expect(result[0].tokenCount).toBeGreaterThanOrEqual(800);
        expect(result[0].tokenCount).toBeLessThanOrEqual(1200);
      }
      expect(result[1]).toBeDefined();
      if (result[1]) {
        expect(result[1].tokenCount).toBeLessThanOrEqual(1200);
      }
    });
  });

  describe('chunkTextBySentences', () => {
    it('should return empty array for empty text', () => {
      const result = chunkTextBySentences('');
      expect(result).toEqual([]);
    });

    it('should chunk text by sentences', () => {
      (countTokens as jest.Mock)
        .mockImplementation((text: string) => text.split('.').length);

      const text = 'First sentence. Second sentence. Third sentence. Fourth sentence.';
      const result = chunkTextBySentences(text, 10);

      expect(result.length).toBeGreaterThan(0);
      // Check that chunks contain sentences
      result.forEach(chunk => {
        expect(chunk.text).toContain('.');
      });
    });
  });
});