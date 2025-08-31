import { countTokens, countTokensInArray } from '../token';

// Mock the tiktoken library
jest.mock('tiktoken', () => ({
  encoding_for_model: () => ({
    encode: (text: string) => ({
      length: Math.ceil(text.length / 4)
    })
  })
}));

describe('token', () => {
  describe('countTokens', () => {
    it('should count tokens in a string', () => {
      const text = 'This is a test string with several words';
      const result = countTokens(text);
      // With our mock, each 4 characters is 1 token
      expect(result).toBe(Math.ceil(text.length / 4));
    });

    it('should return 0 for empty string', () => {
      const result = countTokens('');
      expect(result).toBe(0);
    });

    it('should handle null or undefined gracefully', () => {
      const result1 = countTokens(null as any);
      const result2 = countTokens(undefined as any);
      expect(result1).toBe(0);
      expect(result2).toBe(0);
    });
  });

  describe('countTokensInArray', () => {
    it('should count tokens in an array of strings', () => {
      const texts = [
        'First string',
        'Second string with more words',
        'Third'
      ];
      
      const result = countTokensInArray(texts);
      const expected = texts.reduce((sum, text) => sum + Math.ceil(text.length / 4), 0);
      expect(result).toBe(expected);
    });

    it('should return 0 for empty array', () => {
      const result = countTokensInArray([]);
      expect(result).toBe(0);
    });
  });
});