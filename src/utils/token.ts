import { encoding_for_model } from 'tiktoken';

// Cache the encoding instance
let encoding: ReturnType<typeof encoding_for_model> | null = null;

/**
 * Get the tiktoken encoding instance
 * @returns The encoding instance
 */
function getEncoding() {
  if (!encoding) {
    encoding = encoding_for_model('gpt-4o-mini'); // Using a common model for token counting
  }
  return encoding;
}

/**
 * Count the number of tokens in a string
 * @param text The text to count tokens for
 * @returns The number of tokens
 */
export function countTokens(text: string): number {
  try {
    const enc = getEncoding();
    const tokens = enc.encode(text);
    return tokens.length;
  } catch (error) {
    // Fallback to rough estimation if tiktoken fails
    // Roughly 4 characters per token
    return Math.ceil((text || '').length / 4);
  }
}

/**
 * Count the number of tokens in an array of strings
 * @param texts Array of texts to count tokens for
 * @returns The total number of tokens
 */
export function countTokensInArray(texts: string[]): number {
  return texts.reduce((total, text) => total + countTokens(text), 0);
}