// Normalization helpers for parsed text.
// Delegate to existing normalizeText from utils/fileProcessor to keep behavior consistent.
import { normalizeText } from '@/utils/fileProcessor';

/**
 * Normalize Unicode, trim excessive whitespace, remove control chars, etc.
 * Uses the same pipeline as fileProcessor.normalizeText for consistency across the app.
 */
export function normalizeUnicode(input: string): string {
  return normalizeText(input ?? '');
}
