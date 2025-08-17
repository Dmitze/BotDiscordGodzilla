/*
 * Utilities to split long texts into Discord-safe chunks
 */
export function chunkTextForDiscord(text: string, max = 1900): string[] {
  const s = text || '';
  if (!s) return [];
  const chunks: string[] = [];
  let i = 0;
  while (i < s.length) {
    const end = Math.min(i + max, s.length);
    let slice = s.slice(i, end);
    // avoid cutting in the middle of a word if possible
    if (end < s.length) {
      const lastSpace = slice.lastIndexOf(' ');
      if (lastSpace > max * 0.6) {
        slice = slice.slice(0, lastSpace);
        i += lastSpace + 1;
      } else {
        i = end;
      }
    } else {
      i = end;
    }
    chunks.push(slice);
  }
  return chunks;
}
