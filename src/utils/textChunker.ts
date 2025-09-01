import { countTokens } from './token';

export interface TextChunk {
  text: string;
  start: number;
  end: number;
  tokenCount: number;
}

/**
 * Chunk text into segments with specified token limits and overlap
 * @param text The text to chunk
 * @param targetTokens Target number of tokens per chunk (default: 1000)
 * @param minTokens Minimum number of tokens per chunk (default: 800)
 * @param maxTokens Maximum number of tokens per chunk (default: 1200)
 * @param overlapTokens Number of tokens to overlap between chunks (default: 100)
 * @returns Array of text chunks with metadata
 */
export function chunkTextByTokens(
  text: string,
  _targetTokens: number = 1000,
  minTokens: number = 800,
  maxTokens: number = 1200,
  overlapTokens: number = 100
): TextChunk[] {
  if (!text || text.trim().length === 0) {
    return [];
  }

  const chunks: TextChunk[] = [];
  let position = 0;
  let charPosition = 0;

  while (position < text.length) {
    // Start with a reasonable estimate for chunk size
    let chunkEnd = Math.min(position + (maxTokens * 4), text.length);
    let chunkText = text.substring(position, chunkEnd);
    let tokenCount = countTokens(chunkText);
    
    // Adjust chunk size based on token count
    while (tokenCount > maxTokens && chunkText.length > 0) {
      // Reduce chunk size
      chunkText = chunkText.substring(0, Math.floor(chunkText.length * 0.9));
      tokenCount = countTokens(chunkText);
    }
    
    // If chunk is still too small and we have more text, try to expand
    if (tokenCount < minTokens && position + chunkText.length < text.length) {
      let expandedEnd = Math.min(position + (maxTokens * 4), text.length);
      let expandedText = text.substring(position, expandedEnd);
      let expandedTokenCount = countTokens(expandedText);
      
      // Expand while within limits
      while (expandedTokenCount <= maxTokens && expandedEnd < text.length) {
        chunkText = expandedText;
        tokenCount = expandedTokenCount;
        
        expandedEnd = Math.min(expandedEnd + Math.floor(maxTokens / 2), text.length);
        expandedText = text.substring(position, expandedEnd);
        expandedTokenCount = countTokens(expandedText);
      }
    }
    
    // Add chunk to results
    const endChar = charPosition + chunkText.length;
    
    chunks.push({
      text: chunkText,
      start: charPosition,
      end: endChar,
      tokenCount: tokenCount
    });
    
    // Move position for next chunk
    if (position + chunkText.length >= text.length) {
      // Reached the end
      break;
    }
    
    // Calculate next position with overlap
    // Move forward by the chunk size minus overlap
    const nextPosition = position + Math.floor(chunkText.length * 0.8);
    position = Math.max(nextPosition, position + 1); // Ensure we move forward
    
    // Update character position for next chunk
    charPosition += chunkText.length - overlapTokens;
  }
  
  return chunks;
}

/**
 * Simple sentence-based chunking as a fallback
 * @param text The text to chunk
 * @param maxTokens Maximum tokens per chunk
 * @returns Array of text chunks
 */
export function chunkTextBySentences(text: string, maxTokens: number = 1000): TextChunk[] {
  if (!text || text.trim().length === 0) {
    return [];
  }

  // Split by sentences (simple approach)
  const sentences = text.split(/(?<=[.!?])\s+/);
  const chunks: TextChunk[] = [];
  let currentChunk: string[] = [];
  let currentTokenCount = 0;
  let charPosition = 0;

  for (const sentence of sentences) {
    const sentenceTokens = countTokens(sentence);
    
    // If adding this sentence would exceed the limit, save current chunk and start new one
    if (currentTokenCount + sentenceTokens > maxTokens && currentChunk.length > 0) {
      const chunkText = currentChunk.join(' ');
      chunks.push({
        text: chunkText,
        start: charPosition,
        end: charPosition + chunkText.length,
        tokenCount: currentTokenCount
      });
      
      // For overlap, keep some sentences from the previous chunk
      const overlapSentences = Math.min(2, Math.floor(currentChunk.length / 3));
      if (overlapSentences > 0) {
        currentChunk = currentChunk.slice(-overlapSentences);
        currentTokenCount = countTokens(currentChunk.join(' '));
        // Adjust charPosition for overlap
        const overlapText = currentChunk.join(' ');
        charPosition = charPosition + chunkText.length - overlapText.length;
      } else {
        currentChunk = [];
        currentTokenCount = 0;
        charPosition += chunkText.length + 1; // +1 for space
      }
    }
    
    currentChunk.push(sentence);
    currentTokenCount += sentenceTokens;
  }
  
  // Don't forget the last chunk
  if (currentChunk.length > 0) {
    const chunkText = currentChunk.join(' ');
    chunks.push({
      text: chunkText,
      start: charPosition,
      end: charPosition + chunkText.length,
      tokenCount: currentTokenCount
    });
  }
  
  return chunks;
}