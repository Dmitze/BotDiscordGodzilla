import { HybridRetriever } from '../HybridRetriever';
import type { SearchIndex, SearchHit } from '@/search/SearchIndex';

// Mock search index
const mockSearchIndex: jest.Mocked<SearchIndex> = {
  upsert: jest.fn(),
  search: jest.fn(),
  getDiff: jest.fn()
};

// Mock embeddings service
const mockEmbeddings = {
  embed: jest.fn()
};

describe('HybridRetriever', () => {
  let retriever: HybridRetriever;

  beforeEach(() => {
    retriever = new HybridRetriever(mockSearchIndex, mockEmbeddings);
    jest.clearAllMocks();
  });

  describe('retrieve', () => {
    it('should retrieve documents using FTS when embeddings are not available', async () => {
      const mockHits: SearchHit[] = [
        {
          fileId: 'doc1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'This is document 1'
        },
        {
          fileId: 'doc2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'This is document 2'
        }
      ];

      mockSearchIndex.search.mockResolvedValue({ hits: mockHits, total: 2 });

      const result = await retriever.retrieve('test query', { mode: 'fts' });

      expect(mockSearchIndex.search).toHaveBeenCalledWith({
        text: 'test query',
        limit: 20 // Should get more candidates for reranking even in FTS mode
      });
      expect(result).toHaveLength(2);
      expect(result[0]).toHaveProperty('fusedScore');
    });

    it('should retrieve documents using hybrid search when embeddings are available', async () => {
      const mockHits: SearchHit[] = [
        {
          fileId: 'doc1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'This is document 1'
        }
      ];

      mockSearchIndex.search.mockResolvedValue({ hits: mockHits, total: 1 });
      mockEmbeddings.embed.mockResolvedValue([0.1, 0.2, 0.3]);

      const result = await retriever.retrieve('test query', { mode: 'hybrid' });

      expect(mockSearchIndex.search).toHaveBeenCalledWith({
        text: 'test query',
        limit: 20 // Should get more candidates for reranking
      });
      expect(mockEmbeddings.embed).toHaveBeenCalledWith('test query');
      expect(mockEmbeddings.embed).toHaveBeenCalledWith('This is document 1');
      expect(result).toHaveLength(1);
    });

    it('should limit results to specified k value', async () => {
      const mockHits: SearchHit[] = Array(30).fill(null).map((_, i) => ({
        fileId: `doc${i}`,
        name: `Document ${i}`,
        contentHash: `hash${i}`,
        textLen: 100,
        snippet: `This is document ${i}`
      }));

      mockSearchIndex.search.mockResolvedValue({ hits: mockHits, total: 30 });
      mockEmbeddings.embed.mockResolvedValue([0.1, 0.2, 0.3]);

      const result = await retriever.retrieve('test query', { mode: 'hybrid', k: 5 });

      expect(result).toHaveLength(5);
    });
  });

  describe('rerank', () => {
    it('should rerank documents based on multiple factors', async () => {
      const mockDocs = [
        {
          fileId: 'doc1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'This is document 1',
          fusedScore: 0.8
        },
        {
          fileId: 'doc2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'This is document 2',
          fusedScore: 0.6
        }
      ];

      mockEmbeddings.embed
        .mockResolvedValueOnce([0.1, 0.2, 0.3]) // Query embedding
        .mockResolvedValueOnce([0.2, 0.3, 0.4]) // Doc1 embedding
        .mockResolvedValueOnce([0.1, 0.1, 0.2]); // Doc2 embedding

      // Access private method through reflection for testing
      const rerankMethod = (retriever as any).rerank.bind(retriever);
      const result = await rerankMethod('test query', mockDocs);

      expect(result).toHaveLength(2);
      expect(result[0]).toHaveProperty('rerankMetadata');
      expect(result[0].rerankMetadata).toHaveProperty('tokenOverlap');
      expect(result[0].rerankMetadata).toHaveProperty('lengthScore');
    });
  });
});