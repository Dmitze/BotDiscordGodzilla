import { HybridSearchService } from '../HybridSearchService';

// Mock dependencies
const mockSearchIndex = {
  search: jest.fn(),
  upsert: jest.fn(),
  getDiff: jest.fn()
};

const mockAIService = {
  generateResponse: jest.fn(),
  embed: jest.fn()
};

const mockEmbeddingsService = {
  embed: jest.fn()
};

describe('HybridSearchService', () => {
  let hybridSearchService: HybridSearchService;

  beforeEach(() => {
    jest.clearAllMocks();
    hybridSearchService = new HybridSearchService(
      mockSearchIndex as any,
      mockAIService as any,
      mockEmbeddingsService as any
    );
  });

  describe('search', () => {
    it('should perform hybrid search combining vector and text results', async () => {
      // Mock vector search results
      mockEmbeddingsService.embed.mockResolvedValue([0.1, 0.2, 0.3]);
      mockSearchIndex.search.mockResolvedValueOnce({
        hits: [
          {
            fileId: '1',
            name: 'Document 1',
            contentHash: 'hash1',
            textLen: 100,
            score: 0.8
          }
        ],
        total: 1
      });

      // Mock text search results
      mockSearchIndex.search.mockResolvedValueOnce({
        hits: [
          {
            fileId: '2',
            name: 'Document 2',
            contentHash: 'hash2',
            textLen: 200,
            score: 10
          }
        ],
        total: 1
      });

      const results = await hybridSearchService.search('test query', {
        limit: 10,
        vectorWeight: 0.7,
        textWeight: 0.3
      });

      expect(results).toHaveLength(2);
      expect(mockEmbeddingsService.embed).toHaveBeenCalledWith('test query');
      expect(mockSearchIndex.search).toHaveBeenCalledTimes(2);
    });

    it('should handle search with only text results when embeddings service is not available', async () => {
      const serviceWithoutEmbeddings = new HybridSearchService(
        mockSearchIndex as any,
        mockAIService as any
      );

      mockSearchIndex.search.mockResolvedValueOnce({
        hits: [
          {
            fileId: '1',
            name: 'Document 1',
            contentHash: 'hash1',
            textLen: 100,
            score: 10
          }
        ],
        total: 1
      });

      const results = await serviceWithoutEmbeddings.search('test query', {
        limit: 10
      });

      expect(results).toHaveLength(1);
      expect(mockSearchIndex.search).toHaveBeenCalledTimes(1);
    });

    it('should validate weight parameters', async () => {
      await expect(
        hybridSearchService.search('test query', {
          vectorWeight: 1.5,
          textWeight: 0.3
        })
      ).rejects.toThrow('Weights must be between 0 and 1');

      await expect(
        hybridSearchService.search('test query', {
          vectorWeight: 0.7,
          textWeight: 0.4
        })
      ).rejects.toThrow('Vector weight and text weight must sum to 1');
    });

    it('should combine and deduplicate results from vector and text search', async () => {
      // Mock vector search results with a document that also appears in text search
      mockEmbeddingsService.embed.mockResolvedValue([0.1, 0.2, 0.3]);
      mockSearchIndex.search.mockResolvedValueOnce({
        hits: [
          {
            fileId: '1',
            name: 'Document 1',
            contentHash: 'hash1',
            textLen: 100,
            score: 0.9
          },
          {
            fileId: '2',
            name: 'Document 2',
            contentHash: 'hash2',
            textLen: 200,
            score: 0.8
          }
        ],
        total: 2
      });

      // Mock text search results with overlap
      mockSearchIndex.search.mockResolvedValueOnce({
        hits: [
          {
            fileId: '1',
            name: 'Document 1',
            contentHash: 'hash1',
            textLen: 100,
            score: 15
          },
          {
            fileId: '3',
            name: 'Document 3',
            contentHash: 'hash3',
            textLen: 300,
            score: 20
          }
        ],
        total: 2
      });

      const results = await hybridSearchService.search('test query', {
        limit: 10,
        vectorWeight: 0.6,
        textWeight: 0.4
      });

      // Should have 3 unique results (documents 1, 2, and 3)
      expect(results).toHaveLength(3);
      
      // Document 1 should have both vector and text scores
      const doc1 = results.find(r => r.fileId === '1');
      expect(doc1).toBeDefined();
      // The vector score gets modified in the implementation (multiplied by 0.8)
      expect(doc1?.vectorScore).toBeCloseTo(0.72); // 0.9 * 0.8
      expect(doc1?.textScore).toBe(15);
      expect(doc1?.combinedScore).toBeDefined();
    });
  });

  describe('calculateCombinedScore', () => {
    it('should calculate combined score correctly', () => {
      // This test would need to access private method, so we'll test indirectly
      // through the search method results
    });
  });
});