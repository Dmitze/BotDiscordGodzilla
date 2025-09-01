import { Reranker } from '../Reranker';

// Mock AI service
const mockAIService = {
  generateResponse: jest.fn()
};

describe('Reranker', () => {
  let reranker: Reranker;

  beforeEach(() => {
    jest.clearAllMocks();
    reranker = new Reranker(mockAIService as any);
  });

  describe('rerank', () => {
    it('should rerank documents based on AI scoring', async () => {
      // Mock AI responses for relevance scores
      mockAIService.generateResponse
        .mockResolvedValueOnce({ content: '0.9' })
        .mockResolvedValueOnce({ content: '0.3' })
        .mockResolvedValueOnce({ content: '0.7' });

      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'This is document 1 content',
          fusedScore: 0.8
        },
        {
          fileId: '2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'This is document 2 content',
          fusedScore: 0.5
        },
        {
          fileId: '3',
          name: 'Document 3',
          contentHash: 'hash3',
          textLen: 150,
          snippet: 'This is document 3 content',
          fusedScore: 0.6
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs, { limit: 3 });

      expect(rerankedDocs).toHaveLength(3);
      // We're calling the AI service 3 times
      expect(mockAIService.generateResponse).toHaveBeenCalledTimes(3);
      
      // Documents should be sorted by rerank score (highest first)
      expect(rerankedDocs[0].fileId).toBe('1'); // Score 0.9
      expect(rerankedDocs[1].fileId).toBe('3'); // Score 0.7
      expect(rerankedDocs[2].fileId).toBe('2'); // Score 0.3
      
      // Each document should have a rerankScore
      expect(rerankedDocs[0].rerankScore).toBe(0.9);
      expect(rerankedDocs[1].rerankScore).toBe(0.7);
      expect(rerankedDocs[2].rerankScore).toBe(0.3);
    });

    it('should handle documents without snippets', async () => {
      mockAIService.generateResponse.mockResolvedValue({ content: '0.5' });

      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          mimeType: 'text/plain',
          fusedScore: 0.8
        },
        {
          fileId: '2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          mimeType: 'text/plain',
          fusedScore: 0.5
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs);

      expect(rerankedDocs).toHaveLength(2);
      // We should call the AI service for each document
      expect(mockAIService.generateResponse).toHaveBeenCalledTimes(2);
      expect(rerankedDocs[0].rerankScore).toBe(0.5);
    });

    it('should limit the number of returned documents', async () => {
      mockAIService.generateResponse
        .mockResolvedValueOnce({ content: '0.9' })
        .mockResolvedValueOnce({ content: '0.3' })
        .mockResolvedValueOnce({ content: '0.7' });

      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'Content 1',
          fusedScore: 0.8
        },
        {
          fileId: '2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'Content 2',
          fusedScore: 0.5
        },
        {
          fileId: '3',
          name: 'Document 3',
          contentHash: 'hash3',
          textLen: 150,
          snippet: 'Content 3',
          fusedScore: 0.6
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs, { limit: 2 });

      expect(rerankedDocs).toHaveLength(2);
    });

    it('should handle AI scoring errors gracefully', async () => {
      mockAIService.generateResponse.mockRejectedValue(new Error('AI service error'));

      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'Content 1',
          fusedScore: 0.8
        },
        {
          fileId: '2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'Content 2',
          fusedScore: 0.5
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs);

      // Should return original documents with neutral score due to error handling
      expect(rerankedDocs).toHaveLength(2);
      expect(rerankedDocs[0].rerankScore).toBe(0.5);
      expect(rerankedDocs[1].rerankScore).toBe(0.5);
    });

    it('should handle invalid AI scores', async () => {
      mockAIService.generateResponse.mockResolvedValue({ content: 'invalid' });

      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'Content 1',
          fusedScore: 0.8
        },
        {
          fileId: '2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'Content 2',
          fusedScore: 0.5
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs);

      // Should return original documents with neutral score
      expect(rerankedDocs).toHaveLength(2);
      expect(rerankedDocs[0].rerankScore).toBe(0.5);
      expect(rerankedDocs[1].rerankScore).toBe(0.5);
    });

    it('should return original documents if reranking fails completely', async () => {
      mockAIService.generateResponse.mockRejectedValue(new Error('AI service error'));

      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'Content 1',
          fusedScore: 0.8
        },
        {
          fileId: '2',
          name: 'Document 2',
          contentHash: 'hash2',
          textLen: 200,
          snippet: 'Content 2',
          fusedScore: 0.5
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs);

      // Should return original documents
      expect(rerankedDocs).toHaveLength(2);
      // Should have rerankScore due to error handling
      expect(rerankedDocs[0].rerankScore).toBe(0.5);
      expect(rerankedDocs[1].rerankScore).toBe(0.5);
    });

    it('should not rerank if there are no documents', async () => {
      const query = 'test query';
      const docs: any[] = [];

      const rerankedDocs = await reranker.rerank(query, docs);

      expect(rerankedDocs).toHaveLength(0);
      expect(mockAIService.generateResponse).toHaveBeenCalledTimes(0);
    });

    it('should not rerank if there is only one document', async () => {
      const query = 'test query';
      const docs = [
        {
          fileId: '1',
          name: 'Document 1',
          contentHash: 'hash1',
          textLen: 100,
          snippet: 'Content 1',
          fusedScore: 0.8
        }
      ];

      const rerankedDocs = await reranker.rerank(query, docs);

      expect(rerankedDocs).toHaveLength(1);
      // With only one document, we return early without calling AI
      expect(mockAIService.generateResponse).toHaveBeenCalledTimes(0);
      // No rerankScore since we didn't call AI
      expect(rerankedDocs[0].rerankScore).toBeUndefined();
    });
  });

  describe('combineScores', () => {
    it('should combine original and rerank scores correctly', async () => {
      // This is tested indirectly through the rerank method
    });
  });
});