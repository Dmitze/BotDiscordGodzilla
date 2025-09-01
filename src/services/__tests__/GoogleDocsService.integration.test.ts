import { GoogleDocsService } from '../GoogleDocsService';
import type { BotConfig } from '@/types';
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';
import { mkdirSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

// Mock Google Auth
const mockGoogleAuth = {
  // Mock methods as needed
};

// Mock config
const mockConfig = {
  // Add any config properties that might be needed
} as BotConfig;

describe('GoogleDocsService Integration', () => {
  let googleDocsService: GoogleDocsService;
  let searchIndex: SqliteSearchIndex;
  let tempDbPath: string;

  beforeAll(() => {
    // Create a temporary database file for testing
    const tempDir = join(tmpdir(), 'bot-discord-godzilla-test');
    mkdirSync(tempDir, { recursive: true });
    tempDbPath = join(tempDir, 'test-search-index.db');
    
    // Create an in-memory search index for testing
    searchIndex = new SqliteSearchIndex({ dbPath: tempDbPath });
  });

  beforeEach(() => {
    googleDocsService = new GoogleDocsService(mockConfig, mockGoogleAuth as any);
    googleDocsService.setSearchIndex(searchIndex);
  });

  afterAll(() => {
    // Clean up temporary files
    try {
      // Database file will be automatically cleaned up by the OS
    } catch (e) {
      // Ignore cleanup errors
    }
  });

  describe('indexDoc', () => {
    it('should index a document and its chunks', async () => {
      // This is more of a unit test since we can't actually access Google Docs API in tests
      // But we can verify the integration with the search index works
      
      // Mock the getDocContent method to return test data
      (googleDocsService as any).getDocContent = jest.fn().mockResolvedValue({
        title: 'Test Document',
        content: 'This is a test document with some content for chunking and indexing.',
        blocks: [],
        modifiedTime: new Date().toISOString()
      });

      const documentId = 'test-document-id';
      const result = await googleDocsService.indexDoc(documentId);

      expect(result.success).toBe(true);
      expect(result.documentId).toBe(documentId);
      expect(result.wordCount).toBeGreaterThan(0);
      
      // Verify that the search index was called (we can't easily check the actual content
      // without making this a more complex integration test)
    });
  });
});