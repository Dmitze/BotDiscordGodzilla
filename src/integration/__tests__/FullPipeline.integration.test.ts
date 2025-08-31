import { GoogleDocsService } from '../../services/GoogleDocsService';
import { SqliteSearchIndex } from '../../search/sqlite/SqliteSearchIndex';
import { HybridRetriever } from '../../rag/HybridRetriever';
import { PromptTemplatesService } from '../../services/PromptTemplatesService';
import { mkdirSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

// Mock Google Auth
const mockGoogleAuth = {};

// Mock config
const mockConfig = {};

describe('Full Pipeline Integration', () => {
  let googleDocsService: GoogleDocsService;
  let searchIndex: SqliteSearchIndex;
  let retriever: HybridRetriever;
  let promptTemplates: PromptTemplatesService;
  let tempDbPath: string;

  beforeAll(() => {
    // Create a temporary database file for testing
    const tempDir = join(tmpdir(), 'bot-discord-godzilla-full-test');
    mkdirSync(tempDir, { recursive: true });
    tempDbPath = join(tempDir, 'full-test-search-index.db');
    
    // Create search index
    searchIndex = new SqliteSearchIndex({ dbPath: tempDbPath });
    
    // Create retriever
    retriever = new HybridRetriever(searchIndex);
    
    // Create prompt templates service
    promptTemplates = new PromptTemplatesService(mockConfig as any);
  });

  beforeEach(() => {
    googleDocsService = new GoogleDocsService(mockConfig as any, mockGoogleAuth as any);
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

  describe('Document Processing Pipeline', () => {
    it('should process a document through the full pipeline', async () => {
      // Mock document content
      const mockDocContent = {
        title: 'Test Document',
        content: 'This is a comprehensive test document with multiple sentences. It contains various information that should be properly chunked and indexed. The document is designed to test the full processing pipeline of the bot.',
        blocks: [],
        modifiedTime: new Date().toISOString()
      };

      // Mock the getDocContent method
      (googleDocsService as any).getDocContent = jest.fn().mockResolvedValue(mockDocContent);

      // Step 1: Index the document
      const documentId = 'test-document-id';
      const indexResult = await googleDocsService.indexDoc(documentId);

      expect(indexResult.success).toBe(true);
      expect(indexResult.documentId).toBe(documentId);
      expect(indexResult.wordCount).toBeGreaterThan(0);

      // Step 2: Search for content in the document
      const searchResults = await retriever.retrieve('test document', { k: 5 });

      // We should get some results
      expect(searchResults.length).toBeGreaterThan(0);
      
      // Results should have the expected properties
      expect(searchResults[0]).toHaveProperty('fileId');
      expect(searchResults[0]).toHaveProperty('name');
      expect(searchResults[0]).toHaveProperty('snippet');
      expect(searchResults[0]).toHaveProperty('fusedScore');

      // Step 3: Use prompt templates
      const qaPrompt = promptTemplates.renderPrompt('document_qa', {
        question: 'What is this document about?',
        context: searchResults.map((r, i) => `(${i + 1}) ${r.name}\n${r.snippet}`).join('\n\n')
      });

      expect(qaPrompt).toBeDefined();
      expect(qaPrompt).toContain('What is this document about?');
      expect(qaPrompt).toContain('Test Document');

      // Step 4: Test summary template
      const summaryPrompt = promptTemplates.renderPrompt('document_summary', {
        document_name: 'Test Document',
        document_text: mockDocContent.content
      });

      expect(summaryPrompt).toBeDefined();
      expect(summaryPrompt).toContain('Test Document');
      expect(summaryPrompt).toContain('Створи короткий зміст');

    }, 10000); // 10 second timeout
  });
});