import { GoogleDocsService } from '../GoogleDocsService';
import { DocsService } from '../google/DocsService';
import { CacheService } from '../CacheService';
import { google } from 'googleapis';

// Моки для залежностей
jest.mock('../google/DocsService');
jest.mock('../CacheService');
jest.mock('googleapis');

describe('GoogleDocsService', () => {
  let googleDocsService: GoogleDocsService;
  let mockConfig: any;
  let mockAuth: any;
  let mockMetrics: any;

  beforeEach(() => {
    mockConfig = {
      drive: {
        ttlListSec: 300,
        ttlTextSec: 300,
      },
      google: {
        ocrCacheTTL: 3600,
      },
      performance: {
        cacheTTL: 3600,
      },
    };

    mockAuth = {
      // Порожній об'єкт для мокування JWT авторизації
    };

    mockMetrics = {
      updateGoogleApiMetrics: jest.fn(),
    };

    // Створення екземпляра сервісу
    googleDocsService = new GoogleDocsService(mockConfig, mockAuth, mockMetrics);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('listDocs', () => {
    it('should return an empty array when no documents are found', async () => {
      // Мок для Drive API
      const mockDriveFilesList = jest.fn().mockResolvedValue({
        data: {
          files: [],
        },
      });

      (google.drive as jest.Mock).mockReturnValue({
        files: {
          list: mockDriveFilesList,
        },
      });

      // Мок для Docs API
      (google.docs as jest.Mock).mockReturnValue({
        documents: {
          get: jest.fn(),
        },
      });

      const result = await googleDocsService.listDocs();
      
      expect(result).toEqual([]);
      expect(mockDriveFilesList).toHaveBeenCalledWith({
        q: "mimeType='application/vnd.google-apps.document' and trashed = false",
        pageSize: 100,
        fields: 'files(id,name,mimeType,modifiedTime,owners(displayName,emailAddress))',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });
    });

    it('should return a list of documents when documents are found', async () => {
      const mockFiles = [
        {
          id: 'doc1',
          name: 'Test Document 1',
          mimeType: 'application/vnd.google-apps.document',
          modifiedTime: '2023-01-01T00:00:00Z',
          owners: [{ displayName: 'Test User', emailAddress: 'test@example.com' }],
        },
        {
          id: 'doc2',
          name: 'Test Document 2',
          mimeType: 'application/vnd.google-apps.document',
          modifiedTime: '2023-01-02T00:00:00Z',
          owners: [{ displayName: 'Test User 2', emailAddress: 'test2@example.com' }],
        },
      ];

      // Мок для Drive API
      const mockDriveFilesList = jest.fn().mockResolvedValue({
        data: {
          files: mockFiles,
        },
      });

      (google.drive as jest.Mock).mockReturnValue({
        files: {
          list: mockDriveFilesList,
        },
      });

      // Мок для Docs API
      (google.docs as jest.Mock).mockReturnValue({
        documents: {
          get: jest.fn(),
        },
      });

      const result = await googleDocsService.listDocs();
      
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({
        id: 'doc1',
        name: 'Test Document 1',
        mimeType: 'application/vnd.google-apps.document',
        modifiedTime: '2023-01-01T00:00:00Z',
        owners: [{ displayName: 'Test User', emailAddress: 'test@example.com' }],
      });
      expect(result[1]).toEqual({
        id: 'doc2',
        name: 'Test Document 2',
        mimeType: 'application/vnd.google-apps.document',
        modifiedTime: '2023-01-02T00:00:00Z',
        owners: [{ displayName: 'Test User 2', emailAddress: 'test2@example.com' }],
      });
    });
  });

  describe('getDocContent', () => {
    it('should return document content when document is found', async () => {
      const documentId = 'test-doc-id';
      const mockDocument = {
        title: 'Test Document',
        body: {
          content: [
            {
              paragraph: {
                elements: [
                  {
                    textRun: {
                      content: 'Test content',
                    },
                  },
                ],
              },
            },
          ],
        },
      };

      // Мок для Docs API
      const mockDocsGet = jest.fn().mockResolvedValue({
        data: mockDocument,
      });

      (google.docs as jest.Mock).mockReturnValue({
        documents: {
          get: mockDocsGet,
        },
      });

      // Мок для DocsService
      (DocsService as jest.Mock).mockImplementation(() => {
        return {
          extractTextFromDoc: jest.fn().mockReturnValue('Test content'),
          extractBlocksFromDoc: jest.fn().mockReturnValue([{ kind: 'paragraph', text: 'Test content' }]),
        };
      });

      const result = await googleDocsService.getDocContent(documentId);
      
      expect(result).toEqual({
        title: 'Test Document',
        content: 'Test content',
        blocks: [{ kind: 'paragraph', text: 'Test content' }],
        modifiedTime: undefined,
      });
      expect(mockDocsGet).toHaveBeenCalledWith({
        documentId,
        fields: 'title,body,documentStyle,headers,footers,footnotes,lists,tables,revisions',
      });
    });
  });

  describe('indexDoc', () => {
    it('should return success result when document is indexed', async () => {
      const documentId = 'test-doc-id';

      // Мок для getDocContent
      const mockGetDocContent = jest.spyOn(googleDocsService as any, 'getDocContent').mockResolvedValue({
        title: 'Test Document',
        content: 'This is a test document with some content',
        blocks: [],
      });

      const result = await googleDocsService.indexDoc(documentId);
      
      expect(result).toEqual({
        success: true,
        documentId,
        indexedAt: expect.any(String),
        contentHash: expect.any(String),
        wordCount: 7, // "This is a test document with some content" = 7 words
      });
      
      // Перевірка, що getDocContent був викликаний
      expect(mockGetDocContent).toHaveBeenCalledWith(documentId);
    });
  });

  describe('searchDoc', () => {
    it('should return search results when query matches content', async () => {
      const documentId = 'test-doc-id';
      const query = 'test';

      // Мок для getDocContent
      const mockGetDocContent = jest.spyOn(googleDocsService as any, 'getDocContent').mockResolvedValue({
        title: 'Test Document',
        content: 'This is a test document with some content',
        blocks: [
          { kind: 'paragraph', text: 'This is a test document with some content' },
          { kind: 'heading', level: 1, text: 'Test Heading' },
        ],
      });

      const result = await googleDocsService.searchDoc(documentId, query);
      
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({
        blockIndex: 0,
        blockType: 'paragraph',
        content: 'This is a test document with some content',
        matchPosition: 10, // Position of "test" in the string
        relevanceScore: expect.any(Number),
      });
      expect(result[1]).toEqual({
        blockIndex: 1,
        blockType: 'heading-1',
        content: 'Test Heading',
        matchPosition: 0, // Position of "Test" in the string
        relevanceScore: expect.any(Number),
      });
      
      // Перевірка, що getDocContent був викликаний
      expect(mockGetDocContent).toHaveBeenCalledWith(documentId);
    });
  });

  describe('summarizeDoc', () => {
    it('should return document summary', async () => {
      const documentId = 'test-doc-id';

      // Мок для getDocContent
      const mockGetDocContent = jest.spyOn(googleDocsService as any, 'getDocContent').mockResolvedValue({
        title: 'Test Document',
        content: 'This is a test document with some content. It has multiple sentences for testing purposes.',
        blocks: [
          { kind: 'paragraph', text: 'This is a test document with some content. It has multiple sentences for testing purposes.' },
          { kind: 'heading', level: 1, text: 'Introduction' },
          { kind: 'heading', level: 2, text: 'Main Content' },
        ],
      });

      const result = await googleDocsService.summarizeDoc(documentId);
      
      expect(result).toEqual({
        title: 'Test Document',
        summary: 'This is a test document with some content. It has multiple sentences for testing purposes.',
        keyPoints: ['Introduction', 'Main Content'],
        wordCount: 15,
        readingTimeMinutes: 1, // ceil(15/200) = 1
      });
      
      // Перевірка, що getDocContent був викликаний
      expect(mockGetDocContent).toHaveBeenCalledWith(documentId);
    });
  });
});