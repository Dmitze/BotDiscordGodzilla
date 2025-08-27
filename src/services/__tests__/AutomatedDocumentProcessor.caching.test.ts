import { AutomatedDocumentProcessor } from '../AutomatedDocumentProcessor';
import type { GoogleService } from '../../services/GoogleService';
import type { CacheService } from '../../services/CacheService';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    enableTextIndex: true
  }
} as any;

// Mock services
const mockGoogleService = {
  extractTextForChat: jest.fn()
} as unknown as GoogleService;

const mockCacheService = {
  get: jest.fn(),
  set: jest.fn()
} as unknown as CacheService;

describe('AutomatedDocumentProcessor - Caching', () => {
  let processor: AutomatedDocumentProcessor;

  beforeEach(() => {
    processor = new AutomatedDocumentProcessor(mockConfig);
    jest.clearAllMocks();
  });

  test('should use cache for document summarization', async () => {
    // Initialize with cache service
    (processor as any).cache = mockCacheService;
    (processor as any).google = mockGoogleService;
    
    // Mock cache to return a cached summary
    const cachedSummary = 'This is a cached summary';
    mockCacheService.get = jest.fn().mockResolvedValue(cachedSummary);
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Call summarizeDocument
    const summary = await (processor as any).summarizeDocument(mockFile);
    
    // Verify cache was checked
    expect(mockCacheService.get).toHaveBeenCalledWith('doc:summary:test-file-id');
    expect(summary).toBe(cachedSummary);
    
    // Verify Google service was not called
    expect(mockGoogleService.extractTextForChat).not.toHaveBeenCalled();
  });

  test('should cache document summary when not in cache', async () => {
    // Initialize with cache service
    (processor as any).cache = mockCacheService;
    (processor as any).google = mockGoogleService;
    
    // Mock cache to return null (not found)
    mockCacheService.get = jest.fn().mockResolvedValue(null);
    
    // Mock Google service to return content
    mockGoogleService.extractTextForChat = jest.fn().mockResolvedValue({
      text: 'This is the document content that will be summarized.'
    });
    
    // Mock cache set
    mockCacheService.set = jest.fn().mockResolvedValue(true);
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Call summarizeDocument
    const summary = await (processor as any).summarizeDocument(mockFile);
    
    // Verify cache was checked
    expect(mockCacheService.get).toHaveBeenCalledWith('doc:summary:test-file-id');
    
    // Verify Google service was called
    expect(mockGoogleService.extractTextForChat).toHaveBeenCalledWith('test-file-id');
    
    // Verify result was cached
    expect(mockCacheService.set).toHaveBeenCalledWith(
      'doc:summary:test-file-id',
      expect.any(String),
      3600
    );
    
    // Verify we got a summary
    expect(summary).toContain('document');
  });

  test('should use cache for document tags', async () => {
    // Initialize with cache service
    (processor as any).cache = mockCacheService;
    
    // Mock cache to return cached tags
    const cachedTags = ['tag1', 'tag2', 'tag3'];
    mockCacheService.get = jest.fn().mockResolvedValue(cachedTags);
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Mock trigger
    const mockTrigger = {
      autoTaggingConfig: {
        enabled: true,
        useAI: false,
        keywordThreshold: 0.5,
        maxTags: 5
      }
    } as any;
    
    // Call autoTagDocument
    const tags = await (processor as any).autoTagDocument(mockFile, mockTrigger);
    
    // Verify cache was checked
    expect(mockCacheService.get).toHaveBeenCalledWith('doc:tags:test-file-id');
    expect(tags).toEqual(cachedTags);
  });

  test('should cache document tags when not in cache', async () => {
    // Initialize with cache service
    (processor as any).cache = mockCacheService;
    
    // Mock cache to return null (not found)
    mockCacheService.get = jest.fn().mockResolvedValue(null);
    
    // Mock cache set
    mockCacheService.set = jest.fn().mockResolvedValue(true);
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Mock trigger
    const mockTrigger = {
      autoTaggingConfig: {
        enabled: true,
        useAI: false,
        keywordThreshold: 0.5,
        maxTags: 5
      }
    } as any;
    
    // Call autoTagDocument
    const tags = await (processor as any).autoTagDocument(mockFile, mockTrigger);
    
    // Verify cache was checked
    expect(mockCacheService.get).toHaveBeenCalledWith('doc:tags:test-file-id');
    
    // Verify result was cached
    expect(mockCacheService.set).toHaveBeenCalledWith(
      'doc:tags:test-file-id',
      expect.any(Array),
      3600
    );
    
    // Verify we got tags
    expect(tags).toBeInstanceOf(Array);
  });

  test('should use cache for version comparison', async () => {
    // Initialize with cache service
    (processor as any).cache = mockCacheService;
    
    // Mock cache to return cached comparison
    const cachedComparison = {
      fileId: 'test-file-id',
      fileName: 'Test Document.pdf',
      versions: [],
      differences: {
        added: [],
        removed: [],
        modified: []
      },
      summary: 'Cached comparison'
    };
    mockCacheService.get = jest.fn().mockResolvedValue(cachedComparison);
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Call compareDocumentVersions
    const comparison = await (processor as any).compareDocumentVersions(mockFile);
    
    // Verify cache was checked
    expect(mockCacheService.get).toHaveBeenCalledWith('doc:version-comparison:test-file-id');
    expect(comparison).toEqual(cachedComparison);
  });

  test('should cache version comparison when not in cache', async () => {
    // Initialize with cache service
    (processor as any).cache = mockCacheService;
    
    // Mock cache to return null (not found)
    mockCacheService.get = jest.fn().mockResolvedValue(null);
    
    // Mock cache set
    mockCacheService.set = jest.fn().mockResolvedValue(true);
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Call compareDocumentVersions
    const comparison = await (processor as any).compareDocumentVersions(mockFile);
    
    // Verify cache was checked
    expect(mockCacheService.get).toHaveBeenCalledWith('doc:version-comparison:test-file-id');
    
    // Verify result was cached
    expect(mockCacheService.set).toHaveBeenCalledWith(
      'doc:version-comparison:test-file-id',
      expect.any(Object),
      1800
    );
    
    // Verify we got a comparison
    expect(comparison).toHaveProperty('fileId');
    expect(comparison).toHaveProperty('fileName');
  });

  test('should work without cache service', async () => {
    // Don't initialize cache service
    (processor as any).cache = null;
    (processor as any).google = mockGoogleService;
    
    // Mock Google service to return content
    mockGoogleService.extractTextForChat = jest.fn().mockResolvedValue({
      text: 'This is the document content that will be summarized.'
    });
    
    // Mock file
    const mockFile = {
      id: 'test-file-id',
      name: 'Test Document.pdf',
      mimeType: 'application/pdf'
    } as any;
    
    // Call summarizeDocument
    const summary = await (processor as any).summarizeDocument(mockFile);
    
    // Verify Google service was called
    expect(mockGoogleService.extractTextForChat).toHaveBeenCalledWith('test-file-id');
    
    // Verify we got a summary
    expect(summary).toContain('document');
  });
});