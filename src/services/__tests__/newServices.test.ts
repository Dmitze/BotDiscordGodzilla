import { SmartDocumentClassifier } from '../SmartDocumentClassifier';
import { DocumentCardBuilder } from '../../ui/DocumentCardBuilder';
import { DriveChangesService } from '../DriveChangesService';
import { MultilingualDocumentProcessor } from '../MultilingualDocumentProcessor';
import { DocumentAnalyticsService } from '../DocumentAnalyticsService';
import { DocumentMentionHandler } from '../DocumentMentionHandler';
import { AutomatedDocumentProcessor } from '../AutomatedDocumentProcessor';
import { DocumentExportImportService } from '../DocumentExportImportService';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    enableTextIndex: true
  }
} as any;

describe('New Services', () => {
  test('should be able to import SmartDocumentClassifier', () => {
    const classifier = new SmartDocumentClassifier(mockConfig);
    expect(classifier).toBeInstanceOf(SmartDocumentClassifier);
  });

  test('should be able to import DocumentCardBuilder', () => {
    const mockFile = {
      id: 'test-id',
      name: 'Test Document',
      mimeType: 'application/pdf'
    } as any;
    
    const cardBuilder = new DocumentCardBuilder(mockFile);
    expect(cardBuilder).toBeInstanceOf(DocumentCardBuilder);
  });

  test('should be able to import DriveChangesService', () => {
    const changesService = new DriveChangesService(mockConfig);
    expect(changesService).toBeInstanceOf(DriveChangesService);
  });

  test('should be able to import MultilingualDocumentProcessor', () => {
    const multilingualService = new MultilingualDocumentProcessor(mockConfig);
    expect(multilingualService).toBeInstanceOf(MultilingualDocumentProcessor);
  });

  test('should be able to import DocumentAnalyticsService', () => {
    const analyticsService = new DocumentAnalyticsService(mockConfig);
    expect(analyticsService).toBeInstanceOf(DocumentAnalyticsService);
  });

  test('should be able to import DocumentMentionHandler', () => {
    const mentionHandler = new DocumentMentionHandler(mockConfig);
    expect(mentionHandler).toBeInstanceOf(DocumentMentionHandler);
  });

  test('should be able to import AutomatedDocumentProcessor', () => {
    const autoProcessor = new AutomatedDocumentProcessor(mockConfig);
    expect(autoProcessor).toBeInstanceOf(AutomatedDocumentProcessor);
  });

  test('should be able to import DocumentExportImportService', () => {
    const exportService = new DocumentExportImportService(mockConfig);
    expect(exportService).toBeInstanceOf(DocumentExportImportService);
  });
});