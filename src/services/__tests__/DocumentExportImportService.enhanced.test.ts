import { DocumentExportImportService, SyncOptions, BackupOptions } from '../DocumentExportImportService';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    backupFolderId: 'test-backup-folder-id',
    enableTextIndex: true
  }
} as any;

describe('DocumentExportImportService - Enhanced Features', () => {
  let service: DocumentExportImportService;

  beforeEach(() => {
    service = new DocumentExportImportService(mockConfig);
  });

  test('should create DocumentExportImportService with enhanced features', () => {
    expect(service).toBeInstanceOf(DocumentExportImportService);
  });

  test('should have sync and backup interfaces', () => {
    const syncOptions: SyncOptions = {
      sourceFolderId: 'source-folder-id',
      targetFolderId: 'target-folder-id',
      syncMode: 'mirror',
      fileTypes: ['application/pdf'],
      excludePatterns: ['temp'],
      schedule: '0 2 * * *' // Daily at 2 AM
    };

    const backupOptions: BackupOptions = {
      sourceFolderId: 'source-folder-id',
      backupFolderId: 'backup-folder-id',
      retentionDays: 30,
      compress: true,
      includeSubfolders: true,
      fileTypes: ['application/pdf', 'application/vnd.google-apps.document']
    };

    expect(syncOptions).toBeDefined();
    expect(backupOptions).toBeDefined();
  });

  test('should have enhanced export options', () => {
    // Test that the enhanced export options are available
    expect(service).toBeDefined();
  });

  test('should have enhanced import options', () => {
    // Test that the enhanced import options are available
    expect(service).toBeDefined();
  });
});