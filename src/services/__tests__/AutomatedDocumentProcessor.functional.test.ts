import { AutomatedDocumentProcessor, DocumentTrigger, AutoTaggingConfig, NotificationTemplate } from '../AutomatedDocumentProcessor';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    enableTextIndex: true
  }
} as any;

// Mock DriveFile
const mockDriveFile = {
  id: 'test-file-id',
  name: 'Test Document.pdf',
  mimeType: 'application/pdf',
  size: 1024,
  modifiedTime: '2023-01-01T00:00:00Z'
} as any;

describe('AutomatedDocumentProcessor - Functional Tests', () => {
  let processor: AutomatedDocumentProcessor;

  beforeEach(() => {
    processor = new AutomatedDocumentProcessor(mockConfig);
  });

  test('should handle trigger with no conditions', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: ['user1', 'user2']
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.conditions).toEqual([]);
    expect(trigger.actions).toEqual([]);
  });

  test('should handle trigger with complex conditions', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [
        { type: 'fileType', operator: 'equals', value: 'application/pdf' },
        { type: 'fileNamePattern', operator: 'contains', value: 'report' },
        { type: 'fileSize', operator: 'greaterThan', value: 1024 }
      ],
      actions: [],
      usersToNotify: ['user1', 'user2']
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.conditions).toHaveLength(3);
    expect(trigger.conditions[0].type).toBe('fileType');
    expect(trigger.conditions[1].type).toBe('fileNamePattern');
    expect(trigger.conditions[2].type).toBe('fileSize');
  });

  test('should handle trigger update correctly', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: ['user1', 'user2']
    };

    const trigger = processor.addTrigger(triggerData);
    const triggerId = trigger.id;
    
    const updateResult = processor.updateTrigger(triggerId, {
      enabled: false,
      folderName: 'Updated Folder'
    });
    
    expect(updateResult).toBe(true);
    
    const updatedTriggers = processor.getTriggers();
    const updatedTrigger = updatedTriggers.find(t => t.id === triggerId);
    
    expect(updatedTrigger).toBeDefined();
    expect(updatedTrigger?.enabled).toBe(false);
    expect(updatedTrigger?.folderName).toBe('Updated Folder');
  });

  test('should handle trigger removal correctly', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: ['user1', 'user2']
    };

    const trigger = processor.addTrigger(triggerData);
    const triggerId = trigger.id;
    
    // Verify trigger was added
    let triggers = processor.getTriggers();
    expect(triggers).toHaveLength(1);
    
    // Remove trigger
    const removalResult = processor.removeTrigger(triggerId);
    expect(removalResult).toBe(true);
    
    // Verify trigger was removed
    triggers = processor.getTriggers();
    expect(triggers).toHaveLength(0);
    
    // Try to remove non-existent trigger
    const falseRemoval = processor.removeTrigger('non-existent-id');
    expect(falseRemoval).toBe(false);
  });

  test('should handle edge case with special characters in file name', () => {
    const autoTaggingConfig: AutoTaggingConfig = {
      enabled: true,
      useAI: true,
      keywordThreshold: 0.5,
      maxTags: 5,
      customTags: ['important', 'review']
    };

    const notificationTemplate: NotificationTemplate = {
      title: 'New Document Detected',
      message: 'A new document matching your criteria has been found.',
      includeFileInfo: true,
      includeTags: true,
      includePreview: true,
      previewLength: 100
    };

    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: ['user1', 'user2'],
      autoTaggingConfig,
      notificationTemplate
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.autoTaggingConfig).toEqual(autoTaggingConfig);
    expect(trigger.notificationTemplate).toEqual(notificationTemplate);
  });

  test('should handle trigger with all action types', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'analyze' },
        { type: 'classify' },
        { type: 'tag', parameters: { tags: ['test-tag'] } },
        { type: 'notify' },
        { type: 'export', parameters: { format: 'pdf' } },
        { type: 'move', parameters: { targetFolderId: 'target-folder' } },
        { type: 'delete' }
      ],
      usersToNotify: ['user1', 'user2']
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(7);
    
    const actionTypes = trigger.actions.map(action => action.type);
    expect(actionTypes).toContain('analyze');
    expect(actionTypes).toContain('classify');
    expect(actionTypes).toContain('tag');
    expect(actionTypes).toContain('notify');
    expect(actionTypes).toContain('export');
    expect(actionTypes).toContain('move');
    expect(actionTypes).toContain('delete');
  });

  test('should handle trigger with invalid action type', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'invalid-action' as any }
      ],
      usersToNotify: ['user1', 'user2']
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(1);
    expect(trigger.actions[0].type).toBe('invalid-action');
  });

  test('should handle trigger with empty user list', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: []
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.usersToNotify).toEqual([]);
  });

  test('should handle trigger with null values', () => {
    const triggerData = {
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: ['user1', 'user2'],
      lastRun: null as any,
      autoTaggingConfig: null as any,
      notificationTemplate: null as any
    };

    const trigger = processor.addTrigger(triggerData);
    
    expect(trigger).toBeDefined();
    expect(trigger.usersToNotify).toEqual(['user1', 'user2']);
  });

  test('should get processed documents', () => {
    const processedDocs = processor.getProcessedDocuments();
    expect(processedDocs).toEqual([]);
  });

  test('should handle invalid trigger update', () => {
    const result = processor.updateTrigger('non-existent-id', { enabled: false });
    expect(result).toBe(false);
  });
});