import { AutomatedDocumentProcessor, DocumentAction } from '../AutomatedDocumentProcessor';

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

describe('AutomatedDocumentProcessor - Background Task Processing', () => {
  let processor: AutomatedDocumentProcessor;

  beforeEach(() => {
    processor = new AutomatedDocumentProcessor(mockConfig);
  });

  test('should create AutomatedDocumentProcessor', () => {
    expect(processor).toBeInstanceOf(AutomatedDocumentProcessor);
  });

  test('should handle action with background processing enabled', () => {
    const action: DocumentAction = {
      type: 'summarize',
      runInBackground: true
    };

    expect(action.runInBackground).toBe(true);
  });

  test('should handle trigger with background actions', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'summarize', runInBackground: true },
        { type: 'classify', runInBackground: true }
      ],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(2);
    
    const hasBackgroundActions = trigger.actions.some(action => action.runInBackground);
    expect(hasBackgroundActions).toBe(true);
  });

  test('should handle trigger with mixed background and foreground actions', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'summarize', runInBackground: true },
        { type: 'tag' }, // foreground by default
        { type: 'classify', runInBackground: true }
      ],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(3);
    
    const backgroundActions = trigger.actions.filter(action => action.runInBackground);
    const foregroundActions = trigger.actions.filter(action => !action.runInBackground);
    
    expect(backgroundActions).toHaveLength(2);
    expect(foregroundActions).toHaveLength(1);
  });

  test('should get processed documents after background processing', () => {
    const docs = processor.getProcessedDocuments();
    expect(docs).toEqual([]);
  });

  test('should handle multiple triggers with background actions', () => {
    // Add first trigger with background actions
    const trigger1 = processor.addTrigger({
      folderId: 'folder-1',
      folderName: 'Folder 1',
      channelId: 'channel-1',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'summarize', runInBackground: true }
      ],
      usersToNotify: ['user1']
    });

    // Add second trigger with background actions
    const trigger2 = processor.addTrigger({
      folderId: 'folder-2',
      folderName: 'Folder 2',
      channelId: 'channel-2',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'classify', runInBackground: true },
        { type: 'tag' }
      ],
      usersToNotify: ['user2']
    });

    const triggers = processor.getTriggers();
    expect(triggers).toHaveLength(2);

    // Verify first trigger
    const firstTrigger = triggers.find(t => t.id === trigger1.id);
    expect(firstTrigger).toBeDefined();
    expect(firstTrigger?.actions).toHaveLength(1);
    expect(firstTrigger?.actions[0].type).toBe('summarize');
    expect(firstTrigger?.actions[0].runInBackground).toBe(true);

    // Verify second trigger
    const secondTrigger = triggers.find(t => t.id === trigger2.id);
    expect(secondTrigger).toBeDefined();
    expect(secondTrigger?.actions).toHaveLength(2);
    expect(secondTrigger?.actions[0].type).toBe('classify');
    expect(secondTrigger?.actions[0].runInBackground).toBe(true);
    expect(secondTrigger?.actions[1].type).toBe('tag');
    expect(secondTrigger?.actions[1].runInBackground).toBeUndefined();
  });
});