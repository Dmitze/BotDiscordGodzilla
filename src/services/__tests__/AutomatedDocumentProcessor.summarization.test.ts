import { AutomatedDocumentProcessor } from '../AutomatedDocumentProcessor';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    enableTextIndex: true
  }
} as any;

describe('AutomatedDocumentProcessor - Document Summarization', () => {
  let processor: AutomatedDocumentProcessor;

  beforeEach(() => {
    processor = new AutomatedDocumentProcessor(mockConfig);
  });

  test('should create AutomatedDocumentProcessor', () => {
    expect(processor).toBeInstanceOf(AutomatedDocumentProcessor);
  });

  test('should handle summarize action in trigger', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [{ type: 'summarize' }],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(1);
    expect(trigger.actions[0].type).toBe('summarize');
  });

  test('should handle summarize action with parameters', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [{ type: 'summarize', parameters: { maxLength: 500 } }],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(1);
    expect(trigger.actions[0].type).toBe('summarize');
    expect(trigger.actions[0].parameters).toEqual({ maxLength: 500 });
  });

  test('should get processed documents with summary', () => {
    const docs = processor.getProcessedDocuments();
    expect(docs).toEqual([]);
  });

  test('should handle multiple triggers with summarize actions', () => {
    // Add first trigger with summarize
    const trigger1 = processor.addTrigger({
      folderId: 'folder-1',
      folderName: 'Folder 1',
      channelId: 'channel-1',
      enabled: true,
      conditions: [],
      actions: [{ type: 'summarize' }],
      usersToNotify: ['user1']
    });

    // Add second trigger with summarize
    const trigger2 = processor.addTrigger({
      folderId: 'folder-2',
      folderName: 'Folder 2',
      channelId: 'channel-2',
      enabled: true,
      conditions: [],
      actions: [{ type: 'summarize' }, { type: 'tag' }],
      usersToNotify: ['user2']
    });

    const triggers = processor.getTriggers();
    expect(triggers).toHaveLength(2);

    // Verify first trigger
    const firstTrigger = triggers.find(t => t.id === trigger1.id);
    expect(firstTrigger).toBeDefined();
    expect(firstTrigger?.actions).toHaveLength(1);
    expect(firstTrigger?.actions[0].type).toBe('summarize');

    // Verify second trigger
    const secondTrigger = triggers.find(t => t.id === trigger2.id);
    expect(secondTrigger).toBeDefined();
    expect(secondTrigger?.actions).toHaveLength(2);
    expect(secondTrigger?.actions[0].type).toBe('summarize');
    expect(secondTrigger?.actions[1].type).toBe('tag');
  });
});