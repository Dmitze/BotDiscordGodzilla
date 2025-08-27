import { AutomatedDocumentProcessor } from '../AutomatedDocumentProcessor';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    enableTextIndex: true
  }
} as any;

describe('AutomatedDocumentProcessor - Document Version Comparison', () => {
  let processor: AutomatedDocumentProcessor;

  beforeEach(() => {
    processor = new AutomatedDocumentProcessor(mockConfig);
  });

  test('should create AutomatedDocumentProcessor', () => {
    expect(processor).toBeInstanceOf(AutomatedDocumentProcessor);
  });

  test('should handle compare_versions action in trigger', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [{ type: 'compare_versions' }],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(1);
    expect(trigger.actions[0].type).toBe('compare_versions');
  });

  test('should handle compare_versions action with parameters', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [{ type: 'compare_versions', parameters: { compareAll: true } }],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(1);
    expect(trigger.actions[0].type).toBe('compare_versions');
    expect(trigger.actions[0].parameters).toEqual({ compareAll: true });
  });

  test('should get processed documents with version comparison', () => {
    const docs = processor.getProcessedDocuments();
    expect(docs).toEqual([]);
  });

  test('should handle multiple triggers with compare_versions actions', () => {
    // Add first trigger with compare_versions
    const trigger1 = processor.addTrigger({
      folderId: 'folder-1',
      folderName: 'Folder 1',
      channelId: 'channel-1',
      enabled: true,
      conditions: [],
      actions: [{ type: 'compare_versions' }],
      usersToNotify: ['user1']
    });

    // Add second trigger with compare_versions
    const trigger2 = processor.addTrigger({
      folderId: 'folder-2',
      folderName: 'Folder 2',
      channelId: 'channel-2',
      enabled: true,
      conditions: [],
      actions: [{ type: 'compare_versions' }, { type: 'tag' }],
      usersToNotify: ['user2']
    });

    const triggers = processor.getTriggers();
    expect(triggers).toHaveLength(2);

    // Verify first trigger
    const firstTrigger = triggers.find(t => t.id === trigger1.id);
    expect(firstTrigger).toBeDefined();
    expect(firstTrigger?.actions).toHaveLength(1);
    expect(firstTrigger?.actions[0].type).toBe('compare_versions');

    // Verify second trigger
    const secondTrigger = triggers.find(t => t.id === trigger2.id);
    expect(secondTrigger).toBeDefined();
    expect(secondTrigger?.actions).toHaveLength(2);
    expect(secondTrigger?.actions[0].type).toBe('compare_versions');
    expect(secondTrigger?.actions[1].type).toBe('tag');
  });

  test('should handle trigger with summarize and compare_versions actions', () => {
    const trigger = processor.addTrigger({
      folderId: 'test-folder-id',
      folderName: 'Test Folder',
      channelId: 'test-channel-id',
      enabled: true,
      conditions: [],
      actions: [
        { type: 'summarize' },
        { type: 'compare_versions' }
      ],
      usersToNotify: ['user1']
    });

    expect(trigger).toBeDefined();
    expect(trigger.actions).toHaveLength(2);
    expect(trigger.actions[0].type).toBe('summarize');
    expect(trigger.actions[1].type).toBe('compare_versions');
  });
});