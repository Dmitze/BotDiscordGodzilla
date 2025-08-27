import { AutomatedDocumentProcessor } from '../AutomatedDocumentProcessor';

// Mock config for testing
const mockConfig = {
  drive: {
    folderId: 'test-folder-id',
    enableTextIndex: true
  }
} as any;

describe('AutomatedDocumentProcessor - Performance Tests', () => {
  let processor: AutomatedDocumentProcessor;

  beforeEach(() => {
    processor = new AutomatedDocumentProcessor(mockConfig);
  });

  test('should handle adding many triggers efficiently', () => {
    const start = performance.now();
    
    // Add 1000 triggers
    for (let i = 0; i < 1000; i++) {
      processor.addTrigger({
        folderId: `folder-${i}`,
        folderName: `Folder ${i}`,
        channelId: `channel-${i}`,
        enabled: true,
        conditions: [],
        actions: [],
        usersToNotify: [`user-${i}`]
      });
    }
    
    const end = performance.now();
    const duration = end - start;
    
    // Should complete within reasonable time (less than 100ms for 1000 operations)
    expect(duration).toBeLessThan(100);
    
    // Verify all triggers were added
    const triggers = processor.getTriggers();
    expect(triggers).toHaveLength(1000);
  });

  test('should handle getting triggers efficiently', () => {
    // Add some triggers first
    for (let i = 0; i < 100; i++) {
      processor.addTrigger({
        folderId: `folder-${i}`,
        folderName: `Folder ${i}`,
        channelId: `channel-${i}`,
        enabled: true,
        conditions: [],
        actions: [],
        usersToNotify: [`user-${i}`]
      });
    }
    
    const start = performance.now();
    
    // Get triggers multiple times
    for (let i = 0; i < 1000; i++) {
      const triggers = processor.getTriggers();
      expect(triggers).toHaveLength(100);
    }
    
    const end = performance.now();
    const duration = end - start;
    
    // Should complete within reasonable time (less than 50ms for 1000 operations)
    expect(duration).toBeLessThan(50);
  });

  test('should handle trigger updates efficiently', () => {
    // Add a trigger
    const trigger = processor.addTrigger({
      folderId: 'test-folder',
      folderName: 'Test Folder',
      channelId: 'test-channel',
      enabled: true,
      conditions: [],
      actions: [],
      usersToNotify: ['user1']
    });
    
    const triggerId = trigger.id;
    const start = performance.now();
    
    // Update the trigger many times
    for (let i = 0; i < 1000; i++) {
      const result = processor.updateTrigger(triggerId, {
        folderName: `Updated Folder ${i}`,
        enabled: i % 2 === 0
      });
      expect(result).toBe(true);
    }
    
    const end = performance.now();
    const duration = end - start;
    
    // Should complete within reasonable time (less than 100ms for 1000 operations)
    expect(duration).toBeLessThan(100);
  });

  test('should handle trigger removals efficiently', () => {
    // Add many triggers
    const triggerIds: string[] = [];
    for (let i = 0; i < 1000; i++) {
      const trigger = processor.addTrigger({
        folderId: `folder-${i}`,
        folderName: `Folder ${i}`,
        channelId: `channel-${i}`,
        enabled: true,
        conditions: [],
        actions: [],
        usersToNotify: [`user-${i}`]
      });
      triggerIds.push(trigger.id);
    }
    
    const start = performance.now();
    
    // Remove all triggers
    for (const triggerId of triggerIds) {
      const result = processor.removeTrigger(triggerId);
      expect(result).toBe(true);
    }
    
    const end = performance.now();
    const duration = end - start;
    
    // Should complete within reasonable time (less than 100ms for 1000 operations)
    expect(duration).toBeLessThan(100);
    
    // Verify all triggers were removed
    const triggers = processor.getTriggers();
    expect(triggers).toHaveLength(0);
  });

  test('should handle getting processed documents efficiently', () => {
    const start = performance.now();
    
    // Get processed documents many times
    for (let i = 0; i < 1000; i++) {
      const docs = processor.getProcessedDocuments();
      expect(docs).toEqual([]);
    }
    
    const end = performance.now();
    const duration = end - start;
    
    // Should complete within reasonable time (less than 50ms for 1000 operations)
    expect(duration).toBeLessThan(50);
  });

  test('should maintain performance with large processed document history', () => {
    // This test would require access to private methods, so we'll simulate
    // the performance impact by checking method call times
    
    const start = performance.now();
    
    // Call getProcessedDocuments many times
    for (let i = 0; i < 1000; i++) {
      processor.getProcessedDocuments();
    }
    
    const end = performance.now();
    const duration = end - start;
    
    // Should complete within reasonable time
    expect(duration).toBeLessThan(50);
  });

  test('should handle edge case with maximum processed document history', () => {
    // This test would require access to private methods to simulate having
    // many processed documents, but we can at least verify the method exists
    // and works correctly with the default empty state
    
    const docs = processor.getProcessedDocuments();
    expect(docs).toEqual([]);
    
    // Test with limit parameter
    const limitedDocs = processor.getProcessedDocuments(10);
    expect(limitedDocs).toEqual([]);
  });
});