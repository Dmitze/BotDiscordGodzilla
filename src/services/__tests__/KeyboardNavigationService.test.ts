import { KeyboardNavigationService, keyboardNavigationService } from '../KeyboardNavigationService';
import { ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';

describe('KeyboardNavigationService', () => {
  const userId = 'test-user-id';

  beforeEach(() => {
    // Reset the service before each test
    keyboardNavigationService.__reset();
  });

  test('should create default preferences for new user', () => {
    const prefs = keyboardNavigationService.getUserPreferences(userId);
    
    expect(prefs.userId).toBe(userId);
    expect(prefs.enabled).toBe(true);
    expect(prefs.showHints).toBe(true);
    expect(prefs.shortcuts).toHaveLength(8); // Default shortcuts
  });

  test('should save and retrieve user preferences', () => {
    const prefs = keyboardNavigationService.getUserPreferences(userId);
    prefs.enabled = false;
    
    keyboardNavigationService.setUserPreferences(userId, prefs);
    
    const updatedPrefs = keyboardNavigationService.getUserPreferences(userId);
    expect(updatedPrefs.enabled).toBe(false);
  });

  test('should enable or disable keyboard navigation', () => {
    // Initially enabled
    let prefs = keyboardNavigationService.getUserPreferences(userId);
    expect(prefs.enabled).toBe(true);
    
    // Disable
    keyboardNavigationService.setEnabled(userId, false);
    prefs = keyboardNavigationService.getUserPreferences(userId);
    expect(prefs.enabled).toBe(false);
    
    // Enable
    keyboardNavigationService.setEnabled(userId, true);
    prefs = keyboardNavigationService.getUserPreferences(userId);
    expect(prefs.enabled).toBe(true);
  });

  test('should add numbered labels to buttons', () => {
    // Create test components with buttons
    const row1 = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('button1')
          .setLabel('First Button')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('button2')
          .setLabel('Second Button')
          .setStyle(ButtonStyle.Secondary)
      );
    
    const row2 = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('button3')
          .setLabel('Third Button')
          .setStyle(ButtonStyle.Success)
      );
    
    const components = [row1, row2];
    
    // Apply numbered labels
    const labeledComponents = keyboardNavigationService.addNumberedLabels(components);
    
    // Check that labels were added
    expect(labeledComponents[0].components[0].data.label).toBe('1. First Button');
    expect(labeledComponents[0].components[1].data.label).toBe('2. Second Button');
    expect(labeledComponents[1].components[0].data.label).toBe('3. Third Button');
  });

  test('should not renumber already numbered buttons', () => {
    // Create test components with already numbered buttons
    const row = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('button1')
          .setLabel('1. First Button')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('button2')
          .setLabel('Second Button')
          .setStyle(ButtonStyle.Secondary)
      );
    
    const components = [row];
    
    // Apply numbered labels
    const labeledComponents = keyboardNavigationService.addNumberedLabels(components);
    
    // Check that first button wasn't renumbered
    expect(labeledComponents[0].components[0].data.label).toBe('1. First Button');
    // Check that second button was numbered
    expect(labeledComponents[0].components[1].data.label).toBe('2. Second Button');
  });

  test('should add keyboard shortcut hints to description', () => {
    const originalDescription = 'This is a test description';
    const descriptionWithHints = keyboardNavigationService.addShortcutHints(originalDescription, userId);
    
    // Should contain the original description
    expect(descriptionWithHints).toContain(originalDescription);
    // Should contain keyboard shortcut hints
    expect(descriptionWithHints).toContain('⌨️ Keyboard Shortcuts');
    expect(descriptionWithHints).toContain('1-9 — Select item by number');
  });

  test('should not add hints when disabled', () => {
    // Disable hints
    const prefs = keyboardNavigationService.getUserPreferences(userId);
    prefs.showHints = false;
    keyboardNavigationService.setUserPreferences(userId, prefs);
    
    const originalDescription = 'This is a test description';
    const descriptionWithHints = keyboardNavigationService.addShortcutHints(originalDescription, userId);
    
    // Should only contain the original description
    expect(descriptionWithHints).toBe(originalDescription);
  });

  test('should add custom shortcut', () => {
    const newShortcut = {
      key: 'C',
      description: 'Create new item',
      action: 'create_item'
    };
    
    keyboardNavigationService.addShortcut(userId, newShortcut);
    
    const shortcuts = keyboardNavigationService.getShortcuts(userId);
    expect(shortcuts).toHaveLength(9); // 8 default + 1 custom
    expect(shortcuts.some(s => s.key === 'C')).toBe(true);
  });

  test('should update existing shortcut', () => {
    // First add a shortcut
    const shortcut = {
      key: 'T',
      description: 'Test shortcut',
      action: 'test'
    };
    
    keyboardNavigationService.addShortcut(userId, shortcut);
    
    // Update the shortcut
    const updatedShortcut = {
      key: 'T',
      description: 'Updated test shortcut',
      action: 'test_updated'
    };
    
    keyboardNavigationService.addShortcut(userId, updatedShortcut);
    
    const shortcuts = keyboardNavigationService.getShortcuts(userId);
    expect(shortcuts).toHaveLength(9); // 8 default + 1 custom
    const testShortcut = shortcuts.find(s => s.key === 'T');
    expect(testShortcut?.description).toBe('Updated test shortcut');
  });

  test('should remove shortcut', () => {
    // Add a custom shortcut
    const shortcut = {
      key: 'X',
      description: 'Custom shortcut',
      action: 'custom'
    };
    
    keyboardNavigationService.addShortcut(userId, shortcut);
    
    // Remove the shortcut
    const result = keyboardNavigationService.removeShortcut(userId, 'X');
    expect(result).toBe(true);
    
    // Check that the shortcut was removed
    const shortcuts = keyboardNavigationService.getShortcuts(userId);
    expect(shortcuts).toHaveLength(8); // Back to default
    expect(shortcuts.some(s => s.key === 'X')).toBe(false);
  });

  test('should generate help text', () => {
    const helpText = keyboardNavigationService.generateHelpText(userId);
    
    expect(helpText).toContain('⌨️ Keyboard Navigation Help');
    expect(helpText).toContain('1-9 — Select item by number');
  });

  test('should generate help text when disabled', () => {
    keyboardNavigationService.setEnabled(userId, false);
    const helpText = keyboardNavigationService.generateHelpText(userId);
    
    expect(helpText).toContain('Keyboard navigation is currently disabled');
  });
});