import { MobileOptimizationService, mobileOptimizationService } from '../MobileOptimizationService';
import { EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } from 'discord.js';

describe('MobileOptimizationService', () => {
  const userId = 'test-user-id';

  beforeEach(() => {
    // Reset the service before each test
    mobileOptimizationService.__reset();
  });

  test('should create default preferences for new user', () => {
    const prefs = mobileOptimizationService.getUserPreferences(userId);
    
    expect(prefs.userId).toBe(userId);
    expect(prefs.enabled).toBe(true);
    expect(prefs.compactMode).toBe(true);
    expect(prefs.maxComponentsPerRow).toBe(3);
    expect(prefs.maxActionRows).toBe(3);
  });

  test('should save and retrieve user preferences', () => {
    const prefs = mobileOptimizationService.getUserPreferences(userId);
    prefs.enabled = false;
    prefs.compactMode = false;
    
    mobileOptimizationService.setUserPreferences(userId, prefs);
    
    const updatedPrefs = mobileOptimizationService.getUserPreferences(userId);
    expect(updatedPrefs.enabled).toBe(false);
    expect(updatedPrefs.compactMode).toBe(false);
  });

  test('should enable or disable mobile optimization', () => {
    // Initially enabled
    let prefs = mobileOptimizationService.getUserPreferences(userId);
    expect(prefs.enabled).toBe(true);
    
    // Disable
    mobileOptimizationService.setEnabled(userId, false);
    prefs = mobileOptimizationService.getUserPreferences(userId);
    expect(prefs.enabled).toBe(false);
    
    // Enable
    mobileOptimizationService.setEnabled(userId, true);
    prefs = mobileOptimizationService.getUserPreferences(userId);
    expect(prefs.enabled).toBe(true);
  });

  test('should optimize embed for mobile viewing', () => {
    const embed = new EmbedBuilder()
      .setTitle('Test Embed')
      .setDescription('This is a test description\n\nWith extra whitespace\n\nAnd more whitespace')
      .addFields(
        { name: 'Field 1', value: 'Value 1' },
        { name: 'Field 2', value: 'Value 2' },
        { name: 'Field 3', value: 'Value 3' },
        { name: 'Field 4', value: 'Value 4' },
        { name: 'Field 5', value: 'Value 5' },
        { name: 'Field 6', value: 'Value 6' }
      );
    
    const optimizedEmbed = mobileOptimizationService.optimizeEmbed(embed, userId);
    
    // Check that description was compacted
    const description = optimizedEmbed.data.description;
    expect(description).toBeDefined();
    expect(description).not.toContain('\n\n\n'); // No triple newlines
    
    // Check that fields were limited
    const fields = optimizedEmbed.data.fields || [];
    expect(fields).toHaveLength(6); // 5 original + 1 "more" indicator
    expect(fields[5].name).toBe('...');
  });

  test('should not optimize embed when disabled', () => {
    // Disable mobile optimization
    mobileOptimizationService.setEnabled(userId, false);
    
    const embed = new EmbedBuilder()
      .setTitle('Test Embed')
      .setDescription('This is a test description\n\nWith extra whitespace\n\nAnd more whitespace');
    
    const optimizedEmbed = mobileOptimizationService.optimizeEmbed(embed, userId);
    
    // Should be the same as original
    expect(optimizedEmbed).toBe(embed);
  });

  test('should optimize action rows for mobile touch interaction', () => {
    // Create test components with many buttons
    const row1 = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('button1')
          .setLabel('Very Long Button Label')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('button2')
          .setLabel('Another Long Label')
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId('button3')
          .setLabel('Third Button')
          .setStyle(ButtonStyle.Success),
        new ButtonBuilder()
          .setCustomId('button4')
          .setLabel('Fourth Button')
          .setStyle(ButtonStyle.Danger)
      );
    
    const row2 = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('button5')
          .setLabel('Fifth Button')
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId('button6')
          .setLabel('Sixth Button')
          .setStyle(ButtonStyle.Secondary)
      );
    
    const components = [row1, row2];
    
    // Apply mobile optimization
    const optimizedComponents = mobileOptimizationService.optimizeActionRows(components, userId);
    
    // Check that components per row were limited
    expect(optimizedComponents[0].components).toHaveLength(3); // Limited to 3
    expect(optimizedComponents[1].components).toHaveLength(2); // Kept as is
    
    // Check that long labels were truncated
    const firstButtonLabel = optimizedComponents[0].components[0].data.label;
    expect(firstButtonLabel).toHaveLength(15); // "Very Long But..."
  });

  test('should create mobile-friendly pagination', () => {
    const paginationRow = mobileOptimizationService.createMobilePagination(2, 5);
    
    // Should have 3 buttons: previous, page info, next
    expect(paginationRow.components).toHaveLength(3);
    
    // Check button labels
    expect(paginationRow.components[0].data.label).toBe('⬅️');
    expect(paginationRow.components[1].data.label).toBe('2/5');
    expect(paginationRow.components[2].data.label).toBe('➡️');
    
    // Check that first page disables previous button
    const firstPageRow = mobileOptimizationService.createMobilePagination(1, 5);
    expect(firstPageRow.components[0].data.disabled).toBe(true);
    
    // Check that last page disables next button
    const lastPageRow = mobileOptimizationService.createMobilePagination(5, 5);
    expect(lastPageRow.components[2].data.disabled).toBe(true);
  });

  test('should optimize text content for mobile screens', () => {
    const longText = 'This is a very long line that exceeds the typical mobile screen width and should be truncated for better mobile viewing experience.';
    const optimizedText = mobileOptimizationService.optimizeTextContent(longText, userId);
    
    // Should be truncated to 80 characters + "..."
    expect(optimizedText).toHaveLength(80);
    expect(optimizedText.endsWith('...')).toBe(true);
  });

  test('should get mobile optimization status', () => {
    // Enabled by default
    let status = mobileOptimizationService.getStatus(userId);
    expect(status.enabled).toBe(true);
    expect(status.mode).toBe('compact');
    
    // Disable and check status
    mobileOptimizationService.setEnabled(userId, false);
    status = mobileOptimizationService.getStatus(userId);
    expect(status.enabled).toBe(false);
    expect(status.mode).toBe('disabled');
  });

  test('should apply all optimizations', () => {
    const embed = new EmbedBuilder()
      .setTitle('Test Embed')
      .setDescription('Test description');
    
    const row = new ActionRowBuilder<ButtonBuilder>()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('test')
          .setLabel('Test Button')
          .setStyle(ButtonStyle.Primary)
      );
    
    const components = [row];
    
    const { embed: optimizedEmbed, components: optimizedComponents } = 
      mobileOptimizationService.applyAllOptimizations(embed, components, userId);
    
    // Should return optimized versions
    expect(optimizedEmbed).toBeDefined();
    expect(optimizedComponents).toHaveLength(1);
  });
});