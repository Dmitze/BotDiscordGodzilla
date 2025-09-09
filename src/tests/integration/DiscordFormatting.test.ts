/**
 * Integration tests for enhanced Discord message formatting
 */

import { jest, describe, it, expect, beforeEach } from '@jest/globals';
import { DiscordMarkdownFormatter } from '../../ui/markdownFormatter';
import { BaseCommand } from '../../commands/BaseCommand';
import { createMockConfig } from '../utils/testHelpers';

// Mock the logger
jest.mock('../../utils/logger', () => ({
  default: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    debug: jest.fn(),
    log: jest.fn(),
    apiRequest: jest.fn(),
    apiError: jest.fn(),
    security: jest.fn(),
    performance: jest.fn(),
    system: jest.fn(),
    logStructured: jest.fn(),
    startStructuredTimer: jest.fn().mockReturnValue({ end: jest.fn() }),
    getStats: jest.fn(),
    getLogBuffer: jest.fn(),
    cleanup: jest.fn(),
    isHealthy: jest.fn(),
  },
}));

// Mock discord.js
jest.mock('discord.js', () => ({
  EmbedBuilder: class {
    constructor() {}
    setTitle(title: string) { this.title = title; return this; }
    setDescription(description: string) { this.description = description; return this; }
    setColor(color: string) { this.color = color; return this; }
    setTimestamp() { this.timestamp = new Date(); return this; }
    addFields(fields: any[]) { this.fields = fields; return this; }
  },
  AttachmentBuilder: class {
    constructor(buffer: Buffer, options: { name: string }) {
      this.buffer = buffer;
      this.name = options.name;
    }
    buffer: Buffer;
    name: string;
  },
  SlashCommandBuilder: class {
    constructor() {
      this.options = [];
    }
    setName(name: string) { this.name = name; return this; }
    setDescription(description: string) { this.description = description; return this; }
    setNameLocalizations(localizations: any) { this.nameLocalizations = localizations; return this; }
    setDescriptionLocalizations(localizations: any) { this.descriptionLocalizations = localizations; return this; }
    setDefaultMemberPermissions(permissions: any) { this.defaultMemberPermissions = permissions; return this; }
    setDMPermission(dmPermission: boolean) { this.dmPermission = dmPermission; return this; }
    addStringOption(option: any) { this.options.push(option); return this; }
    addBooleanOption(option: any) { this.options.push(option); return this; }
    addChoices(choices: any[]) { this.choices = choices; return this; }
    setRequired(required: boolean) { this.required = required; return this; }
    setMaxLength(maxLength: number) { this.maxLength = maxLength; return this; }
  },
}));

describe('Discord Message Formatting Integration', () => {
  let formatter: DiscordMarkdownFormatter;
  let mockConfig: any;
  
  beforeEach(() => {
    mockConfig = createMockConfig();
    formatter = new DiscordMarkdownFormatter(mockConfig);
  });
  
  describe('DiscordMarkdownFormatter integration', () => {
    it('should format markdown with enhanced features', async () => {
      const content = '# Hello World\nThis is **bold** text!';
      
      // Test text formatting
      const textResult = await formatter.formatMarkdown(content, { format: 'text' });
      expect(textResult.content).toBeDefined();
      
      // Test embed formatting
      const embedResult = await formatter.formatMarkdown(content, { format: 'embed' });
      expect(embedResult.embeds).toBeDefined();
    });
    
    it('should create action buttons', () => {
      const content = '# Hello World\n' + 'This is a long content. '.repeat(100);
      const buttons = formatter.createActionButtons(content);
      expect(buttons.length).toBeGreaterThan(0);
    });
  });
});