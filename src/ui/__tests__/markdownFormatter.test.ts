import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { DiscordMarkdownFormatter } from '../markdownFormatter';

// Mock the logger
jest.mock('@/utils/logger', () => ({
  default: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    debug: jest.fn(),
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
}));

describe('DiscordMarkdownFormatter', () => {
  let formatter: DiscordMarkdownFormatter;
  
  beforeEach(() => {
    formatter = new DiscordMarkdownFormatter({
      // Minimal config for testing
    } as any);
  });
  
  describe('formatMarkdown', () => {
    it('should format simple markdown as text', async () => {
      const content = '# Hello World\nThis is **bold** text!';
      const result = await formatter.formatMarkdown(content, { format: 'text' });
      expect(result.content).toBeDefined();
    });
    
    it.skip('should format markdown as image', async () => {
      const content = '# Hello World\nThis is **bold** text!';
      const result = await formatter.formatMarkdown(content, { format: 'image' });
      expect(result.files).toBeDefined();
      expect(result.files?.length).toBeGreaterThan(0);
    });
    
    it('should format markdown as embed', async () => {
      const content = '# Hello World\nThis is **bold** text!';
      const result = await formatter.formatMarkdown(content, { format: 'embed' });
      expect(result.embeds).toBeDefined();
      expect(result.embeds?.length).toBeGreaterThan(0);
    });
    
    it('should handle Ukrainian language content', async () => {
      const ukrainianContent = '# Привіт Світ\nЦе **жирний** текст!';
      const result = await formatter.formatMarkdown(ukrainianContent, { format: 'text' });
      expect(result.content).toContain('Привіт Світ');
    });
  });
  
  describe('createActionButtons', () => {
    it('should create action buttons for long content', () => {
      const longContent = '# Hello World\n' + 'This is a long content. '.repeat(100);
      const buttons = formatter.createActionButtons(longContent);
      expect(buttons.length).toBeGreaterThan(0);
      expect(buttons.some(b => b.label === 'Show Full Document')).toBeTruthy();
    });
    
    it('should create standard action buttons', () => {
      const content = '# Hello World\nThis is some content.';
      const buttons = formatter.createActionButtons(content);
      expect(buttons.length).toBeGreaterThan(0);
      expect(buttons.some(b => b.label === 'Generate Summary')).toBeTruthy();
      expect(buttons.some(b => b.label === 'Find Similar Documents')).toBeTruthy();
    });
  });
  
  describe('formatCodeBlocks', () => {
    it('should format code blocks correctly', () => {
      const content = 'Some text\n```javascript\nconsole.log("Hello");\n```\nMore text';
      const result = formatter.formatCodeBlocks(content);
      expect(result).toContain('```javascript');
    });
  });
  
  describe('formatTables', () => {
    it('should format tables correctly', () => {
      const content = '| Name | Age |\n|------|-----|\n| John | 30  |';
      const result = formatter.formatTables(content);
      expect(result).toContain('\n| Name | Age |');
    });
  });
});