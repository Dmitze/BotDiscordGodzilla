// Mock the logger
jest.mock('../../utils/logger', () => ({
  info: jest.fn(),
  error: jest.fn()
}));

import { MarkdownRenderService } from '../../services/MarkdownRenderService';
import { AttachmentBuilder } from 'discord.js';

describe('MarkdownRenderService', () => {
  let service: MarkdownRenderService;

  beforeEach(() => {
    service = new MarkdownRenderService();
  });

  test('should render simple markdown to attachment', async () => {
    const markdown = '# Hello World\nThis is a test.';
    const attachment = await service.renderToImage(markdown);
    
    expect(attachment).toBeInstanceOf(AttachmentBuilder);
  });

  test('should render markdown with Discord dark theme', async () => {
    const markdown = '# Hello World\nThis is a test.';
    const attachment = await service.renderDiscordDarkTheme(markdown);
    
    expect(attachment).toBeInstanceOf(AttachmentBuilder);
  });

  test('should render markdown with light theme', async () => {
    const markdown = '# Hello World\nThis is a test.';
    const attachment = await service.renderLightTheme(markdown);
    
    expect(attachment).toBeInstanceOf(AttachmentBuilder);
  });

  test('should handle Ukrainian text correctly', async () => {
    const markdown = '# Привіт, світ!\nЦе тест українського тексту.';
    const attachment = await service.renderDiscordDarkTheme(markdown);
    
    expect(attachment).toBeInstanceOf(AttachmentBuilder);
  });
});