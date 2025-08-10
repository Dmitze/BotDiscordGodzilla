/**
 * Тести для загальних допоміжних функцій
 * Версія 1.0.0 - Комплексне тестування утілітарних класів
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { ChatInputCommandInteraction, EmbedBuilder, User, Guild } from 'discord.js';
import {
  EmbedFactory,
  TimeUtils,
  ValidationUtils,
  DataUtils,
  DiscordUtils,
  ErrorUtils,
  RetryUtils,
  EMBED_COLORS,
  TIME_CONSTANTS
} from '@/utils/commonHelpers';

// Mock logger for error utils
jest.mock('@/utils/logger');

describe('CommonHelpers', () => {
  describe('EmbedFactory', () => {
    describe('createBase', () => {
      it('should create a basic embed with title and description', () => {
        const embed = EmbedFactory.createBase('Test Title', 'Test Description');
        
        expect(embed).toBeInstanceOf(EmbedBuilder);
        expect(embed.data.title).toBe('Test Title');
        expect(embed.data.description).toBe('Test Description');
        expect(embed.data.color).toBe(EMBED_COLORS.PRIMARY);
        expect(embed.data.timestamp).toBeDefined();
        expect(embed.data.footer?.text).toBe('Discord AI Assistant Bot');
      });

      it('should truncate long titles and descriptions', () => {
        const longTitle = 'A'.repeat(300);
        const longDescription = 'B'.repeat(5000);
        
        const embed = EmbedFactory.createBase(longTitle, longDescription);
        
        expect(embed.data.title?.length).toBeLessThanOrEqual(256);
        expect(embed.data.description?.length).toBeLessThanOrEqual(4096);
        expect(embed.data.title).toContain('...');
        expect(embed.data.description).toContain('...');
      });

      it('should use custom color when provided', () => {
        const customColor = 0xFF5733;
        const embed = EmbedFactory.createBase('Title', 'Description', customColor);
        
        expect(embed.data.color).toBe(customColor);
      });
    });

    describe('success', () => {
      it('should create success embed with green color', () => {
        const embed = EmbedFactory.success('Success', 'Operation completed');
        
        expect(embed.data.title).toBe('✅ Success');
        expect(embed.data.color).toBe(EMBED_COLORS.SUCCESS);
      });
    });

    describe('error', () => {
      it('should create error embed with red color', () => {
        const embed = EmbedFactory.error('Error', 'Something went wrong');
        
        expect(embed.data.title).toBe('❌ Error');
        expect(embed.data.color).toBe(EMBED_COLORS.ERROR);
      });

      it('should include support field by default', () => {
        const embed = EmbedFactory.error('Error', 'Something went wrong');
        
        expect(embed.data.fields?.some(field => 
          field.name === '📞 Потрібна допомога?'
        )).toBe(true);
      });

      it('should exclude support field when specified', () => {
        const embed = EmbedFactory.error('Error', 'Something went wrong', false);
        
        expect(embed.data.fields?.some(field => 
          field.name === '📞 Потрібна допомога?'
        )).toBe(false);
      });
    });

    describe('dataFields', () => {
      it('should create embed with data fields', () => {
        const fields = [
          { name: 'Field 1', value: 'Value 1', inline: true },
          { name: 'Field 2', value: 'Value 2', inline: false }
        ];
        
        const embed = EmbedFactory.dataFields('Data', 'Information', fields);
        
        expect(embed.data.fields).toHaveLength(2);
        expect(embed.data.fields?.[0].name).toBe('Field 1');
        expect(embed.data.fields?.[0].inline).toBe(true);
        expect(embed.data.fields?.[1].inline).toBe(false);
      });

      it('should limit fields to maximum allowed', () => {
        const manyFields = Array.from({ length: 30 }, (_, i) => ({
          name: `Field ${i}`,
          value: `Value ${i}`
        }));
        
        const embed = EmbedFactory.dataFields('Data', 'Information', manyFields);
        
        expect(embed.data.fields?.length).toBeLessThanOrEqual(25);
      });

      it('should truncate long field names and values', () => {
        const fields = [
          { 
            name: 'A'.repeat(300), 
            value: 'B'.repeat(2000) 
          }
        ];
        
        const embed = EmbedFactory.dataFields('Data', 'Information', fields);
        
        expect(embed.data.fields?.[0].name.length).toBeLessThanOrEqual(256);
        expect(embed.data.fields?.[0].value.length).toBeLessThanOrEqual(1024);
      });
    });

    describe('paginated', () => {
      it('should create paginated embed with page info', () => {
        const embed = EmbedFactory.paginated('Results', 'Page content', 2, 5);
        
        expect(embed.data.footer?.text).toContain('Сторінка 2/5');
      });
    });
  });

  describe('TimeUtils', () => {
    describe('formatDuration', () => {
      it('should format milliseconds correctly', () => {
        expect(TimeUtils.formatDuration(500)).toBe('500ms');
        expect(TimeUtils.formatDuration(1500)).toBe('2s');
        expect(TimeUtils.formatDuration(65000)).toBe('1m 5s');
        expect(TimeUtils.formatDuration(3665000)).toBe('1h 1m');
        expect(TimeUtils.formatDuration(90061000)).toBe('1d 1h');
      });

      it('should handle zero and negative values', () => {
        expect(TimeUtils.formatDuration(0)).toBe('0ms');
        expect(TimeUtils.formatDuration(-1000)).toBe('-1000ms');
      });

      it('should not show zero components', () => {
        expect(TimeUtils.formatDuration(60000)).toBe('1m');
        expect(TimeUtils.formatDuration(3600000)).toBe('1h');
        expect(TimeUtils.formatDuration(86400000)).toBe('1d');
      });
    });

    describe('formatTimestamp', () => {
      const now = Date.now();
      const minute = TIME_CONSTANTS.MINUTE;
      const hour = TIME_CONSTANTS.HOUR;
      const day = TIME_CONSTANTS.DAY;

      it('should format relative timestamps', () => {
        expect(TimeUtils.formatTimestamp(now - 30000, 'relative')).toBe('щойно');
        expect(TimeUtils.formatTimestamp(now - minute * 5, 'relative')).toBe('5 хв тому');
        expect(TimeUtils.formatTimestamp(now - hour * 2, 'relative')).toBe('2 год тому');
        expect(TimeUtils.formatTimestamp(now - day * 3, 'relative')).toBe('3 дн тому');
      });

      it('should format absolute dates', () => {
        const timestamp = new Date('2024-01-15').getTime();
        const result = TimeUtils.formatTimestamp(timestamp, 'absolute');
        expect(result).toMatch(/\d{1,2}\.\d{1,2}\.\d{4}/);
      });

      it('should format datetime', () => {
        const timestamp = new Date('2024-01-15T12:30:00').getTime();
        const result = TimeUtils.formatTimestamp(timestamp, 'datetime');
        expect(result).toContain('2024');
        expect(result).toContain(':');
      });
    });

    describe('isWithinRange', () => {
      it('should check if timestamp is within range', () => {
        const now = Date.now();
        expect(TimeUtils.isWithinRange(now - 1000, 2000)).toBe(true);
        expect(TimeUtils.isWithinRange(now - 3000, 2000)).toBe(false);
      });
    });

    describe('getTimeUntilNextInterval', () => {
      it('should calculate time until next interval', () => {
        const result = TimeUtils.getTimeUntilNextInterval(60000); // 1 minute
        expect(result).toBeGreaterThan(0);
        expect(result).toBeLessThanOrEqual(60000);
      });
    });
  });

  describe('ValidationUtils', () => {
    describe('isValidEmail', () => {
      it('should validate email addresses', () => {
        expect(ValidationUtils.isValidEmail('test@example.com')).toBe(true);
        expect(ValidationUtils.isValidEmail('user.name+tag@domain.co.uk')).toBe(true);
        expect(ValidationUtils.isValidEmail('invalid-email')).toBe(false);
        expect(ValidationUtils.isValidEmail('test@')).toBe(false);
        expect(ValidationUtils.isValidEmail('@example.com')).toBe(false);
      });
    });

    describe('isValidURL', () => {
      it('should validate URLs', () => {
        expect(ValidationUtils.isValidURL('https://example.com')).toBe(true);
        expect(ValidationUtils.isValidURL('http://localhost:3000')).toBe(true);
        expect(ValidationUtils.isValidURL('ftp://files.example.com')).toBe(true);
        expect(ValidationUtils.isValidURL('not-a-url')).toBe(false);
        expect(ValidationUtils.isValidURL('://missing-protocol')).toBe(false);
      });
    });

    describe('isValidDiscordId', () => {
      it('should validate Discord IDs', () => {
        expect(ValidationUtils.isValidDiscordId('123456789012345678')).toBe(true);
        expect(ValidationUtils.isValidDiscordId('1234567890123456789')).toBe(true);
        expect(ValidationUtils.isValidDiscordId('12345')).toBe(false);
        expect(ValidationUtils.isValidDiscordId('abc123')).toBe(false);
        expect(ValidationUtils.isValidDiscordId('')).toBe(false);
      });
    });

    describe('sanitizeText', () => {
      it('should sanitize dangerous characters', () => {
        expect(ValidationUtils.sanitizeText('<script>alert("xss")</script>'))
          .toBe('scriptalert("xss")/script');
        expect(ValidationUtils.sanitizeText('  text with spaces  '))
          .toBe('text with spaces');
      });

      it('should neutralize mass mentions', () => {
        expect(ValidationUtils.sanitizeText('@everyone hello'))
          .toBe('@​everyone hello');
        expect(ValidationUtils.sanitizeText('@here test'))
          .toBe('@​here test');
      });
    });

    describe('isInRange', () => {
      it('should check if number is in range', () => {
        expect(ValidationUtils.isInRange(5, 1, 10)).toBe(true);
        expect(ValidationUtils.isInRange(1, 1, 10)).toBe(true);
        expect(ValidationUtils.isInRange(10, 1, 10)).toBe(true);
        expect(ValidationUtils.isInRange(0, 1, 10)).toBe(false);
        expect(ValidationUtils.isInRange(11, 1, 10)).toBe(false);
      });
    });

    describe('isValidLength', () => {
      it('should check string length', () => {
        expect(ValidationUtils.isValidLength('hello', 3, 10)).toBe(true);
        expect(ValidationUtils.isValidLength('hi', 3, 10)).toBe(false);
        expect(ValidationUtils.isValidLength('very long string', 3, 10)).toBe(false);
      });
    });
  });

  describe('DataUtils', () => {
    describe('formatFileSize', () => {
      it('should format file sizes correctly', () => {
        expect(DataUtils.formatFileSize(0)).toBe('0.0 B');
        expect(DataUtils.formatFileSize(1024)).toBe('1.0 KB');
        expect(DataUtils.formatFileSize(1024 * 1024)).toBe('1.0 MB');
        expect(DataUtils.formatFileSize(1024 * 1024 * 1024)).toBe('1.0 GB');
        expect(DataUtils.formatFileSize(1536)).toBe('1.5 KB');
      });
    });

    describe('formatNumber', () => {
      it('should format numbers with locale-specific separators', () => {
        expect(DataUtils.formatNumber(1234567)).toContain('1');
        expect(DataUtils.formatNumber(1000)).toContain('1');
      });
    });

    describe('formatPercentage', () => {
      it('should calculate and format percentages', () => {
        expect(DataUtils.formatPercentage(25, 100)).toBe('25.0%');
        expect(DataUtils.formatPercentage(1, 3, 2)).toBe('33.33%');
        expect(DataUtils.formatPercentage(0, 0)).toBe('0.0%');
      });
    });

    describe('deepClone', () => {
      it('should clone primitive values', () => {
        expect(DataUtils.deepClone(5)).toBe(5);
        expect(DataUtils.deepClone('hello')).toBe('hello');
        expect(DataUtils.deepClone(null)).toBe(null);
        expect(DataUtils.deepClone(undefined)).toBe(undefined);
      });

      it('should clone arrays', () => {
        const original = [1, 2, [3, 4]];
        const cloned = DataUtils.deepClone(original);
        
        expect(cloned).toEqual(original);
        expect(cloned).not.toBe(original);
        expect(cloned[2]).not.toBe(original[2]);
      });

      it('should clone objects', () => {
        const original = { a: 1, b: { c: 2 } };
        const cloned = DataUtils.deepClone(original);
        
        expect(cloned).toEqual(original);
        expect(cloned).not.toBe(original);
        expect(cloned.b).not.toBe(original.b);
      });

      it('should clone dates', () => {
        const original = new Date('2024-01-01');
        const cloned = DataUtils.deepClone(original);
        
        expect(cloned).toEqual(original);
        expect(cloned).not.toBe(original);
      });
    });

    describe('safeJsonParse', () => {
      it('should parse valid JSON', () => {
        expect(DataUtils.safeJsonParse('{"a": 1}', {})).toEqual({ a: 1 });
        expect(DataUtils.safeJsonParse('[1, 2, 3]', [])).toEqual([1, 2, 3]);
      });

      it('should return default value for invalid JSON', () => {
        expect(DataUtils.safeJsonParse('invalid json', { default: true }))
          .toEqual({ default: true });
        expect(DataUtils.safeJsonParse('', 'default')).toBe('default');
      });
    });

    describe('groupBy', () => {
      it('should group array elements by key function', () => {
        const items = [
          { type: 'A', value: 1 },
          { type: 'B', value: 2 },
          { type: 'A', value: 3 }
        ];
        
        const grouped = DataUtils.groupBy(items, item => item.type);
        
        expect(grouped.A).toHaveLength(2);
        expect(grouped.B).toHaveLength(1);
        expect(grouped.A[0].value).toBe(1);
        expect(grouped.A[1].value).toBe(3);
      });
    });

    describe('paginate', () => {
      it('should paginate arrays correctly', () => {
        const items = Array.from({ length: 25 }, (_, i) => i);
        
        const page1 = DataUtils.paginate(items, 1, 10);
        expect(page1.items).toHaveLength(10);
        expect(page1.currentPage).toBe(1);
        expect(page1.totalPages).toBe(3);
        expect(page1.hasNext).toBe(true);
        expect(page1.hasPrev).toBe(false);
        
        const page2 = DataUtils.paginate(items, 2, 10);
        expect(page2.hasNext).toBe(true);
        expect(page2.hasPrev).toBe(true);
        
        const page3 = DataUtils.paginate(items, 3, 10);
        expect(page3.items).toHaveLength(5);
        expect(page3.hasNext).toBe(false);
        expect(page3.hasPrev).toBe(true);
      });

      it('should handle edge cases', () => {
        const emptyResult = DataUtils.paginate([], 1, 10);
        expect(emptyResult.items).toHaveLength(0);
        expect(emptyResult.totalPages).toBe(0);
        
        const invalidPage = DataUtils.paginate([1, 2, 3], 10, 10);
        expect(invalidPage.currentPage).toBe(1);
      });
    });
  });

  describe('DiscordUtils', () => {
    let mockInteraction: ChatInputCommandInteraction;

    beforeEach(() => {
      mockInteraction = {
        replied: false,
        deferred: false,
        reply: jest.fn().mockResolvedValue(undefined),
        followUp: jest.fn().mockResolvedValue(undefined)
      } as unknown as ChatInputCommandInteraction;
    });

    describe('safeReply', () => {
      it('should reply to interaction when not replied', async () => {
        const result = await DiscordUtils.safeReply(mockInteraction, {
          content: 'Test message'
        });
        
        expect(result).toBe(true);
        expect(mockInteraction.reply).toHaveBeenCalledWith({
          content: 'Test message'
        });
      });

      it('should follow up when already replied', async () => {
        mockInteraction.replied = true;
        
        const result = await DiscordUtils.safeReply(mockInteraction, {
          content: 'Test message'
        });
        
        expect(result).toBe(true);
        expect(mockInteraction.followUp).toHaveBeenCalledWith({
          content: 'Test message'
        });
      });

      it('should handle errors gracefully', async () => {
        (mockInteraction.reply as jest.Mock).mockRejectedValue(new Error('API Error'));
        
        const result = await DiscordUtils.safeReply(mockInteraction, {
          content: 'Test message'
        });
        
        expect(result).toBe(false);
      });
    });

    describe('getUserDisplayName', () => {
      it('should return display name from guild member', () => {
        const user = { username: 'user', displayName: 'User' } as User;
        const guild = {
          members: {
            cache: new Map([['123', { displayName: 'Guild User' }]])
          }
        } as unknown as Guild;
        
        user.id = '123';
        const name = DiscordUtils.getUserDisplayName(user, guild);
        expect(name).toBe('Guild User');
      });

      it('should fallback to user display name', () => {
        const user = { 
          id: '456', 
          username: 'user', 
          displayName: 'User' 
        } as User;
        const guild = {
          members: {
            cache: new Map()
          }
        } as unknown as Guild;
        
        const name = DiscordUtils.getUserDisplayName(user, guild);
        expect(name).toBe('User');
      });

      it('should work without guild', () => {
        const user = { username: 'user', displayName: 'User' } as User;
        const name = DiscordUtils.getUserDisplayName(user);
        expect(name).toBe('User');
      });
    });

    describe('Format mention methods', () => {
      it('should format user mentions', () => {
        expect(DiscordUtils.formatUserMention('123456789'))
          .toBe('<@123456789>');
      });

      it('should format channel mentions', () => {
        expect(DiscordUtils.formatChannelMention('987654321'))
          .toBe('<#987654321>');
      });

      it('should format role mentions', () => {
        expect(DiscordUtils.formatRoleMention('555555555'))
          .toBe('<@&555555555>');
      });
    });
  });

  describe('ErrorUtils', () => {
    beforeEach(() => {
      jest.clearAllMocks();
    });

    describe('logError', () => {
      it('should log errors with context', () => {
        const mockLogger = require('@/utils/logger');
        const error = new Error('Test error');
        
        ErrorUtils.logError(error, {
          operation: 'test-operation',
          userId: '123',
          commandName: 'test-command'
        });
        
        expect(mockLogger.error).toHaveBeenCalledWith(
          expect.stringContaining('test-operation'),
          expect.objectContaining({
            error: 'Test error',
            userId: '123',
            command: 'test-command'
          })
        );
      });
    });

    describe('createErrorEmbed', () => {
      it('should create error embed without details by default', () => {
        const error = new Error('Sensitive error message');
        const embed = ErrorUtils.createErrorEmbed(error);
        
        expect(embed.data.title).toContain('❌');
        expect(embed.data.description).not.toContain('Sensitive error message');
        expect(embed.data.description).toContain('неочікувана помилка');
      });

      it('should show details when requested', () => {
        const error = new Error('Detailed error message');
        const embed = ErrorUtils.createErrorEmbed(error, true);
        
        expect(embed.data.description).toContain('Detailed error message');
      });
    });

    describe('isCriticalError', () => {
      it('should identify critical errors', () => {
        expect(ErrorUtils.isCriticalError(new Error('ECONNREFUSED')))
          .toBe(true);
        expect(ErrorUtils.isCriticalError(new Error('Database connection failed')))
          .toBe(true);
        expect(ErrorUtils.isCriticalError(new Error('Permission denied')))
          .toBe(true);
        expect(ErrorUtils.isCriticalError(new Error('Regular error')))
          .toBe(false);
      });

      it('should handle non-Error objects', () => {
        expect(ErrorUtils.isCriticalError('string error')).toBe(false);
        expect(ErrorUtils.isCriticalError(null)).toBe(false);
      });
    });
  });

  describe('RetryUtils', () => {
    describe('withRetry', () => {
      it('should succeed on first try', async () => {
        const mockFn = jest.fn().mockResolvedValue('success');
        
        const result = await RetryUtils.withRetry(mockFn);
        
        expect(result).toBe('success');
        expect(mockFn).toHaveBeenCalledTimes(1);
      });

      it('should retry on failure and eventually succeed', async () => {
        const mockFn = jest.fn()
          .mockRejectedValueOnce(new Error('First failure'))
          .mockRejectedValueOnce(new Error('Second failure'))
          .mockResolvedValue('success');
        
        const result = await RetryUtils.withRetry(mockFn, {
          maxAttempts: 3,
          delay: 10
        });
        
        expect(result).toBe('success');
        expect(mockFn).toHaveBeenCalledTimes(3);
      });

      it('should fail after max attempts', async () => {
        const mockFn = jest.fn().mockRejectedValue(new Error('Persistent failure'));
        
        await expect(RetryUtils.withRetry(mockFn, {
          maxAttempts: 2,
          delay: 10
        })).rejects.toThrow('Persistent failure');
        
        expect(mockFn).toHaveBeenCalledTimes(2);
      });

      it('should respect shouldRetry predicate', async () => {
        const mockFn = jest.fn().mockRejectedValue(new Error('Non-retryable'));
        
        await expect(RetryUtils.withRetry(mockFn, {
          maxAttempts: 3,
          shouldRetry: () => false
        })).rejects.toThrow('Non-retryable');
        
        expect(mockFn).toHaveBeenCalledTimes(1);
      });

      it('should use exponential backoff', async () => {
        const mockFn = jest.fn()
          .mockRejectedValueOnce(new Error('First'))
          .mockRejectedValueOnce(new Error('Second'))
          .mockResolvedValue('success');
        
        const startTime = Date.now();
        
        await RetryUtils.withRetry(mockFn, {
          maxAttempts: 3,
          delay: 10,
          backoff: 'exponential'
        });
        
        const duration = Date.now() - startTime;
        expect(duration).toBeGreaterThan(30); // 10 + 20 = 30ms minimum
      });

      it('should use linear backoff', async () => {
        const mockFn = jest.fn()
          .mockRejectedValueOnce(new Error('First'))
          .mockRejectedValueOnce(new Error('Second'))
          .mockResolvedValue('success');
        
        const startTime = Date.now();
        
        await RetryUtils.withRetry(mockFn, {
          maxAttempts: 3,
          delay: 10,
          backoff: 'linear'
        });
        
        const duration = Date.now() - startTime;
        expect(duration).toBeGreaterThan(30); // 10 + 20 = 30ms minimum
      });
    });
  });

  describe('Constants', () => {
    it('should have correct embed colors', () => {
      expect(EMBED_COLORS.SUCCESS).toBe(0x00FF00);
      expect(EMBED_COLORS.ERROR).toBe(0xFF0000);
      expect(EMBED_COLORS.WARNING).toBe(0xFFA500);
      expect(EMBED_COLORS.INFO).toBe(0x0099FF);
      expect(EMBED_COLORS.PRIMARY).toBe(0x00AE86);
    });

    it('should have correct time constants', () => {
      expect(TIME_CONSTANTS.SECOND).toBe(1000);
      expect(TIME_CONSTANTS.MINUTE).toBe(60000);
      expect(TIME_CONSTANTS.HOUR).toBe(3600000);
      expect(TIME_CONSTANTS.DAY).toBe(86400000);
      expect(TIME_CONSTANTS.WEEK).toBe(604800000);
    });
  });
});