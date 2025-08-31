import { DocLoadCommand } from '../DocLoadCommand';
// Remove unused imports
// import { CommandInteraction, Client, ChatInputCommandInteraction } from 'discord.js';

// Моки для Discord.js
jest.mock('discord.js', () => {
  return {
    SlashCommandBuilder: jest.fn().mockImplementation(() => {
      return {
        setName: jest.fn().mockReturnThis(),
        setDescription: jest.fn().mockReturnThis(),
        setDescriptionLocalizations: jest.fn().mockReturnThis(),
        addStringOption: jest.fn().mockReturnThis(),
        addIntegerOption: jest.fn().mockReturnThis(),
        setDMPermission: jest.fn().mockReturnThis(),
      };
    }),
    EmbedBuilder: jest.fn().mockImplementation(() => {
      return {
        setTitle: jest.fn().mockReturnThis(),
        setDescription: jest.fn().mockReturnThis(),
        addFields: jest.fn().mockReturnThis(),
        setColor: jest.fn().mockReturnThis(),
        setTimestamp: jest.fn().mockReturnThis(),
      };
    }),
    ActionRowBuilder: jest.fn().mockImplementation(() => {
      return {
        addComponents: jest.fn().mockReturnThis(),
      };
    }),
    ButtonBuilder: jest.fn().mockImplementation(() => {
      return {
        setCustomId: jest.fn().mockReturnThis(),
        setLabel: jest.fn().mockReturnThis(),
        setStyle: jest.fn().mockReturnThis(),
      };
    }),
    ButtonStyle: {
      Primary: 1,
      Secondary: 2,
    },
  };
});

describe('DocLoadCommand', () => {
  let docLoadCommand: DocLoadCommand;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = {
      // Порожній об'єкт конфігурації для тестів
    };

    docLoadCommand = new DocLoadCommand(mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('getName', () => {
    it('should return the correct command name', () => {
      expect(docLoadCommand.name).toBe('doc-load');
    });
  });

  describe('getDescription', () => {
    it('should return the correct command description', () => {
      expect(docLoadCommand.description).toBe('Завантажити та проіндексувати Google Docs документ');
    });
  });

  describe('extractDocumentId', () => {
    it('should extract document ID from a valid Google Docs URL', () => {
      const url = 'https://docs.google.com/document/d/1a2b3c4d5e6f7g8h9i0j/edit';
      const result = (docLoadCommand as any).extractDocumentId(url);
      expect(result).toBe('1a2b3c4d5e6f7g8h9i0j');
    });

    it('should return null for invalid URL', () => {
      const url = 'https://example.com/invalid';
      const result = (docLoadCommand as any).extractDocumentId(url);
      expect(result).toBeNull();
    });

    it('should return null for URL without document ID', () => {
      const url = 'https://docs.google.com/document/';
      const result = (docLoadCommand as any).extractDocumentId(url);
      expect(result).toBeNull();
    });
  });

  describe('register', () => {
    it('should register the command with correct options', () => {
      const command = docLoadCommand.register();
      
      // Перевірка, що команда була створена
      expect(command).toBeDefined();
    });
  });
});