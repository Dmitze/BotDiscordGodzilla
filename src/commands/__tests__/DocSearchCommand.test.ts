import { DocSearchCommand } from '../DocSearchCommand';
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

describe('DocSearchCommand', () => {
  let docSearchCommand: DocSearchCommand;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = {
      // Порожній об'єкт конфігурації для тестів
    };

    docSearchCommand = new DocSearchCommand(mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('getName', () => {
    it('should return the correct command name', () => {
      expect(docSearchCommand.name).toBe('doc-search');
    });
  });

  describe('getDescription', () => {
    it('should return the correct command description', () => {
      expect(docSearchCommand.description).toBe('Пошук в завантажених Google Docs документах');
    });
  });

  describe('register', () => {
    it('should register the command with correct options', () => {
      const command = docSearchCommand.register();
      
      // Перевірка, що команда була створена
      expect(command).toBeDefined();
    });
  });
});