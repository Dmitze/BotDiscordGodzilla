import { MobileCommand } from '../MobileCommand';
import { mobileOptimizationService } from '@/services/MobileOptimizationService';
import { BaseCommand } from '@/commands/BaseCommand';
import type { BotConfig } from '@/types';
import { t } from '@/i18n';

// Mock the i18n module
jest.mock('@/i18n', () => ({
  t: jest.fn().mockImplementation((key) => `translated:${key}`)
}));

// Mock the mobile optimization service
jest.mock('@/services/MobileOptimizationService', () => ({
  mobileOptimizationService: {
    setEnabled: jest.fn(),
    getStatus: jest.fn(),
    getUserPreferences: jest.fn(),
    setUserPreferences: jest.fn()
  }
}));

// Mock the reply utility
jest.mock('@/ui/reply', () => ({
  replyWithPrivacy: jest.fn()
}));

// Mock the UIHelper
jest.mock('@/utils/uiHelpers', () => ({
  UIHelper: {
    createBaseEmbed: jest.fn().mockImplementation((title, description, color) => ({
      data: { title, description, color },
      addFields: jest.fn().mockReturnThis()
    })),
    COLORS: {
      SUCCESS: 0x00ff00,
      WARNING: 0xffff00
    }
  }
}));

describe('MobileCommand', () => {
  let mobileCommand: MobileCommand;
  let mockConfig: BotConfig;
  
  beforeEach(() => {
    mockConfig = {
      // Add required config properties
      google: {
        credentials: {
          client_email: 'test@example.com',
          private_key: 'test-key'
        },
        folderId: 'test-folder-id'
      }
    } as unknown as BotConfig;
    
    mobileCommand = new MobileCommand(mockConfig);
    
    // Clear all mocks
    jest.clearAllMocks();
  });

  test('should be an instance of BaseCommand', () => {
    expect(mobileCommand).toBeInstanceOf(BaseCommand);
  });

  test('should have correct command name and description', () => {
    expect(mobileCommand.data.name).toBe('mobile');
    // The description is translated, so we check if it's calling the translation function
    expect(t).toHaveBeenCalledWith('mobile.command.description');
  });

  test('should handle enable subcommand', async () => {
    const mockInteraction = {
      user: { id: 'test-user-id' },
      options: {
        getSubcommand: jest.fn().mockReturnValue('enable')
      }
    };
    
    await mobileCommand['handleEnable'](mockInteraction as any);
    
    expect(mobileOptimizationService.setEnabled).toHaveBeenCalledWith('test-user-id', true);
  });

  test('should handle disable subcommand', async () => {
    const mockInteraction = {
      user: { id: 'test-user-id' },
      options: {
        getSubcommand: jest.fn().mockReturnValue('disable')
      }
    };
    
    await mobileCommand['handleDisable'](mockInteraction as any);
    
    expect(mobileOptimizationService.setEnabled).toHaveBeenCalledWith('test-user-id', false);
  });

  test('should handle status subcommand when enabled', async () => {
    const mockInteraction = {
      user: { id: 'test-user-id' },
      options: {
        getSubcommand: jest.fn().mockReturnValue('status')
      }
    };
    
    (mobileOptimizationService.getStatus as jest.Mock).mockReturnValue({
      enabled: true,
      mode: 'compact'
    });
    
    (mobileOptimizationService.getUserPreferences as jest.Mock).mockReturnValue({
      compactMode: true,
      contrastMode: false,
      maxComponentsPerRow: 3,
      maxActionRows: 3
    });
    
    await mobileCommand['handleStatus'](mockInteraction as any);
    
    expect(mobileOptimizationService.getStatus).toHaveBeenCalledWith('test-user-id');
    expect(mobileOptimizationService.getUserPreferences).toHaveBeenCalledWith('test-user-id');
  });

  test('should handle status subcommand when disabled', async () => {
    const mockInteraction = {
      user: { id: 'test-user-id' },
      options: {
        getSubcommand: jest.fn().mockReturnValue('status')
      }
    };
    
    (mobileOptimizationService.getStatus as jest.Mock).mockReturnValue({
      enabled: false,
      mode: 'disabled'
    });
    
    await mobileCommand['handleStatus'](mockInteraction as any);
    
    expect(mobileOptimizationService.getStatus).toHaveBeenCalledWith('test-user-id');
  });

  test('should handle compact subcommand', async () => {
    const mockInteraction = {
      user: { id: 'test-user-id' },
      options: {
        getSubcommand: jest.fn().mockReturnValue('compact'),
        getBoolean: jest.fn().mockReturnValue(true)
      }
    };
    
    (mobileOptimizationService.getUserPreferences as jest.Mock).mockReturnValue({
      compactMode: false,
      contrastMode: false
    });
    
    await mobileCommand['handleCompact'](mockInteraction as any);
    
    expect(mobileOptimizationService.getUserPreferences).toHaveBeenCalledWith('test-user-id');
    expect(mobileOptimizationService.setUserPreferences).toHaveBeenCalledWith('test-user-id', {
      compactMode: true,
      contrastMode: false
    });
  });

  test('should handle contrast subcommand', async () => {
    const mockInteraction = {
      user: { id: 'test-user-id' },
      options: {
        getSubcommand: jest.fn().mockReturnValue('contrast'),
        getBoolean: jest.fn().mockReturnValue(true)
      }
    };
    
    (mobileOptimizationService.getUserPreferences as jest.Mock).mockReturnValue({
      compactMode: false,
      contrastMode: false
    });
    
    await mobileCommand['handleContrast'](mockInteraction as any);
    
    expect(mobileOptimizationService.getUserPreferences).toHaveBeenCalledWith('test-user-id');
    expect(mobileOptimizationService.setUserPreferences).toHaveBeenCalledWith('test-user-id', {
      compactMode: false,
      contrastMode: true
    });
  });
});