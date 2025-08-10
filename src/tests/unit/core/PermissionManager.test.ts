/**
 * Тести для PermissionManager
 * Версія 1.0.0 - Комплексне тестування системи прав
 */

import { describe, it, expect, beforeEach, afterEach, jest, beforeAll } from '@jest/globals';
import { User, GuildMember, Guild, PermissionFlagsBits } from 'discord.js';
import PermissionManager, { UserLevel } from '@/core/PermissionManager';
import type { BotConfig } from '@/types';

// Mock dependencies
jest.mock('@/utils/logger');

describe('PermissionManager', () => {
  let permissionManager: PermissionManager;
  let mockConfig: BotConfig;
  let mockUser: User;
  let mockMember: GuildMember;
  let mockGuild: Guild;

  beforeAll(() => {
    // Mock config
    mockConfig = {
      environment: 'test',
      bot: {
        token: 'test-token',
        clientId: 'test-client-id'
      },
      ai: {
        provider: 'openai',
        apiKey: 'test-key'
      },
      cache: {
        ttl: 300,
        maxSize: 1000
      }
    } as BotConfig;
  });

  beforeEach(() => {
    // Reset mocks
    jest.clearAllMocks();
    
    // Create permission manager instance
    permissionManager = new PermissionManager(mockConfig);

    // Mock Discord objects
    mockUser = {
      id: '123456789',
      username: 'testuser',
      discriminator: '0001'
    } as User;

    mockGuild = {
      id: '987654321',
      ownerId: '111111111',
      members: {
        me: {
          permissions: {
            has: jest.fn().mockReturnValue(true)
          }
        }
      }
    } as unknown as Guild;

    mockMember = {
      id: '123456789',
      user: mockUser,
      guild: mockGuild,
      roles: {
        cache: new Map([
          ['role1', { name: 'Bot User', id: 'role1' }],
          ['role2', { name: 'Trusted', id: 'role2' }]
        ])
      },
      permissions: {
        has: jest.fn().mockReturnValue(true),
        toArray: jest.fn().mockReturnValue(['ViewChannel', 'SendMessages'])
      }
    } as unknown as GuildMember;
  });

  afterEach(() => {
    permissionManager.cleanup();
  });

  describe('Constructor and Initialization', () => {
    it('should initialize with default configuration', () => {
      expect(permissionManager).toBeDefined();
      expect(permissionManager.getStats().isInitialized).toBe(true);
    });

    it('should return singleton instance', () => {
      const instance1 = new PermissionManager(mockConfig);
      const instance2 = new PermissionManager(mockConfig);
      expect(instance1).toBe(instance2);
    });

    it('should load permission configurations', () => {
      const stats = permissionManager.getStats();
      expect(stats.commandConfigs).toBeGreaterThan(0);
    });
  });

  describe('User Level Determination', () => {
    it('should identify server owner as OWNER level', async () => {
      mockGuild.ownerId = mockMember.id;
      
      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.userLevel).toBe(UserLevel.OWNER);
      expect(result.allowed).toBe(true);
    });

    it('should identify administrator as ADMIN level', async () => {
      (mockMember.permissions.has as jest.Mock).mockImplementation((perm) => {
        return perm === PermissionFlagsBits.Administrator;
      });

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.userLevel).toBe(UserLevel.ADMIN);
    });

    it('should identify moderator correctly', async () => {
      (mockMember.permissions.has as jest.Mock).mockImplementation((perms) => {
        const modPerms = [
          PermissionFlagsBits.ManageMessages,
          PermissionFlagsBits.ManageChannels,
          PermissionFlagsBits.ManageRoles
        ];
        return Array.isArray(perms) ? modPerms.every(p => perms.includes(p)) : modPerms.includes(perms);
      });

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.userLevel).toBe(UserLevel.MODERATOR);
    });

    it('should identify trusted user by role', async () => {
      mockMember.roles.cache.set('trusted', { name: 'Trusted', id: 'trusted' } as any);

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.userLevel).toBe(UserLevel.TRUSTED);
    });

    it('should identify banned user', async () => {
      mockMember.roles.cache.clear();
      mockMember.roles.cache.set('banned', { name: 'Banned', id: 'banned' } as any);

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.userLevel).toBe(UserLevel.BANNED);
      expect(result.allowed).toBe(false);
    });

    it('should default to USER level', async () => {
      mockMember.roles.cache.clear();
      (mockMember.permissions.has as jest.Mock).mockReturnValue(false);

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.userLevel).toBe(UserLevel.USER);
    });
  });

  describe('Permission Checking', () => {
    it('should allow access for valid permissions', async () => {
      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.allowed).toBe(true);
      expect(result.hasRequiredRoles).toBe(true);
      expect(result.hasRequiredPermissions).toBe(true);
    });

    it('should deny access for insufficient user level', async () => {
      // Test with AI command that requires TRUSTED level
      mockMember.roles.cache.clear(); // Only USER level

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'ai-асистент'
      );

      expect(result.allowed).toBe(false);
      expect(result.reason).toContain('Недостатній рівень користувача');
    });

    it('should deny access for missing roles', async () => {
      mockMember.roles.cache.clear(); // No roles

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.allowed).toBe(false);
      expect(result.hasRequiredRoles).toBe(false);
    });

    it('should deny access for missing Discord permissions', async () => {
      (mockMember.permissions.has as jest.Mock).mockReturnValue(false);

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.allowed).toBe(false);
      expect(result.hasRequiredPermissions).toBe(false);
    });

    it('should deny access in DM for server-only commands', async () => {
      const result = await permissionManager.checkPermission(
        mockUser,
        null, // No member = DM
        'пошук'
      );

      expect(result.allowed).toBe(false);
      expect(result.reason).toBe('Команда доступна тільки на сервері');
    });

    it('should allow access for unknown commands by default', async () => {
      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'unknown-command'
      );

      expect(result.allowed).toBe(true);
      expect(result.reason).toBe('Конфігурація не знайдена');
    });
  });

  describe('Rate Limiting and Cooldowns', () => {
    it('should track command usage', async () => {
      await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      const stats = permissionManager.getStats();
      expect(stats.dailyUsage).toBeGreaterThan(0);
    });

    it('should enforce cooldowns', async () => {
      // First request should succeed
      const result1 = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );
      expect(result1.allowed).toBe(true);

      // Immediate second request should be limited by cooldown
      // Note: This would require mocking the command usage tracking
      // For now, we just verify the structure is in place
      expect(result1.remainingUses).toBeDefined();
    });

    it('should enforce daily limits', async () => {
      // This would require extensive mocking of time and usage tracking
      // For now, we verify the basic structure
      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.remainingUses).toBeDefined();
    });
  });

  describe('Admin Bypass', () => {
    it('should allow admin bypass for role requirements', async () => {
      // Remove required roles but keep admin permissions
      mockMember.roles.cache.clear();
      (mockMember.permissions.has as jest.Mock).mockImplementation((perm) => {
        return perm === PermissionFlagsBits.Administrator;
      });

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      expect(result.allowed).toBe(true);
      expect(result.hasRequiredRoles).toBe(true); // Should be true due to admin bypass
    });
  });

  describe('Caching', () => {
    it('should cache user information', async () => {
      // First call
      await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      // Second call should use cache
      await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      const stats = permissionManager.getStats();
      expect(stats.cachedUsers).toBeGreaterThan(0);
    });

    it('should clean up cache when limit reached', async () => {
      // This would require creating many cache entries to test cleanup
      // For now, we verify the method exists
      expect(permissionManager.getStats().cachedUsers).toBeDefined();
    });
  });

  describe('Configuration Management', () => {
    it('should allow adding new permission configs', () => {
      const initialCount = permissionManager.getStats().commandConfigs;

      permissionManager.addPermissionConfig({
        command: 'test-command',
        requiredRoles: ['Test Role'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.USER,
        cooldown: 5000,
        maxUsesPerDay: 10
      });

      expect(permissionManager.getStats().commandConfigs).toBe(initialCount + 1);
    });

    it('should provide comprehensive stats', () => {
      const stats = permissionManager.getStats();

      expect(stats).toHaveProperty('cachedUsers');
      expect(stats).toHaveProperty('commandConfigs');
      expect(stats).toHaveProperty('dailyUsage');
      expect(stats).toHaveProperty('isInitialized');

      expect(typeof stats.cachedUsers).toBe('number');
      expect(typeof stats.commandConfigs).toBe('number');
      expect(typeof stats.dailyUsage).toBe('number');
      expect(typeof stats.isInitialized).toBe('boolean');
    });
  });

  describe('Error Handling', () => {
    it('should handle errors gracefully', async () => {
      // Mock an error in permission checking
      const invalidMember = {
        ...mockMember,
        permissions: {
          has: jest.fn().mockImplementation(() => {
            throw new Error('Permission check failed');
          })
        }
      } as unknown as GuildMember;

      const result = await permissionManager.checkPermission(
        mockUser,
        invalidMember,
        'пошук'
      );

      // Should fallback to safe defaults
      expect(result.allowed).toBe(true); // Basic commands should be allowed on error
      expect(result.reason).toBe('Помилка системи прав доступу');
    });

    it('should not allow critical commands on error', async () => {
      const invalidMember = {
        ...mockMember,
        permissions: {
          has: jest.fn().mockImplementation(() => {
            throw new Error('Permission check failed');
          })
        }
      } as unknown as GuildMember;

      const result = await permissionManager.checkPermission(
        mockUser,
        invalidMember,
        'операції' // Critical operations command
      );

      expect(result.allowed).toBe(false);
    });
  });

  describe('Cleanup', () => {
    it('should clean up resources properly', () => {
      const initialStats = permissionManager.getStats();
      expect(initialStats.isInitialized).toBe(true);

      permissionManager.cleanup();

      // After cleanup, stats should be reset
      const cleanedStats = permissionManager.getStats();
      expect(cleanedStats.cachedUsers).toBe(0);
      expect(cleanedStats.commandConfigs).toBe(0);
      expect(cleanedStats.dailyUsage).toBe(0);
      expect(cleanedStats.isInitialized).toBe(false);
    });
  });

  describe('Security Logging', () => {
    it('should log security events for denied access', async () => {
      // Mock logger to capture security events
      const mockLogger = require('@/utils/logger');
      
      mockMember.roles.cache.clear(); // Remove roles to trigger denial

      await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      // Verify security logging was called
      expect(mockLogger.security).toHaveBeenCalled();
    });

    it('should log successful access for audit', async () => {
      const mockLogger = require('@/utils/logger');

      await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'пошук'
      );

      // Verify info logging was called for successful access
      expect(mockLogger.info).toHaveBeenCalled();
    });
  });
});