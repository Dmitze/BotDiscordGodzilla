/**
 * Enhanced tests for PermissionManager with detailed permissions
 * Testing the new resource-based access control functionality
 */

import { describe, it, expect, beforeEach, afterEach, jest, beforeAll } from '@jest/globals';
import { User, GuildMember, Guild, PermissionFlagsBits } from 'discord.js';
import PermissionManager, { UserLevel } from '@/core/PermissionManager';
import type { BotConfig } from '@/types';

// Mock dependencies
jest.mock('@/utils/logger');

describe('PermissionManager - Enhanced Features', () => {
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
            has: jest.fn().mockImplementation((perm: unknown) => {
              // Bot can generally view/send, but not Administrator by default
              const allow = [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages];
              return Array.isArray(perm)
                ? perm.every(p => allow.includes(p as any))
                : allow.includes(perm as any);
            })
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
        has: jest.fn().mockImplementation((perm: unknown) => {
          // Fix: Include AttachFiles permission which is required for 'файли' command
          const allow = [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages, PermissionFlagsBits.AttachFiles];
          return Array.isArray(perm)
            ? perm.every(p => allow.includes(p as any))
            : allow.includes(perm as any);
        }),
        toArray: jest.fn().mockReturnValue(['ViewChannel', 'SendMessages', 'AttachFiles'])
      }
    } as unknown as GuildMember;
  });

  afterEach(() => {
    permissionManager.cleanup();
  });

  describe('Resource-Based Permissions', () => {
    it('should allow setting and getting document access control', () => {
      const documentId = 'doc-123';
      const accessControl = {
        documentId,
        allowedUsers: ['123456789'],
        allowedRoles: ['Bot User'],
        permissions: ['view', 'edit'] as const,
      };

      permissionManager.setDocumentAccessControl(documentId, accessControl);
      
      const retrieved = permissionManager.getDocumentAccessControl(documentId);
      expect(retrieved).toEqual(accessControl);
    });

    it('should allow removing document access control', () => {
      const documentId = 'doc-123';
      const accessControl = {
        documentId,
        allowedUsers: ['123456789'],
        allowedRoles: ['Bot User'],
        permissions: ['view', 'edit'] as const,
      };

      permissionManager.setDocumentAccessControl(documentId, accessControl);
      expect(permissionManager.getDocumentAccessControl(documentId)).toEqual(accessControl);

      const result = permissionManager.removeDocumentAccessControl(documentId);
      expect(result).toBe(true);
      expect(permissionManager.getDocumentAccessControl(documentId)).toBeUndefined();
    });

    it('should check document access permissions correctly', async () => {
      const documentId = 'doc-123';
      const accessControl = {
        documentId,
        allowedUsers: ['123456789'],
        allowedRoles: ['Bot User'],
        permissions: ['view', 'edit'] as const,
      };

      permissionManager.setDocumentAccessControl(documentId, accessControl);

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'файли',
        undefined,
        [{ resourceId: documentId, action: 'view' }]
      );

      expect(result.allowed).toBe(true);
      expect(result.resourceAccess).toBeDefined();
      expect(result.resourceAccess?.length).toBe(1);
      expect(result.resourceAccess?.[0].allowed).toBe(true);
    });

    it('should deny access when user is not authorized for document', async () => {
      const documentId = 'doc-123';
      const accessControl = {
        documentId,
        allowedUsers: ['999999999'], // Different user ID
        allowedRoles: ['Admin'],
        permissions: ['view', 'edit'] as const,
      };

      permissionManager.setDocumentAccessControl(documentId, accessControl);

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'файли',
        undefined,
        [{ resourceId: documentId, action: 'view' }]
      );

      expect(result.allowed).toBe(false);
      expect(result.resourceAccess).toBeDefined();
      expect(result.resourceAccess?.length).toBe(1);
      expect(result.resourceAccess?.[0].allowed).toBe(false);
    });

    it('should allow admin users to access documents without specific permissions', async () => {
      // Make the user an admin
      (mockMember.permissions.has as jest.Mock).mockImplementation((perm) => {
        return perm === PermissionFlagsBits.Administrator;
      });

      const documentId = 'doc-123';
      // No access control set for this document

      const result = await permissionManager.checkPermission(
        mockUser,
        mockMember,
        'файли',
        undefined,
        [{ resourceId: documentId, action: 'view' }]
      );

      expect(result.allowed).toBe(true);
      expect(result.userLevel).toBe(UserLevel.ADMIN);
    });
  });

  describe('Enhanced Permission Stats', () => {
    it('should include document access controls in stats', () => {
      const documentId = 'doc-123';
      const accessControl = {
        documentId,
        allowedUsers: ['123456789'],
        allowedRoles: ['Bot User'],
        permissions: ['view', 'edit'] as const,
      };

      permissionManager.setDocumentAccessControl(documentId, accessControl);
      
      const stats = permissionManager.getStats();
      expect(stats.documentAccessControls).toBe(1);
    });
  });

  describe('Detailed Permission Configurations', () => {
    it('should support resource permissions in command configs', () => {
      const customConfig = {
        command: 'custom-command',
        requiredRoles: ['Bot User'],
        requiredPermissions: [PermissionFlagsBits.ViewChannel],
        minUserLevel: UserLevel.USER,
        resourcePermissions: [
          {
            resourceId: 'custom-resource',
            resourceType: 'feature' as const,
            actions: ['read', 'write'] as const,
          }
        ]
      };

      permissionManager.addPermissionConfig(customConfig);
      
      const stats = permissionManager.getStats();
      expect(stats.commandConfigs).toBeGreaterThan(6); // Original configs + our custom one
    });
  });
});