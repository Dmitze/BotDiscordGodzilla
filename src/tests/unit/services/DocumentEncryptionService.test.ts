/**
 * Unit tests for DocumentEncryptionService
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { DocumentEncryptionService } from '../../../services/DocumentEncryptionService';
import { createMockConfig } from '../../utils/testHelpers';

describe('DocumentEncryptionService', () => {
  let encryptionService: DocumentEncryptionService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    encryptionService = new DocumentEncryptionService(mockConfig);
  });

  describe('encryptDocumentContent', () => {
    it('should encrypt document content successfully', () => {
      const content = 'This is sensitive document content';
      
      const encrypted = encryptionService.encryptDocumentContent(content);
      
      expect(encrypted).toBeDefined();
      expect(encrypted.encryptedData).toBeDefined();
      expect(encrypted.iv).toBeDefined();
      expect(encrypted.authTag).toBeDefined();
      expect(encrypted.salt).toBeDefined();
      expect(encrypted.algorithm).toBe('aes-256-gcm');
      expect(encrypted.createdAt).toBeDefined();
    });

    it('should encrypt content with password', () => {
      const content = 'This is sensitive document content';
      const password = 'strongPassword123';
      
      const encrypted = encryptionService.encryptDocumentContent(content, password);
      
      expect(encrypted).toBeDefined();
      expect(encrypted.encryptedData).toBeDefined();
    });

    it('should throw error for invalid content', () => {
      expect(() => {
        encryptionService.encryptDocumentContent(null as any);
      }).toThrow();
    });
  });

  describe('decryptDocumentContent', () => {
    it('should decrypt document content successfully', () => {
      const originalContent = 'This is sensitive document content';
      
      const encrypted = encryptionService.encryptDocumentContent(originalContent);
      const decrypted = encryptionService.decryptDocumentContent(encrypted);
      
      expect(decrypted).toBe(originalContent);
    });

    it('should decrypt content with password', () => {
      const originalContent = 'This is sensitive document content';
      const password = 'strongPassword123';
      
      const encrypted = encryptionService.encryptDocumentContent(originalContent, password);
      const decrypted = encryptionService.decryptDocumentContent(encrypted, password);
      
      expect(decrypted).toBe(originalContent);
    });

    it('should throw error for invalid encrypted document', () => {
      const invalidEncryptedDoc = {
        encryptedData: 'invalid',
        iv: 'invalid',
        authTag: 'invalid',
        salt: 'invalid',
        algorithm: 'aes-256-gcm',
        createdAt: new Date(),
      };
      
      expect(() => {
        encryptionService.decryptDocumentContent(invalidEncryptedDoc);
      }).toThrow();
    });

    it('should throw error for corrupted encrypted document', () => {
      const encrypted = encryptionService.encryptDocumentContent('test content');
      
      // Corrupt the encrypted data
      encrypted.encryptedData = 'corrupted';
      
      expect(() => {
        encryptionService.decryptDocumentContent(encrypted);
      }).toThrow();
    });
  });

  describe('shouldEncryptDocument', () => {
    it('should return true for sensitive content', () => {
      const content = 'This document contains confidential information';
      const fileName = 'regular-document.txt';
      
      const result = encryptionService.shouldEncryptDocument(content, fileName);
      
      expect(result).toBe(true);
    });

    it('should return true for sensitive filename', () => {
      const content = 'Regular content';
      const fileName = 'confidential-report.txt';
      
      const result = encryptionService.shouldEncryptDocument(content, fileName);
      
      expect(result).toBe(true);
    });

    it('should return false for non-sensitive content and filename', () => {
      const content = 'Regular content';
      const fileName = 'regular-document.txt';
      
      const result = encryptionService.shouldEncryptDocument(content, fileName);
      
      expect(result).toBe(false);
    });
  });

  describe('encryptIfSensitive', () => {
    it('should encrypt sensitive document', () => {
      const content = 'This document contains confidential information';
      const fileName = 'regular-document.txt';
      
      const result = encryptionService.encryptIfSensitive(content, fileName);
      
      expect(result).toBeDefined();
      expect(result?.encryptedData).toBeDefined();
    });

    it('should not encrypt non-sensitive document', () => {
      const content = 'Regular content';
      const fileName = 'regular-document.txt';
      
      const result = encryptionService.encryptIfSensitive(content, fileName);
      
      expect(result).toBeNull();
    });
  });

  describe('getStats', () => {
    it('should return correct statistics after operations', () => {
      const content = 'Test content';
      
      // Perform some operations
      const encrypted = encryptionService.encryptDocumentContent(content);
      encryptionService.decryptDocumentContent(encrypted);
      
      const stats = encryptionService.getStats();
      
      expect(stats.totalDocumentsEncrypted).toBe(1);
      expect(stats.totalDocumentsDecrypted).toBe(1);
      expect(stats.failedEncryptions).toBe(0);
      expect(stats.failedDecryptions).toBe(0);
      expect(stats.totalEncryptionTime).toBeGreaterThan(0);
      expect(stats.totalDecryptionTime).toBeGreaterThan(0);
    });
  });
});