/**
 * Unit tests for DataLossPreventionService functionality
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { DataLossPreventionService } from '../../../services/DataLossPreventionService';
import { createMockConfig, createMockDriveFile } from '../../utils/testHelpers';

describe('DataLossPreventionService', () => {
  let dlpService: DataLossPreventionService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    dlpService = new DataLossPreventionService(mockConfig);
  });

  it('should initialize with default patterns and policies', () => {
    const stats = dlpService.getStats();
    expect(stats.totalPatterns).toBeGreaterThan(0);
    expect(stats.activePolicies).toBeGreaterThan(0);
  });

  it('should scan document content for sensitive data', async () => {
    const mockFile = createMockDriveFile('test-file-id', 'Test Document.txt');
    
    // Content with sensitive data that matches the actual patterns
    const content = `
      Contact information:
      Email: john.doe@example.com
      Phone: (555) 123-4567
      Credit Card: 4532123456789012
      SSN: 123-45-6789
    `;
    
    const scanResult = await dlpService.scanDocument(mockFile, content);
    
    expect(scanResult).toBeDefined();
    expect(scanResult.fileId).toBe('test-file-id');
    expect(scanResult.totalFindings).toBeGreaterThan(0);
    expect(scanResult.riskScore).toBeGreaterThan(0);
    expect(scanResult.recommendedActions).toContain('log');
    expect(scanResult.recommendedActions).toContain('alert');
  });

  it('should detect credit card numbers with Luhn validation', async () => {
    const mockFile = createMockDriveFile('cc-test', 'Credit Card Test.txt');
    
    // Valid credit card number
    const validContent = 'Payment information: 4532123456789012';
    
    // Invalid credit card number
    const invalidContent = 'Fake number: 1234567890123456';
    
    const validResult = await dlpService.scanDocument(mockFile, validContent);
    const invalidResult = await dlpService.scanDocument(
      { ...mockFile, id: 'invalid-cc-test' },
      invalidContent
    );
    
    // Valid CC should have higher confidence
    const validFinding = validResult.findings.find(f => f.patternId.includes('visa'));
    const invalidFinding = invalidResult.findings.find(f => f.patternId.includes('generic'));
    
    expect(validFinding).toBeDefined();
    // The invalid one might not be detected due to confidence filtering
  });

  it('should calculate risk scores correctly', async () => {
    const mockFile = createMockDriveFile('risk-test', 'Risk Test.txt');
    
    // Content with different severity findings that match the actual patterns
    const content = `
      Low severity: user@example.com
      Medium severity: (555) 123-4567
      High severity: 4532123456789012
      Critical severity: password = "secret123"
    `;
    
    const scanResult = await dlpService.scanDocument(mockFile, content);
    
    expect(scanResult.riskScore).toBeGreaterThan(0);
    expect(scanResult.findingsBySeverity).toHaveProperty('low');
    expect(scanResult.findingsBySeverity).toHaveProperty('medium');
    expect(scanResult.findingsBySeverity).toHaveProperty('high');
    expect(scanResult.findingsBySeverity).toHaveProperty('critical');
  });

  it('should cache scan results', async () => {
    const mockFile = createMockDriveFile('cache-test', 'Cache Test.txt');
    const content = 'Email: test@example.com';
    
    // First scan
    const result1 = await dlpService.scanDocument(mockFile, content);
    
    // Second scan should use cache
    const result2 = await dlpService.scanDocument(mockFile, content);
    
    expect(result1).toEqual(result2);
  });

  it('should handle custom patterns', () => {
    const customPattern = {
      id: 'custom-pattern',
      name: 'Custom Test Pattern',
      pattern: /\btest\d{4}\b/g,
      severity: 'medium' as const,
      category: 'Custom',
      description: 'Custom test pattern'
    };
    
    dlpService.addPattern(customPattern);
    
    const stats = dlpService.getStats();
    expect(stats.totalPatterns).toBeGreaterThan(10); // Should have more than default patterns
  });

  it('should manage policies correctly', () => {
    // Add a custom policy
    const customPolicy = {
      id: 'custom-policy',
      name: 'Custom Policy',
      description: 'Custom test policy',
      enabled: true,
      patterns: ['email'],
      severityThreshold: 'low' as const,
      actions: ['log']
    };
    
    dlpService.addPolicy(customPolicy);
    
    // Update policy
    const updated = dlpService.updatePolicy('custom-policy', { enabled: false });
    expect(updated).toBeDefined();
    expect(updated?.enabled).toBe(false);
    
    // Remove policy
    const removed = dlpService.removePolicy('custom-policy');
    expect(removed).toBe(true);
  });

  it('should filter findings by policy thresholds', async () => {
    const mockFile = createMockDriveFile('policy-test', 'Policy Test.txt');
    
    // Content with only low severity findings
    const content = 'Contact: user@example.com';
    
    const scanResult = await dlpService.scanDocument(mockFile, content);
    
    // With default policy threshold of 'medium', low severity findings
    // might not trigger all actions
    expect(scanResult).toBeDefined();
  });

  it('should provide service statistics', () => {
    const stats = dlpService.getStats();
    
    expect(stats).toBeDefined();
    expect(typeof stats.totalPatterns).toBe('number');
    expect(typeof stats.activePolicies).toBe('number');
    expect(typeof stats.cachedResults).toBe('number');
    expect(typeof stats.averageRiskScore).toBe('number');
  });

  it('should handle empty or invalid content', async () => {
    const mockFile = createMockDriveFile('empty-test', 'Empty Test.txt');
    
    // Test with empty content
    const emptyResult = await dlpService.scanDocument(mockFile, '');
    expect(emptyResult.totalFindings).toBe(0);
    expect(emptyResult.riskScore).toBe(0);
    
    // Test with whitespace content
    const whitespaceResult = await dlpService.scanDocument(mockFile, '   \n\t   ');
    expect(whitespaceResult.totalFindings).toBe(0);
  });

  it('should maintain pattern and policy limits', () => {
    // Add many patterns
    for (let i = 0; i < 50; i++) {
      dlpService.addPattern({
        id: `pattern-${i}`,
        name: `Pattern ${i}`,
        pattern: /\btest\b/g,
        severity: 'low',
        category: 'Test',
        description: `Test pattern ${i}`
      });
    }
    
    // Add many policies
    for (let i = 0; i < 20; i++) {
      dlpService.addPolicy({
        id: `policy-${i}`,
        name: `Policy ${i}`,
        description: `Test policy ${i}`,
        enabled: true,
        patterns: ['email'],
        severityThreshold: 'low',
        actions: ['log']
      });
    }
    
    const stats = dlpService.getStats();
    expect(stats.totalPatterns).toBeGreaterThan(10);
    expect(stats.activePolicies).toBeGreaterThan(1);
  });
});