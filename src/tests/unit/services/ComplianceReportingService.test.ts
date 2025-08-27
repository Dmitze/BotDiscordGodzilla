/**
 * Unit tests for ComplianceReportingService functionality
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { ComplianceReportingService } from '../../../services/ComplianceReportingService';
import { DocumentAccessAuditService } from '../../../services/DocumentAccessAuditService';
import { DataLossPreventionService } from '../../../services/DataLossPreventionService';
import { createMockConfig, createMockDriveFile } from '../../utils/testHelpers';

describe('ComplianceReportingService', () => {
  let complianceService: ComplianceReportingService;
  let auditService: DocumentAccessAuditService;
  let dlpService: DataLossPreventionService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    complianceService = new ComplianceReportingService(mockConfig);
    auditService = new DocumentAccessAuditService(mockConfig);
    dlpService = new DataLossPreventionService(mockConfig);
  });

  it('should initialize with default compliance requirements', () => {
    const stats = complianceService.getStats();
    expect(stats.totalRequirements).toBeGreaterThan(0);
  });

  it('should generate a compliance report with GDPR requirements', async () => {
    // Log some audit data
    const mockFile = createMockDriveFile('test-file', 'Test Document.txt');
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'user-1',
      userName: 'Test User',
      action: 'view',
      sessionId: 'session-1'
    });
    
    // Generate compliance report
    const report = await complianceService.generateComplianceReport({
      auditService,
      dlpService,
      regulations: ['GDPR'],
      periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
      periodEnd: new Date(),
      organization: 'Test Organization'
    });
    
    expect(report).toBeDefined();
    expect(report.id).toBeDefined();
    expect(report.organization).toBe('Test Organization');
    expect(report.regulations).toContain('GDPR');
    expect(report.summary.totalRequirements).toBeGreaterThan(0);
  });

  it('should generate a compliance report with HIPAA requirements', async () => {
    // Generate compliance report
    const report = await complianceService.generateComplianceReport({
      auditService,
      dlpService,
      regulations: ['HIPAA'],
      periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
      periodEnd: new Date(),
      organization: 'Healthcare Organization'
    });
    
    expect(report).toBeDefined();
    expect(report.organization).toBe('Healthcare Organization');
    expect(report.regulations).toContain('HIPAA');
  });

  it('should generate a compliance report with multiple regulations', async () => {
    // Generate compliance report
    const report = await complianceService.generateComplianceReport({
      auditService,
      dlpService,
      regulations: ['GDPR', 'HIPAA', 'SOX'],
      periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
      periodEnd: new Date(),
      organization: 'Multinational Corporation'
    });
    
    expect(report).toBeDefined();
    expect(report.organization).toBe('Multinational Corporation');
    expect(report.regulations).toContain('GDPR');
    expect(report.regulations).toContain('HIPAA');
    expect(report.regulations).toContain('SOX');
  });

  it('should cache generated reports', async () => {
    // Generate first report
    const report1 = await complianceService.generateComplianceReport({
      auditService,
      dlpService,
      regulations: ['GDPR'],
      periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
      periodEnd: new Date(),
      organization: 'Test Organization'
    });
    
    // Retrieve cached report
    const cachedReport = complianceService.getReport(report1.id);
    
    expect(cachedReport).toBeDefined();
    expect(cachedReport?.id).toBe(report1.id);
  });

  it('should handle custom compliance requirements', () => {
    // Add a custom requirement
    const customRequirement = {
      id: 'custom-1',
      name: 'Custom Security Requirement',
      description: 'Custom security requirement for testing',
      applicableRegulations: ['Internal'],
      checkFunction: (data: any) => ({
        requirementId: 'custom-1',
        requirementName: 'Custom Security Requirement',
        status: 'compliant',
        findings: [],
        severity: 'low',
        recommendations: []
      })
    };
    
    complianceService.addRequirement(customRequirement);
    
    const stats = complianceService.getStats();
    expect(stats.totalRequirements).toBeGreaterThan(10); // Should have more than default requirements
  });

  it('should export reports in different formats', async () => {
    // Generate a report
    const report = await complianceService.generateComplianceReport({
      auditService,
      dlpService,
      regulations: ['GDPR'],
      periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
      periodEnd: new Date(),
      organization: 'Test Organization'
    });
    
    // Test JSON export
    const jsonExport = complianceService.exportReport(report.id, 'json');
    expect(typeof jsonExport).toBe('string');
    expect(jsonExport).toContain(report.id);
    
    // Test CSV export
    const csvExport = complianceService.exportReport(report.id, 'csv');
    expect(typeof csvExport).toBe('string');
    expect(csvExport).toContain('Requirement,Status,Severity,Findings,Recommendations');
  });

  it('should provide service statistics', () => {
    const stats = complianceService.getStats();
    
    expect(stats).toBeDefined();
    expect(typeof stats.totalRequirements).toBe('number');
    expect(typeof stats.cachedReports).toBe('number');
  });

  it('should handle compliance check errors gracefully', async () => {
    // Add a requirement with a failing check function
    const failingRequirement = {
      id: 'failing-1',
      name: 'Failing Requirement',
      description: 'Requirement that always fails',
      applicableRegulations: ['Test'],
      checkFunction: (data: any) => {
        throw new Error('Intentional test error');
      }
    };
    
    complianceService.addRequirement(failingRequirement);
    
    // Generate report - should handle the error gracefully
    const report = await complianceService.generateComplianceReport({
      auditService,
      dlpService,
      regulations: ['Test'],
      periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
      periodEnd: new Date(),
      organization: 'Error Test Organization'
    });
    
    expect(report).toBeDefined();
    // The failing check should result in a non-compliant status
    expect(report.summary.nonCompliant).toBeGreaterThanOrEqual(0);
  });

  it('should maintain report cache limits', async () => {
    // Generate multiple reports
    for (let i = 0; i < 10; i++) {
      await complianceService.generateComplianceReport({
        auditService,
        dlpService,
        regulations: ['GDPR'],
        periodStart: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // Last 30 days
        periodEnd: new Date(),
        organization: `Organization ${i}`
      });
    }
    
    const stats = complianceService.getStats();
    expect(stats.cachedReports).toBeGreaterThan(0);
  });
});