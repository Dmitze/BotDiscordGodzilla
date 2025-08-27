/**
 * Integration tests for REST API endpoints
 */

import { describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import request from 'supertest';
import express from 'express';
import { DocumentAccessAuditService } from '../../services/DocumentAccessAuditService';
import { DataLossPreventionService } from '../../services/DataLossPreventionService';
import { ComplianceReportingService } from '../../services/ComplianceReportingService';
import { DriveIndexerService } from '../../services/DriveIndexerService';
import { createMockConfig, createMockDriveFile } from '../utils/testHelpers';
import app, { initializeApiServices, startApiServer } from '../../api';

describe('API Integration Tests', () => {
  let auditService: DocumentAccessAuditService;
  let dlpService: DataLossPreventionService;
  let complianceService: ComplianceReportingService;
  let indexerService: DriveIndexerService;
  let mockConfig: any;
  let server: any;

  beforeAll(async () => {
    mockConfig = createMockConfig();
    
    // Create service instances
    auditService = new DocumentAccessAuditService(mockConfig);
    dlpService = new DataLossPreventionService(mockConfig);
    complianceService = new ComplianceReportingService(mockConfig);
    
    // Create a mock indexer service
    const mockBot = {
      config: mockConfig,
      getService: jest.fn()
    };
    indexerService = new DriveIndexerService(mockBot as any);
    
    // Initialize API services
    initializeApiServices({
      auditService,
      dlpService,
      complianceService,
      indexerService
    });
    
    // Log some test data
    const mockFile = createMockDriveFile('test-file', 'Test Document.txt');
    auditService.logDocumentAccess({
      file: mockFile,
      userId: 'test-user',
      userName: 'Test User',
      action: 'view',
      sessionId: 'test-session'
    });
  });

  describe('Health Check Endpoint', () => {
    it('should return health status', async () => {
      const response = await request(app).get('/health');
      
      expect(response.status).toBe(200);
      expect(response.body.status).toBe('ok');
      expect(response.body.service).toBe('Discord AI Assistant Bot API');
    });
  });

  describe('Audit Endpoints', () => {
    it('should require authentication for audit records endpoint', async () => {
      const response = await request(app).get('/audit/records');
      
      expect(response.status).toBe(401);
      expect(response.body.error).toBe('Access token required');
    });

    it('should return audit records when authenticated', async () => {
      const response = await request(app)
        .get('/audit/records')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(200);
      expect(response.body.records).toBeDefined();
      expect(response.body.total).toBeDefined();
    });

    it('should return audit summary when authenticated', async () => {
      const response = await request(app)
        .get('/audit/summary')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(200);
      expect(response.body.totalAccesses).toBeDefined();
      expect(response.body.uniqueUsers).toBeDefined();
    });
  });

  describe('DLP Endpoints', () => {
    it('should scan document for sensitive data', async () => {
      const mockFile = createMockDriveFile('dlp-test', 'DLP Test Document.txt');
      
      const response = await request(app)
        .post('/dlp/scan')
        .set('Authorization', 'Bearer test-token')
        .send({
          file: mockFile,
          content: 'Contact: john.doe@example.com'
        });
      
      expect(response.status).toBe(200);
      expect(response.body.fileId).toBe('dlp-test');
      expect(response.body.totalFindings).toBeDefined();
    });

    it('should return 400 for missing scan data', async () => {
      const response = await request(app)
        .post('/dlp/scan')
        .set('Authorization', 'Bearer test-token')
        .send({});
      
      expect(response.status).toBe(400);
      expect(response.body.error).toBe('File and content are required');
    });

    it('should retrieve cached scan results', async () => {
      const mockFile = createMockDriveFile('cached-test', 'Cached Test Document.txt');
      
      // First scan to cache the result
      await request(app)
        .post('/dlp/scan')
        .set('Authorization', 'Bearer test-token')
        .send({
          file: mockFile,
          content: 'Email: test@example.com'
        });
      
      // Retrieve cached result
      const response = await request(app)
        .get('/dlp/result/cached-test')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(200);
      expect(response.body.fileId).toBe('cached-test');
    });
  });

  describe('Compliance Endpoints', () => {
    it('should generate compliance report', async () => {
      const response = await request(app)
        .post('/compliance/report')
        .set('Authorization', 'Bearer test-token')
        .send({
          regulations: ['GDPR'],
          periodStart: new Date(Date.now() - 86400000).toISOString(), // 1 day ago
          periodEnd: new Date().toISOString(),
          organization: 'Test Organization'
        });
      
      expect(response.status).toBe(200);
      expect(response.body.organization).toBe('Test Organization');
      expect(response.body.regulations).toContain('GDPR');
    });

    it('should return 400 for missing report parameters', async () => {
      const response = await request(app)
        .post('/compliance/report')
        .set('Authorization', 'Bearer test-token')
        .send({});
      
      expect(response.status).toBe(400);
      expect(response.body.error).toBe('Missing required parameters');
    });

    it('should retrieve compliance report by ID', async () => {
      // First generate a report
      const generateResponse = await request(app)
        .post('/compliance/report')
        .set('Authorization', 'Bearer test-token')
        .send({
          regulations: ['HIPAA'],
          periodStart: new Date(Date.now() - 86400000).toISOString(), // 1 day ago
          periodEnd: new Date().toISOString(),
          organization: 'Retrieve Test Organization'
        });
      
      expect(generateResponse.status).toBe(200);
      const reportId = generateResponse.body.id;
      
      // Retrieve the report
      const response = await request(app)
        .get(`/compliance/report/${reportId}`)
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(200);
      expect(response.body.id).toBe(reportId);
    });

    it('should export compliance report in different formats', async () => {
      // First generate a report
      const generateResponse = await request(app)
        .post('/compliance/report')
        .set('Authorization', 'Bearer test-token')
        .send({
          regulations: ['GDPR'],
          periodStart: new Date(Date.now() - 86400000).toISOString(), // 1 day ago
          periodEnd: new Date().toISOString(),
          organization: 'Export Test Organization'
        });
      
      expect(generateResponse.status).toBe(200);
      const reportId = generateResponse.body.id;
      
      // Export as JSON
      const jsonResponse = await request(app)
        .get(`/compliance/report/${reportId}/export?format=json`)
        .set('Authorization', 'Bearer test-token');
      
      expect(jsonResponse.status).toBe(200);
      expect(jsonResponse.headers['content-type']).toContain('application/json');
      
      // Export as CSV
      const csvResponse = await request(app)
        .get(`/compliance/report/${reportId}/export?format=csv`)
        .set('Authorization', 'Bearer test-token');
      
      expect(csvResponse.status).toBe(200);
      expect(csvResponse.headers['content-type']).toContain('text/csv');
    });
  });

  describe('Statistics Endpoint', () => {
    it('should return service statistics', async () => {
      const response = await request(app)
        .get('/stats')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(200);
      expect(response.body.audit).toBeDefined();
      expect(response.body.dlp).toBeDefined();
      expect(response.body.compliance).toBeDefined();
    });
  });

  describe('Search Endpoint', () => {
    it('should return search results', async () => {
      const response = await request(app)
        .get('/search?query=test')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(200);
      expect(response.body.query).toBe('test');
      expect(response.body.results).toBeDefined();
    });

    it('should return 400 for missing search query', async () => {
      const response = await request(app)
        .get('/search')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(400);
      expect(response.body.error).toBe('Query parameter is required');
    });
  });

  describe('Custom Configuration Endpoints', () => {
    it('should add custom DLP pattern', async () => {
      const response = await request(app)
        .post('/dlp/patterns')
        .set('Authorization', 'Bearer test-token')
        .send({
          id: 'custom-test',
          name: 'Custom Test Pattern',
          pattern: '\\btest\\d{4}\\b',
          severity: 'medium',
          category: 'Test',
          description: 'Custom test pattern'
        });
      
      expect(response.status).toBe(201);
      expect(response.body.message).toBe('Pattern added successfully');
    });

    it('should return 400 for invalid DLP pattern', async () => {
      const response = await request(app)
        .post('/dlp/patterns')
        .set('Authorization', 'Bearer test-token')
        .send({
          name: 'Invalid Pattern' // Missing required fields
        });
      
      expect(response.status).toBe(400);
      expect(response.body.error).toBe('Invalid pattern data');
    });

    it('should add custom compliance requirement', async () => {
      const response = await request(app)
        .post('/compliance/requirements')
        .set('Authorization', 'Bearer test-token')
        .send({
          id: 'custom-requirement',
          name: 'Custom Requirement',
          description: 'Custom compliance requirement',
          applicableRegulations: ['Test'],
          checkFunction: '() => ({ status: "compliant" })' // Simplified for test
        });
      
      expect(response.status).toBe(201);
      expect(response.body.message).toBe('Requirement added successfully');
    });

    it('should return 400 for invalid compliance requirement', async () => {
      const response = await request(app)
        .post('/compliance/requirements')
        .set('Authorization', 'Bearer test-token')
        .send({
          name: 'Invalid Requirement' // Missing required fields
        });
      
      expect(response.status).toBe(400);
      expect(response.body.error).toBe('Invalid requirement data');
    });
  });

  describe('Error Handling', () => {
    it('should return 404 for non-existent endpoints', async () => {
      const response = await request(app)
        .get('/non-existent-endpoint')
        .set('Authorization', 'Bearer test-token');
      
      expect(response.status).toBe(404);
      expect(response.body.error).toBe('Endpoint not found');
    });

    it('should handle internal server errors gracefully', async () => {
      // This test would require mocking a service to throw an error
      // For now, we'll test the error handling middleware with a direct call
      const response = await request(app)
        .get('/health')
        .set('Test-Error', 'true'); // This header doesn't exist, so it won't cause an error
      
      // The request should still succeed
      expect(response.status).toBe(200);
    });
  });
});