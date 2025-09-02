import express = require('express');
import { Request, Response, NextFunction } from 'express';
import cors = require('cors');
import helmet = require('helmet');
import rateLimit = require('express-rate-limit');
import type { BotConfig } from '@/types/index';
import logger from '@/utils/logger';
import { DocumentAccessAuditService } from '@/services/DocumentAccessAuditService';
import { DataLossPreventionService } from '@/services/DataLossPreventionService';
import { ComplianceReportingService } from '@/services/ComplianceReportingService';
import { DriveIndexerService } from '@/services/DriveIndexerService';

// Define the API service interface
interface ApiService {
  auditService: DocumentAccessAuditService;
  dlpService: DataLossPreventionService;
  complianceService: ComplianceReportingService;
  indexerService: DriveIndexerService;
}

// Create Express app
const app = express();

// Middleware
app.use(helmet()); // Security headers
app.use(cors()); // Enable CORS
app.use(express.json({ limit: '10mb' })); // Parse JSON bodies
app.use(express.urlencoded({ extended: true, limit: '10mb' })); // Parse URL-encoded bodies

// Rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // limit each IP to 100 requests per windowMs
  message: 'Too many requests, please try again later.'
});
app.use(limiter);

// API services container
let apiServices: ApiService | null = null;

// Initialize API services
export function initializeApiServices(services: ApiService): void {
  apiServices = services;
  logger.info('API services initialized', { component: 'ApiService' });
}

// Authentication middleware
function authenticateToken(req: Request, res: Response, next: NextFunction): void {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1]; // Bearer TOKEN
  
  if (!token) {
    res.status(401).json({ error: 'Access token required' });
    return;
  }
  
  // In a real implementation, you would verify the token
  // For now, we'll just check if it exists
  if (token !== (process.env['API_ACCESS_TOKEN'] || '')) {
    res.status(403).json({ error: 'Invalid access token' });
    return;
  }
  
  next();
}

// Error handling middleware
function errorHandler(err: Error, req: Request, res: Response, next: NextFunction): void {
  logger.error('API error', {
    component: 'ApiService',
    error: err.message,
    stack: err.stack,
    url: req.url,
    method: req.method
  });
  
  res.status(500).json({
    error: 'Internal server error',
    message: (process.env['NODE_ENV'] || '') === 'development' ? err.message : 'An error occurred'
  });
}

// Routes

// Health check endpoint
app.get('/health', (_req: Request, res: Response) => {
  res.status(200).json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    service: 'Discord AI Assistant Bot API'
  });
});

// Enhanced health check endpoint
app.get('/health/detailed', async (_req: Request, res: Response) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }

    // This would be replaced with actual service health checks
    const health = {
      status: 'ok',
      timestamp: new Date().toISOString(),
      services: {
        audit: await apiServices.auditService.onHealthCheck?.() || { healthy: true, service: 'audit' },
        dlp: await apiServices.dlpService.onHealthCheck?.() || { healthy: true, service: 'dlp' },
        compliance: await apiServices.complianceService.onHealthCheck?.() || { healthy: true, service: 'compliance' },
        indexer: await (apiServices.indexerService as any).onHealthCheck?.() || { healthy: true, service: 'indexer' }
      }
    };

    res.json(health);
  } catch (error) {
    res.status(500).json({
      status: 'error',
      error: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

// AI Service health check
app.get('/health/ai', async (_req: Request, res: Response) => {
  try {
    // This would integrate with the actual AI service
    res.json({
      healthy: true,
      service: 'ai',
      message: 'AI service health check endpoint'
    });
  } catch (error) {
    res.status(500).json({
      healthy: false,
      service: 'ai',
      error: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

// Google Service health check
app.get('/health/google', async (_req: Request, res: Response) => {
  try {
    // This would integrate with the actual Google service
    res.json({
      healthy: true,
      service: 'google',
      message: 'Google service health check endpoint'
    });
  } catch (error) {
    res.status(500).json({
      healthy: false,
      service: 'google',
      error: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

// Get audit records
app.get('/audit/records', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { page, limit, userId, fileId, action } = req.query;
    
    const records = apiServices.auditService.getAuditRecords({
      page: page ? parseInt(page as string) : 1,
      limit: limit ? parseInt(limit as string) : 10,
      userId: userId as string || '',
      fileId: fileId as string || '',
      action: action as any
    });
    
    res.json(records);
  } catch (error) {
    next(error);
  }
});

// Get audit summary
app.get('/audit/summary', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { startDate, endDate, userId, fileId } = req.query;
    
    const summary = apiServices.auditService.generateAccessSummary({
      startDate: startDate ? new Date(startDate as string) : new Date(0),
      endDate: endDate ? new Date(endDate as string) : new Date(),
      userId: userId as string || '',
      fileId: fileId as string || ''
    });
    
    res.json(summary);
  } catch (error) {
    next(error);
  }
});

// Scan document for sensitive data
app.post('/dlp/scan', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { file, content } = req.body;
    
    if (!file || !content) {
      res.status(400).json({ error: 'File and content are required' });
      return;
    }
    
    const result = await apiServices.dlpService.scanDocument(file, content);
    
    res.json(result);
  } catch (error) {
    next(error);
  }
});

// Get DLP scan result
app.get('/dlp/result/:fileId', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { fileId } = req.params;
    const { modifiedTime } = req.query;
    
    const result = apiServices.dlpService.getScanResult(fileId, modifiedTime as string);
    
    if (!result) {
      res.status(404).json({ error: 'Scan result not found' });
      return;
    }
    
    res.json(result);
  } catch (error) {
    next(error);
  }
});

// Generate compliance report
app.post('/compliance/report', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { regulations, periodStart, periodEnd, organization } = req.body;
    
    if (!regulations || !periodStart || !periodEnd || !organization) {
      res.status(400).json({ error: 'Missing required parameters' });
      return;
    }
    
    const report = await apiServices.complianceService.generateComplianceReport({
      auditService: apiServices.auditService,
      dlpService: apiServices.dlpService,
      regulations,
      periodStart: new Date(periodStart),
      periodEnd: new Date(periodEnd),
      organization
    });
    
    res.json(report);
  } catch (error) {
    next(error);
  }
});

// Get compliance report
app.get('/compliance/report/:reportId', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { reportId } = req.params;
    const report = apiServices.complianceService.getReport(reportId);
    
    if (!report) {
      res.status(404).json({ error: 'Report not found' });
      return;
    }
    
    res.json(report);
  } catch (error) {
    next(error);
  }
});

// Export compliance report
app.get('/compliance/report/:reportId/export', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { reportId } = req.params;
    const { format } = req.query;
    
    const validFormats = ['json', 'csv', 'pdf'];
    const exportFormat = validFormats.includes(format as string) ? format as 'json' | 'csv' | 'pdf' : 'json';
    
    const exportedData = apiServices.complianceService.exportReport(reportId, exportFormat);
    
    if (exportFormat === 'json') {
      res.setHeader('Content-Type', 'application/json');
      res.send(exportedData);
    } else if (exportFormat === 'csv') {
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', `attachment; filename="compliance-report-${reportId}.csv"`);
      res.send(exportedData);
    } else {
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="compliance-report-${reportId}.pdf"`);
      res.send(exportedData);
    }
  } catch (error) {
    next(error);
  }
});

// Get service statistics
app.get('/stats', authenticateToken, async (_req: Request, res: Response) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const stats = {
      audit: apiServices.auditService.getStats(),
      dlp: apiServices.dlpService.getStats(),
      compliance: apiServices.complianceService.getStats(),
      indexer: (apiServices.indexerService as any).getStats?.() || { notAvailable: true }
    };
    
    res.json(stats);
  } catch (error) {
    next(error);
  }
});

// Search documents
app.get('/search', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const { query, limit, offset } = req.query;
    
    if (!query) {
      res.status(400).json({ error: 'Query parameter is required' });
      return;
    }
    
    // This would integrate with the search index
    // For now, we'll return a placeholder response
    res.json({
      query,
      results: [],
      total: 0,
      limit: parseInt(limit as string) || 10,
      offset: parseInt(offset as string) || 0
    });
  } catch (error) {
    next(error);
  }
});

// Webhook endpoint for n8n integration - receives file updates from Google Drive
app.post('/webhook/n8n/drive', async (req: Request, res: Response) => {
  try {
    logger.info('Received n8n webhook for Google Drive file update', {
      component: 'ApiService',
      event: 'n8n_webhook_received',
      body: req.body
    });

    if (!apiServices) {
      logger.error('API services not initialized for n8n webhook', {
        component: 'ApiService',
        event: 'n8n_webhook_error'
      });
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }

    const { fileId, fileName, mimeType, chunks, embeddings } = req.body;

    // Validate required fields
    if (!fileId || !fileName) {
      logger.warn('Missing required fields in n8n webhook', {
        component: 'ApiService',
        event: 'n8n_webhook_invalid',
        missingFields: [!fileId ? 'fileId' : null, !fileName ? 'fileName' : null].filter(Boolean)
      });
      res.status(400).json({ error: 'Missing required fields: fileId and fileName are required' });
      return;
    }

    // Process the file update
    logger.info('Processing Google Drive file update', {
      component: 'ApiService',
      event: 'file_update_processing',
      fileId,
      fileName,
      mimeType
    });

    // If we have chunks and embeddings, process them for RAG
    if (chunks && embeddings && Array.isArray(chunks) && Array.isArray(embeddings)) {
      logger.info('Processing document chunks for RAG', {
        component: 'ApiService',
        event: 'rag_processing',
        fileId,
        chunkCount: chunks.length
      });

      // In a real implementation, this would:
      // 1. Store the chunks and embeddings in the vector database
      // 2. Update the search index
      // 3. Notify relevant Discord channels
    }

    // Acknowledge the webhook
    res.status(200).json({
      success: true,
      message: 'File update received and queued for processing',
      fileId,
      fileName
    });
  } catch (error) {
    logger.error('Error processing n8n webhook', {
      component: 'ApiService',
      event: 'n8n_webhook_error',
      error: error instanceof Error ? error.message : String(error)
    });
    res.status(500).json({
      error: 'Internal server error processing webhook',
      message: process.env['NODE_ENV'] === 'development' ? error instanceof Error ? error.message : String(error) : 'An error occurred'
    });
  }
});

// Add custom DLP pattern
app.post('/dlp/patterns', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const pattern = req.body;
    
    if (!pattern || !pattern.id || !pattern.name || !pattern.pattern) {
      res.status(400).json({ error: 'Invalid pattern data' });
      return;
    }
    
    apiServices.dlpService.addPattern(pattern);
    
    res.status(201).json({ message: 'Pattern added successfully' });
  } catch (error) {
    next(error);
  }
});

// Add custom compliance requirement
app.post('/compliance/requirements', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    
    const requirement = req.body;
    
    if (!requirement || !requirement.id || !requirement.name || !requirement.checkFunction) {
      res.status(400).json({ error: 'Invalid requirement data' });
      return;
    }
    
    apiServices.complianceService.addRequirement(requirement);
    
    res.status(201).json({ message: 'Requirement added successfully' });
  } catch (error) {
    next(error);
  }
});

// Apply error handling middleware
app.use(errorHandler);

// Handle 404 errors
app.use((_req: Request, res: Response) => {
  res.status(404).json({ error: 'Endpoint not found' });
});

// Start server function
export function startApiServer(config: BotConfig, services: ApiService): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      // Initialize services
      initializeApiServices(services);
      
      // Get port from config or default to 3000
      const port = (config as any).api?.port ?? (process.env['API_PORT'] ? parseInt(process.env['API_PORT'], 10) : 3000);
      
      // Start server
      const server = app.listen(port, () => {
        logger.info(`API server started on port ${port}`, { component: 'ApiService' });
        resolve();
      });
      
      // Handle server errors
      server.on('error', (error: Error) => {
        logger.error('API server error', { component: 'ApiService', error });
        reject(error);
      });
      
      // Graceful shutdown
      process.on('SIGTERM', () => {
        logger.info('SIGTERM received, shutting down API server', { component: 'ApiService' });
        server.close(() => {
          logger.info('API server closed', { component: 'ApiService' });
        });
      });
      
      process.on('SIGINT', () => {
        logger.info('SIGINT received, shutting down API server', { component: 'ApiService' });
        server.close(() => {
          logger.info('API server closed', { component: 'ApiService' });
        });
      });
    } catch (error) {
      reject(error);
    }
  });
}

export default app;