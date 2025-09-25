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
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';
import { z } from 'zod';

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

// Lazy singleton for search index
let searchIndexSingleton: SqliteSearchIndex | null = null;
function getSearchIndex(): SqliteSearchIndex {
  if (!searchIndexSingleton) {
    searchIndexSingleton = new SqliteSearchIndex();
    logger.info('Search index initialized', { component: 'ApiService' });
  }
  return searchIndexSingleton;
}

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
  
  if (token !== (process.env['API_ACCESS_TOKEN'] || '')) {
    res.status(403).json({ error: 'Invalid access token' });
    return;
  }
  
  next();
}

// Error handling middleware
function errorHandler(err: Error, req: Request, res: Response, _next: NextFunction): void {
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
app.get('/health', (_req: Request, res: Response) => {
  res.status(200).json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    service: 'Discord AI Assistant Bot API'
  });
});

// Search documents (real implementation)
app.get('/search', authenticateToken, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const schema = z.object({
      query: z.string().min(1, 'query is required'),
      limit: z.coerce.number().int().min(1).max(200).optional(),
      offset: z.coerce.number().int().min(0).optional(),
    });

    const parsed = schema.safeParse(req.query);
    if (!parsed.success) {
      res.status(400).json({ error: 'Invalid query parameters', issues: parsed.error.issues });
      return;
    }

    const q: any = parsed.data;
    const searchQuery: any = {
      text: q.query,
      limit: q.limit ?? 10,
      offset: q.offset ?? 0,
    };

    const index = getSearchIndex();
    const { hits, total } = await index.search(searchQuery);

    res.json({ query: q.query, results: hits, total, limit: searchQuery.limit, offset: searchQuery.offset });
  } catch (error) {
    next(error);
  }
});

// Webhook endpoint for n8n integration
app.post('/webhook/n8n/drive', async (req: Request, res: Response) => {
  try {
    if (!apiServices) {
      res.status(503).json({ error: 'API services not initialized' });
      return;
    }
    const { fileId, fileName } = req.body;
    if (!fileId || !fileName) {
      res.status(400).json({ error: 'Missing required fields: fileId and fileName are required' });
      return;
    }
    const mockDriveFile = {
      id: fileId,
      name: fileName,
      mimeType: req.body.mimeType || 'application/octet-stream',
      modifiedTime: new Date().toISOString()
    };
    await apiServices.indexerService.indexOneFileByMeta(mockDriveFile);
    res.status(200).json({ success: true, message: 'File update received' });
  } catch (error) {
    logger.error('Error processing n8n webhook', { error: error instanceof Error ? error.message : String(error) });
    res.status(500).json({ error: 'Internal server error' });
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
      initializeApiServices(services);
      const port = (config as any).api?.port ?? 3000;
      const server = app.listen(port, () => {
        logger.info(`API server started on port ${port}`, { component: 'ApiService' });
        resolve();
      });
      server.on('error', (error: Error) => {
        logger.error('API server error', { component: 'ApiService', error });
        reject(error);
      });
    } catch (error) {
      reject(error);
    }
  });
}

export default app;

