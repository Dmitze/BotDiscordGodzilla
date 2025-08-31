/**
 * Example usage of GoogleSheetsService
 * This demonstrates how to use the enhanced GoogleSheetsService with the google-spreadsheet library
 */

import { GoogleSheetsService } from '../src/services/GoogleSheetsService';
import type { BotConfig } from '../src/types';

async function example() {
  // Mock configuration (in a real app, this would come from environment variables)
  const config: BotConfig = {
    discord: {
      token: 'your-discord-token',
      clientId: 'your-client-id',
      prefix: '!',
      intents: ['Guilds', 'GuildMessages', 'MessageContent'],
    },
    google: {
      spreadsheetId: 'your-spreadsheet-id',
      driveFolderId: 'your-drive-folder-id',
      credentials: {
        client_email: 'your-service-account-email@project.iam.gserviceaccount.com',
        private_key: '-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n',
        project_id: 'your-project-id',
      },
    },
    ai: {
      provider: 'openai',
      openai: {
        apiKey: 'your-openai-api-key',
        model: 'gpt-3.5-turbo',
        maxTokens: 1000,
        temperature: 0.7,
      },
      ollama: {
        host: 'http://localhost:11434',
        model: 'llama2',
      },
    },
    cache: {
      redis: {
        host: 'localhost',
        port: 6379,
        password: '',
        database: 0,
      },
      ttl: 3600,
    },
    metrics: {
      enabled: true,
      port: 9090,
      path: '/metrics',
    },
    security: {
      rateLimitWindow: 60000,
      rateLimitMax: 100,
      adminRole: 'admin',
      botUserRole: 'bot-user',
    },
    performance: {
      cacheTTL: 300,
      maxSearchResults: 50,
      maxAnalysisRows: 1000,
      requestTimeout: 30000,
      maxRetries: 3,
    },
    logging: {
      level: 'info',
      maxFiles: 5,
      maxSize: '10m',
      directory: './logs',
    },
    drive: {
      allowedMime: ['*'],
      ttlListSec: 300,
      ttlTextSec: 300,
      maxResults: 1000,
      rateQps: 5,
      rateBurst: 10,
    },
    features: {
      defaultLocale: 'uk',
    },
  } as unknown as BotConfig;

  // Create an instance of GoogleSheetsService
  const googleSheetsService = new GoogleSheetsService(config);

  try {
    // Initialize the service
    await googleSheetsService.initialize();
    console.log('✅ GoogleSheetsService initialized successfully');

    // Example 1: List sheets in a spreadsheet
    const spreadsheetId = 'your-spreadsheet-id';
    const sheetNames = await googleSheetsService.listSheets(spreadsheetId);
    console.log('📋 Sheet names:', sheetNames);

    // Example 2: Get data from a specific sheet
    const sheetData = await googleSheetsService.getSheetData(spreadsheetId, 'Sheet1!A1:D10');
    console.log('📊 Sheet data:', sheetData);

    // Example 3: Search data
    const searchData = await googleSheetsService.searchData('test query', 20);
    console.log('🔍 Search results:', searchData);

    // Example 4: Extract text for chat
    const fileId = 'your-file-id';
    const textData = await googleSheetsService.extractTextForChat(fileId);
    console.log('📝 Extracted text:', textData.text);

    // Example 5: Read a range of data
    const rangeData = await googleSheetsService.readRange(spreadsheetId, 'Sheet1', 'A1:D10');
    console.log('📄 Range data:', rangeData);

    // Example 6: Find a sheet by name
    const sheetInfo = await googleSheetsService.findSheetByName(spreadsheetId, 'Sheet1');
    console.log('🔍 Found sheet:', sheetInfo);

    // Health check
    const health = await googleSheetsService.healthCheck();
    console.log('❤️ Health status:', health);

    // Get stats
    const stats = googleSheetsService.getStats();
    console.log('📈 Service stats:', stats);

  } catch (error) {
    console.error('❌ Error using GoogleSheetsService:', error);
  } finally {
    // Shutdown the service
    await googleSheetsService.shutdown();
    console.log('🛑 GoogleSheetsService shutdown complete');
  }
}

// Run the example
example().catch(console.error);