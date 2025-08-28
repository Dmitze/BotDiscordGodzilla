#!/usr/bin/env ts-node

/**
 * Test Google Drive Access
 * Simple script to test if the bot can access Google Drive
 */

import { config } from 'dotenv';
import path from 'path';
import { GoogleService } from '../services/GoogleService';
import { Config } from '../config/Config';

// Load environment variables
config({ path: path.resolve(process.cwd(), '.env') });

async function testGoogleAccess() {
  console.log('🔍 Testing Google Drive access...');
  
  try {
    // Load configuration
    const botConfig = Config.load();
    
    // Initialize Google Service
    const googleService = new GoogleService(botConfig);
    
    // Try to initialize the service
    await googleService.initialize();
    
    console.log('✅ Google Service initialized successfully');
    
    // Try to list files in the configured folder
    console.log('📂 Attempting to list files in Drive folder...');
    const files = await googleService.searchFiles('');
    
    console.log(`✅ Found ${files.length} files in Drive folder:`);
    files.forEach((file: any, index: number) => {
      console.log(`  ${index + 1}. ${file.name} (${file.id}) - ${file.mimeType}`);
    });
    
    console.log('\n🎉 Google Drive access test completed successfully!');
    
  } catch (error) {
    console.error('❌ Error testing Google Drive access:', error);
    if (error instanceof Error) {
      console.error('Error message:', error.message);
      console.error('Stack trace:', error.stack);
    }
  }
}

// Run the test
if (require.main === module) {
  testGoogleAccess();
}