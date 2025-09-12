
import { OllamaService } from '../src/services/OllamaService';
import type { BotConfig } from '../src/types';

async function main() {
  // Configuration for the Ollama service
  const config: BotConfig = {
    ai: {
      ollama: {
        host: process.env.OLLAMA_HOST || 'http://localhost:11434',
        model: process.env.OLLAMA_MODEL || 'llama3',
        ctx: parseInt(process.env.OLLAMA_CTX || '2048'),
        chatMaxLength: parseInt(process.env.OLLAMA_CHAT_MAX_LENGTH || '500'),
      },
    },
  } as BotConfig;

  // Create the Ollama service
  const ollamaService = new OllamaService(config);

  try {
    // Initialize the service
    console.log('Initializing Ollama service...');
    await ollamaService.initialize();
    
    // Check health
    const health = await ollamaService.healthCheck();
    console.log('Health check:', health);
    
    if (!health.healthy) {
      console.error('Ollama is not available:', health.message);
      return;
    }
    
    // List available models
    console.log('\nAvailable models:');
    const models = await ollamaService.listModels();
    models.forEach((model: any) => {
      console.log(`- ${model.name} (size: ${model.size})`);
    });
    
    // Generate a simple response
    console.log('\nGenerating response...');
    const response = await ollamaService.generate('Hello, what is your name?', {
      temperature: 0.7,
    });
    console.log('Response:', response);
    
    // Example of conversation with history
    console.log('\nTesting conversation history...');
    const channelId = 'example-channel-123';
    
    const response1 = await ollamaService.generate('What is 2+2?', {
      channelId,
    });
    console.log('Response 1:', response1);
    
    const response2 = await ollamaService.generate('What did I just ask?', {
      channelId,
    });
    console.log('Response 2:', response2);
    
    // Reset history
    console.log('\nResetting conversation history...');
    await ollamaService.resetChannelHistory(channelId);
    console.log('History reset');
    
    // Check final stats
    const stats = ollamaService.getStats();
    console.log('\nService stats:', stats);
    
  } catch (error) {
    console.error('Error:', error);
  } finally {
    // Shutdown the service
    await ollamaService.shutdown();
    console.log('\nOllama service shutdown complete');
  }
}

// Run the example
if (require.main === module) {
  main().catch(console.error);
}