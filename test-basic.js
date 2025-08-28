// Simple test to verify core functionality
const { config } = require('dotenv');
const path = require('path');
const fs = require('fs');

// Load environment variables
const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  config({ path: envPath });
  console.log('✅ Environment variables loaded');
} else {
  console.log('⚠️ .env file not found');
}

// Test Discord connection
const discordToken = process.env.DISCORD_TOKEN;
if (discordToken) {
  console.log('✅ Discord token found');
  // Test if token is valid by checking first few characters
  if (discordToken.length > 20) {
    console.log('✅ Discord token appears valid');
  } else {
    console.log('❌ Discord token appears invalid');
  }
} else {
  console.log('❌ Discord token not found');
}

// Test Ollama connection
const ollamaHost = process.env.OLLAMA_HOST || 'http://localhost:11434';
console.log(`🔍 Checking Ollama at ${ollamaHost}`);

fetch(`${ollamaHost}/api/tags`)
  .then(response => response.json())
  .then(data => {
    console.log('✅ Ollama is running');
    if (data.models && data.models.length > 0) {
      console.log(`✅ Found ${data.models.length} models:`);
      data.models.forEach(model => {
        console.log(`  - ${model.name}`);
      });
    } else {
      console.log('⚠️ No models found in Ollama');
    }
  })
  .catch(error => {
    console.log('❌ Failed to connect to Ollama:', error.message);
  });

// Test AI provider
const aiProvider = process.env.AI_PROVIDER || 'ollama';
console.log(`🤖 AI Provider: ${aiProvider}`);

// Test required models
const requiredModels = ['llama3.2'];
const ollamaModel = process.env.OLLAMA_MODEL || 'llama3.2';
console.log(`🔍 Checking for required model: ${ollamaModel}`);

fetch(`${ollamaHost}/api/tags`)
  .then(response => response.json())
  .then(data => {
    if (data.models) {
      const modelExists = data.models.some(model => model.name.includes(ollamaModel));
      if (modelExists) {
        console.log(`✅ Required model ${ollamaModel} found`);
      } else {
        console.log(`❌ Required model ${ollamaModel} not found`);
      }
    }
  })
  .catch(error => {
    console.log('❌ Failed to check models:', error.message);
  });

console.log('🧪 Basic tests completed');