# Ollama Integration

This document describes the Ollama integration in the BotDiscordGodzilla project, which enables local AI model capabilities for the Discord bot.

## Overview

The Ollama integration provides the following capabilities:
- Interaction with local AI models through Ollama
- Conversation history management per Discord channel
- Model management (listing, pulling)
- Context-aware responses with conversation memory

## Architecture

The integration consists of two main components:
1. **OllamaService** - Core service that handles communication with the Ollama API
2. **OllamaCommand** - Discord slash command that exposes Ollama functionality to users

### OllamaService

The OllamaService is responsible for:
- Managing connection to the Ollama API
- Handling conversation history using Redis caching
- Generating responses from AI models
- Managing model operations (list, pull)

#### Key Methods

- `generate(prompt, options)` - Generate a response from the AI model
- `resetChannelHistory(channelId)` - Clear conversation history for a channel
- `listModels()` - List available models
- `pullModel(modelName)` - Download a model
- `healthCheck()` - Check if Ollama is available

### OllamaCommand

The OllamaCommand provides a Discord interface for users to:
- Send prompts to the AI model
- Specify which model to use
- Reset conversation history

## Configuration

### Environment Variables

```env
# Ollama Configuration
OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=llama3
OLLAMA_CTX=2048
OLLAMA_CHAT_MAX_LENGTH=500
```

### Service Configuration

```typescript
interface OllamaConfig {
  host: string;
  model: string;
  ctx: number;
  chatMaxLength: number;
}
```

## Usage

### Discord Command

Users can interact with the Ollama integration through the `/ollama` slash command:

```
/ollama prompt:"Explain quantum computing" model:"llama3" reset:false
```

Options:
- `prompt` (required) - The prompt to send to the AI model
- `model` (optional) - Specify which model to use
- `reset` (optional) - Reset the conversation history for this channel

### Examples

1. **Simple query:**
   ```
   /ollama prompt:"What is the capital of France?"
   ```

2. **Using a specific model:**
   ```
   /ollama prompt:"Write a poem" model:"mistral"
   ```

3. **Reset conversation history:**
   ```
   /ollama prompt:"Hello" reset:true
   ```

## Conversation History

The Ollama integration maintains conversation history per Discord channel using Redis caching:
- History is stored for 7 days by default
- Each channel has its own conversation context
- History can be reset using the `reset` option

## Supported Models

The integration supports any model available through Ollama:
- Llama 3 (llama3, llama3.2)
- Mistral (mistral, mixtral)
- Gemma (gemma, gemma2)
- Phi (phi, phi3)
- And many others

Models can be pulled using the `pullModel` method or through Ollama's CLI.

## Error Handling

The service includes comprehensive error handling:
- Network errors when connecting to Ollama
- API errors from Ollama
- Cache errors when managing conversation history
- Model errors during generation

## Performance Considerations

- Response times depend on the local hardware and model size
- Conversation history is cached to reduce repeated context transmission
- Large responses are truncated to fit Discord's message limits

## Security

- Input sanitization to prevent injection attacks
- Rate limiting through Discord's built-in mechanisms
- Secure storage of conversation history in Redis

## Testing

Unit tests are provided for the OllamaService:
- Configuration validation
- Method functionality
- Error handling
- Cache integration

## Future Enhancements

Planned improvements:
- Streaming responses for faster interaction
- Multi-modal support (images, audio)
- Advanced model management (unload, delete)
- Custom system prompts per channel
- Model performance monitoring