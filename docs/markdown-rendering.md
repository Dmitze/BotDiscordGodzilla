# Markdown Rendering Service

## Overview

The Markdown Rendering Service provides enhanced markdown rendering capabilities for the BotDiscordGodzilla project, leveraging the CordMd library to create beautiful formatted responses in Discord.

## Features

- **Text Rendering**: Enhanced markdown formatting for Discord messages
- **Image Rendering**: Convert markdown to images for complex documents
- **Code Block Syntax Highlighting**: Beautiful code formatting with language support
- **Table Formatting**: Improved table display in Discord
- **Input Sanitization**: Safe handling of user-provided markdown
- **Caching**: Performance optimization through rendered content caching
- **Memory Management**: Efficient handling of large documents
- **Metrics Collection**: Performance monitoring and statistics

## Usage

### Text Rendering

```typescript
const markdownService = interaction.client.serviceContainer.get('markdownRendering');
const renderedText = await markdownService.renderToText('# Hello World\nThis is **bold** text!');
```

### Image Rendering

```typescript
const markdownService = interaction.client.serviceContainer.get('markdownRendering');
const attachment = await markdownService.renderToImage('# Hello World\nThis is **bold** text!');
```

### Command Usage

Users can use the `/markdown` command to render markdown content:

```
/markdown content:"# Hello World
This is **bold** text!" format:text
```

## Configuration

The service can be configured through environment variables:

```env
# Markdown Rendering Configuration
MARKDOWN_THEME=dark
MARKDOWN_FONT_SIZE=16
MARKDOWN_MAX_WIDTH=800
MARKDOWN_MAX_HEIGHT=600
```

## API Reference

### `renderToText(markdown: string): Promise<string>`

Render markdown content as formatted Discord text.

### `renderToImage(markdown: string, options?: RenderOptions): Promise<AttachmentBuilder>`

Render markdown content as an image attachment.

### `validateMarkdown(markdown: string): ValidationResult`

Validate markdown content for syntax errors.

### `extractCodeBlocks(markdown: string): CodeBlock[]`

Extract code blocks from markdown content.

### `getMetrics(): RenderingMetrics`

Get performance metrics for the rendering service.

## Error Handling

The service includes comprehensive error handling:

- Invalid markdown syntax
- Rendering timeouts
- Memory overflow protection
- Font loading failures
- Canvas rendering errors

## Performance Optimization

### Caching

The service implements a caching mechanism to avoid re-rendering the same content:

```typescript
class MarkdownCache {
  private cache: Map<string, { buffer: Buffer; timestamp: number }>;
  private maxSize: number = 100;
  private ttl: number = 300000; // 5 minutes
}
```

### Memory Management

Memory usage is monitored to prevent overflow:

```typescript
class MemoryMonitor {
  private threshold: number = 500 * 1024 * 1024; // 500MB
}
```

## Internationalization

The service supports Ukrainian language content through proper Unicode handling and font support.

## Security

Input sanitization prevents XSS and other security issues:

```typescript
class InputSanitizer {
  public sanitizeMarkdown(markdown: string): string {
    // Remove potentially dangerous HTML
    let sanitized = markdown.replace(/<[^>]*>/g, '');
    
    // Limit length
    if (sanitized.length > 10000) {
      sanitized = sanitized.substring(0, 10000);
    }
    
    // Escape Discord-specific characters
    sanitized = sanitized.replace(/@/g, '@\u200b'); // Prevent mentions
    sanitized = sanitized.replace(/`/g, '\\`'); // Escape backticks
    
    return sanitized;
  }
}
```

## Testing

The service includes comprehensive unit and integration tests:

```bash
npm run test:services -- MarkdownRenderingService
npm run test:commands -- MarkdownCommand
```

## Future Enhancements

Planned features include:

1. Real-time collaborative markdown editing
2. Export to multiple formats (PDF, DOCX, HTML)
3. Version control for markdown documents
4. Template system for common document types
5. Integration with cloud storage providers