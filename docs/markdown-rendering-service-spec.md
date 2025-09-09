# Markdown Rendering Service Specification

## Overview
This document specifies the implementation of a Markdown Rendering Service for the BotDiscordGodzilla project, leveraging the CordMd library to provide enhanced markdown rendering capabilities in Discord responses.

## Requirements

### Functional Requirements
1. Render markdown text to formatted Discord messages
2. Support syntax highlighting for code blocks
3. Provide table formatting improvements
4. Enable image embedding in responses
5. Support mathematical expressions rendering
6. Allow customization of rendering options (colors, fonts, sizes)

### Non-Functional Requirements
1. Response time under 2 seconds for typical markdown content
2. Support for Ukrainian language content
3. Error handling for malformed markdown
4. Memory efficient processing for large documents
5. Compatibility with existing Discord message formatting

## Architecture

### Component Diagram
```
┌─────────────────────┐    ┌──────────────────────┐    ┌──────────────────────┐
│   Discord Client    │────│  Markdown Service    │────│    CordMd Library    │
└─────────────────────┘    └──────────────────────┘    └──────────────────────┘
                                    │                           │
                                    ▼                           ▼
                         ┌──────────────────────┐    ┌──────────────────────┐
                         │   Discord.js API     │    │   Canvas Rendering   │
                         └──────────────────────┘    └──────────────────────┘
```

### Service Interface
```typescript
interface MarkdownRenderingService {
  renderToText(markdown: string): Promise<string>;
  renderToImage(markdown: string, options?: RenderOptions): Promise<Buffer>;
  validateMarkdown(markdown: string): ValidationResult;
  extractCodeBlocks(markdown: string): CodeBlock[];
}

interface RenderOptions {
  theme?: 'light' | 'dark';
  fontSize?: number;
  fontFamily?: string;
  maxWidth?: number;
  maxHeight?: number;
}

interface ValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

interface CodeBlock {
  language: string;
  content: string;
  startLine: number;
  endLine: number;
}
```

## Implementation Details

### Core Service Implementation
The service will be implemented as a singleton class that wraps the CordMd library functionality:

```typescript
import { renderMarkdown, validateMarkdown } from 'cordmd';
import { AttachmentBuilder } from 'discord.js';
import logger from '@/utils/logger';

export class MarkdownRenderingService {
  private static instance: MarkdownRenderingService;
  
  private constructor() {}
  
  public static getInstance(): MarkdownRenderingService {
    if (!MarkdownRenderingService.instance) {
      MarkdownRenderingService.instance = new MarkdownRenderingService();
    }
    return MarkdownRenderingService.instance;
  }
  
  /**
   * Render markdown to formatted text
   */
  public async renderToText(markdown: string): Promise<string> {
    try {
      // Validate markdown first
      const validation = this.validateMarkdown(markdown);
      if (!validation.isValid) {
        logger.warn('Invalid markdown provided for rendering', { errors: validation.errors });
      }
      
      // For text rendering, we'll use Discord's native markdown support
      // with some enhancements for better formatting
      return this.enhanceMarkdownForDiscord(markdown);
    } catch (error) {
      logger.error('Error rendering markdown to text', { error });
      throw new Error('Failed to render markdown to text');
    }
  }
  
  /**
   * Render markdown to image buffer
   */
  public async renderToImage(markdown: string, options?: RenderOptions): Promise<Buffer> {
    try {
      // Validate markdown first
      const validation = this.validateMarkdown(markdown);
      if (!validation.isValid) {
        logger.warn('Invalid markdown provided for image rendering', { errors: validation.errors });
      }
      
      // Use CordMd to render to image
      const buffer = await renderMarkdown(markdown);
      return buffer;
    } catch (error) {
      logger.error('Error rendering markdown to image', { error });
      throw new Error('Failed to render markdown to image');
    }
  }
  
  /**
   * Validate markdown content
   */
  public validateMarkdown(markdown: string): ValidationResult {
    try {
      // Use CordMd validation
      const validated = validateMarkdown(markdown);
      return {
        isValid: true,
        errors: [],
        warnings: []
      };
    } catch (error) {
      return {
        isValid: false,
        errors: [error instanceof Error ? error.message : 'Invalid markdown'],
        warnings: []
      };
    }
  }
  
  /**
   * Extract code blocks from markdown
   */
  public extractCodeBlocks(markdown: string): CodeBlock[] {
    const codeBlocks: CodeBlock[] = [];
    const codeBlockRegex = /```(\w*)\n([\s\S]*?)```/g;
    let match;
    let lineCounter = 1;
    
    while ((match = codeBlockRegex.exec(markdown)) !== null) {
      const [, language, content] = match;
      const startLine = lineCounter;
      const lines = content.split('\n').length;
      const endLine = startLine + lines - 1;
      
      codeBlocks.push({
        language: language || 'text',
        content,
        startLine,
        endLine
      });
      
      lineCounter += lines + 2; // +2 for the opening and closing ```
    }
    
    return codeBlocks;
  }
  
  /**
   * Enhance markdown for better Discord display
   */
  private enhanceMarkdownForDiscord(markdown: string): string {
    // Apply Discord-specific formatting enhancements
    let enhanced = markdown;
    
    // Improve code block formatting
    enhanced = enhanced.replace(/```(\w+)\n([\s\S]*?)```/g, (match, lang, code) => {
      return `\`\`\`${lang}\n${code.trim()}\n\`\`\``;
    });
    
    // Improve table formatting
    enhanced = enhanced.replace(/(\|[^\n]*\|\n\|[^\n]*\|\n\|[^\n]*\|)/g, (table) => {
      return `\n${table}\n`;
    });
    
    // Limit message length for Discord
    if (enhanced.length > 2000) {
      enhanced = enhanced.substring(0, 1997) + '...';
    }
    
    return enhanced;
  }
}
```

### Command Integration
The markdown rendering service will be integrated into existing commands through dependency injection:

```typescript
import { BaseCommand } from '@/commands/BaseCommand';
import { MarkdownRenderingService } from '@/services/MarkdownRenderingService';

export class EnhancedAICommand extends BaseCommand {
  private markdownService: MarkdownRenderingService;
  
  constructor(config: BotConfig, googleService?: GoogleService) {
    super(/* ... */);
    this.markdownService = MarkdownRenderingService.getInstance();
  }
  
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    // Generate AI response
    const aiResponse = await this.generateAIResponse(options);
    
    // Render response with markdown
    const renderedResponse = await this.markdownService.renderToText(aiResponse);
    
    // Send response to Discord
    await interaction.reply({ content: renderedResponse });
  }
}
```

## Configuration

### Environment Variables
```env
# Markdown Rendering Configuration
MARKDOWN_THEME=dark
MARKDOWN_FONT_SIZE=16
MARKDOWN_MAX_WIDTH=800
MARKDOWN_MAX_HEIGHT=600
```

### Service Configuration
```typescript
interface MarkdownConfig {
  theme: 'light' | 'dark';
  fontSize: number;
  fontFamily: string;
  maxWidth: number;
  maxHeight: number;
  enableImageRendering: boolean;
  enableSyntaxHighlighting: boolean;
}

const defaultMarkdownConfig: MarkdownConfig = {
  theme: 'dark',
  fontSize: 16,
  fontFamily: 'sans-serif',
  maxWidth: 800,
  maxHeight: 600,
  enableImageRendering: true,
  enableSyntaxHighlighting: true
};
```

## Error Handling

### Common Error Scenarios
1. Invalid markdown syntax
2. Rendering timeout
3. Memory overflow for large documents
4. Font loading failures
5. Canvas rendering errors

### Error Recovery Strategies
1. Fallback to plain text rendering
2. Truncate large content
3. Retry rendering with reduced quality
4. Return error message to user with suggestion

## Testing Strategy

### Unit Tests
```typescript
describe('MarkdownRenderingService', () => {
  let service: MarkdownRenderingService;
  
  beforeEach(() => {
    service = MarkdownRenderingService.getInstance();
  });
  
  describe('renderToText', () => {
    it('should render simple markdown correctly', async () => {
      const markdown = '# Hello World\nThis is **bold** text!';
      const result = await service.renderToText(markdown);
      expect(result).toContain('Hello World');
      expect(result).toContain('**bold**');
    });
    
    it('should handle code blocks', async () => {
      const markdown = '```javascript\nconsole.log("Hello");\n```';
      const result = await service.renderToText(markdown);
      expect(result).toContain('```javascript');
    });
  });
  
  describe('renderToImage', () => {
    it('should generate image buffer', async () => {
      const markdown = '# Hello World';
      const buffer = await service.renderToImage(markdown);
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });
  
  describe('validateMarkdown', () => {
    it('should validate correct markdown', () => {
      const markdown = '# Valid Markdown';
      const result = service.validateMarkdown(markdown);
      expect(result.isValid).toBe(true);
    });
  });
});
```

### Integration Tests
```typescript
describe('Markdown Rendering Integration', () => {
  it('should integrate with Discord commands', async () => {
    // Test integration with actual Discord command execution
    // This would require a mock Discord client
  });
  
  it('should handle Ukrainian language content', async () => {
    const ukrainianMarkdown = '# Привіт Світ\nЦе **жирний** текст!';
    const result = await service.renderToText(ukrainianMarkdown);
    expect(result).toContain('Привіт Світ');
  });
});
```

## Performance Optimization

### Caching Strategy
```typescript
class MarkdownCache {
  private cache: Map<string, { buffer: Buffer; timestamp: number }>;
  private maxSize: number;
  private ttl: number;
  
  constructor(maxSize: number = 100, ttl: number = 300000) { // 5 minutes
    this.cache = new Map();
    this.maxSize = maxSize;
    this.ttl = ttl;
  }
  
  public get(key: string): Buffer | null {
    const entry = this.cache.get(key);
    if (!entry) return null;
    
    if (Date.now() - entry.timestamp > this.ttl) {
      this.cache.delete(key);
      return null;
    }
    
    return entry.buffer;
  }
  
  public set(key: string, buffer: Buffer): void {
    if (this.cache.size >= this.maxSize) {
      // Remove oldest entry
      const firstKey = this.cache.keys().next().value;
      if (firstKey) this.cache.delete(firstKey);
    }
    
    this.cache.set(key, { buffer, timestamp: Date.now() });
  }
}
```

### Memory Management
```typescript
class MemoryMonitor {
  private threshold: number;
  
  constructor(threshold: number = 500 * 1024 * 1024) { // 500MB
    this.threshold = threshold;
  }
  
  public checkMemory(): { safe: boolean; usage: number } {
    const usage = process.memoryUsage().heapUsed;
    return {
      safe: usage < this.threshold,
      usage
    };
  }
  
  public async waitForSafeMemory(): Promise<void> {
    while (!this.checkMemory().safe) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }
  }
}
```

## Internationalization Support

### Ukrainian Language Support
The service will support Ukrainian language content through:
1. Proper Unicode handling
2. Ukrainian font support
3. Localization of error messages
4. RTL language support (if needed in future)

## Security Considerations

### Input Sanitization
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

## Monitoring and Logging

### Performance Metrics
```typescript
interface RenderingMetrics {
  renderCount: number;
  averageRenderTime: number;
  errorCount: number;
  cacheHitRate: number;
}

class MetricsCollector {
  private metrics: RenderingMetrics = {
    renderCount: 0,
    averageRenderTime: 0,
    errorCount: 0,
    cacheHitRate: 0
  };
  
  public recordRender(startTime: number, success: boolean): void {
    this.metrics.renderCount++;
    const renderTime = Date.now() - startTime;
    this.metrics.averageRenderTime = 
      (this.metrics.averageRenderTime * (this.metrics.renderCount - 1) + renderTime) / this.metrics.renderCount;
    
    if (!success) {
      this.metrics.errorCount++;
    }
  }
  
  public getMetrics(): RenderingMetrics {
    return { ...this.metrics };
  }
}
```

## Deployment Considerations

### Docker Configuration
```dockerfile
# Install canvas dependencies
RUN apt-get update && apt-get install -y \
    libcairo2-dev \
    libjpeg-dev \
    libpango1.0-dev \
    libgif-dev \
    build-essential \
    libpixman-1-0 \
    libcairo2 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Install node-canvas dependencies
RUN npm install --build-from-source canvas
```

### Kubernetes Deployment
```yaml
apiVersion: apps/v1
kind: Deployment
metadata:
  name: discord-bot-markdown
spec:
  template:
    spec:
      containers:
      - name: bot
        resources:
          requests:
            memory: "512Mi"
            cpu: "250m"
          limits:
            memory: "1Gi"
            cpu: "500m"
```

## Future Enhancements

### Planned Features
1. Real-time collaborative markdown editing
2. Export to multiple formats (PDF, DOCX, HTML)
3. Version control for markdown documents
4. Template system for common document types
5. Integration with cloud storage providers

### Research Areas
1. AI-powered markdown generation
2. Voice-to-markdown conversion
3. Handwriting recognition for markdown
4. Advanced visualization capabilities
5. Cross-platform synchronization

## Conclusion

This specification provides a comprehensive plan for implementing a Markdown Rendering Service in the BotDiscordGodzilla project. By leveraging the CordMd library and following best practices for Discord bot development, we can provide enhanced markdown rendering capabilities that will significantly improve the user experience, especially for non-technical users who need to work with formatted documents in Discord.