import { renderMarkdown, validateMarkdown } from 'cordmd';
import { AttachmentBuilder } from 'discord.js';
import logger from '@/utils/logger';
import type { BotConfig } from '@/types';

// Define interfaces
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

interface RenderingMetrics {
  renderCount: number;
  averageRenderTime: number;
  errorCount: number;
  cacheHitRate: number;
}

// Cache implementation
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

// Memory monitor
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

// Input sanitizer
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

// Metrics collector
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

export class MarkdownRenderingService {
  private static instance: MarkdownRenderingService;
  private cache: MarkdownCache;
  private memoryMonitor: MemoryMonitor;
  private sanitizer: InputSanitizer;
  private metricsCollector: MetricsCollector;
  private config: BotConfig;
  
  private constructor(config: BotConfig) {
    this.config = config;
    this.cache = new MarkdownCache();
    this.memoryMonitor = new MemoryMonitor();
    this.sanitizer = new InputSanitizer();
    this.metricsCollector = new MetricsCollector();
  }
  
  public static getInstance(config?: BotConfig): MarkdownRenderingService {
    if (!MarkdownRenderingService.instance) {
      if (!config) {
        throw new Error('Configuration required for first initialization');
      }
      MarkdownRenderingService.instance = new MarkdownRenderingService(config);
    }
    return MarkdownRenderingService.instance;
  }
  
  /**
   * Render markdown to formatted text
   */
  public async renderToText(markdown: string): Promise<string> {
    const startTime = Date.now();
    let success = false;
    
    try {
      // Sanitize input
      const sanitizedMarkdown = this.sanitizer.sanitizeMarkdown(markdown);
      
      // Validate markdown first
      const validation = this.validateMarkdown(sanitizedMarkdown);
      if (!validation.isValid) {
        logger.warn('Invalid markdown provided for rendering', { errors: validation.errors });
      }
      
      // For text rendering, we'll use Discord's native markdown support
      // with some enhancements for better formatting
      const result = this.enhanceMarkdownForDiscord(sanitizedMarkdown);
      success = true;
      return result;
    } catch (error) {
      logger.error('Error rendering markdown to text', { error });
      throw new Error('Failed to render markdown to text');
    } finally {
      this.metricsCollector.recordRender(startTime, success);
    }
  }
  
  /**
   * Render markdown to image buffer
   */
  public async renderToImage(markdown: string, options?: RenderOptions): Promise<AttachmentBuilder> {
    const startTime = Date.now();
    let success = false;
    
    try {
      // Check memory first
      if (!this.memoryMonitor.checkMemory().safe) {
        await this.memoryMonitor.waitForSafeMemory();
      }
      
      // Sanitize input
      const sanitizedMarkdown = this.sanitizer.sanitizeMarkdown(markdown);
      
      // Check cache first
      const cacheKey = this.generateCacheKey(sanitizedMarkdown, options);
      const cachedBuffer = this.cache.get(cacheKey);
      if (cachedBuffer) {
        logger.debug('Returning cached markdown rendering');
        success = true;
        const attachment = new AttachmentBuilder(cachedBuffer, { name: 'markdown-render.png' });
        return attachment;
      }
      
      // Validate markdown first
      const validation = this.validateMarkdown(sanitizedMarkdown);
      if (!validation.isValid) {
        logger.warn('Invalid markdown provided for image rendering', { errors: validation.errors });
      }
      
      // Use CordMd to render to image
      const buffer = await renderMarkdown(sanitizedMarkdown);
      
      // Cache the result
      this.cache.set(cacheKey, buffer);
      
      // Create Discord attachment
      const attachment = new AttachmentBuilder(buffer, { name: 'markdown-render.png' });
      success = true;
      return attachment;
    } catch (error) {
      logger.error('Error rendering markdown to image', { error });
      throw new Error('Failed to render markdown to image');
    } finally {
      this.metricsCollector.recordRender(startTime, success);
    }
  }
  
  /**
   * Validate markdown content
   */
  public validateMarkdown(markdown: string): ValidationResult {
    try {
      // Use CordMd validation
      validateMarkdown(markdown);
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
    let match: RegExpExecArray | null;
    let lineCounter = 1;
    
    while ((match = codeBlockRegex.exec(markdown)) !== null) {
      const [, language, content] = match;
      const startLine = lineCounter;
      const lines = content ? content.split('\n').length : 0;
      const endLine = startLine + lines - 1;
      
      if (content) {
        codeBlocks.push({
          language: language || 'text',
          content,
          startLine,
          endLine
        });
      }
      
      lineCounter += lines + 2; // +2 for the opening and closing ```
    }
    
    return codeBlocks;
  }
  
  /**
   * Get rendering metrics
   */
  public getMetrics(): RenderingMetrics {
    return this.metricsCollector.getMetrics();
  }
  
  /**
   * Get service configuration
   */
  public getConfig(): BotConfig {
    return this.config;
  }
  
  /**
   * Enhance markdown for better Discord display
   */
  private enhanceMarkdownForDiscord(markdown: string): string {
    // Apply Discord-specific formatting enhancements
    let enhanced = markdown;
    
    // Improve code block formatting
    enhanced = enhanced.replace(/```(\w+)\n([\s\S]*?)```/g, (_match, lang, code) => {
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
  
  /**
   * Generate cache key for rendered content
   */
  private generateCacheKey(markdown: string, options?: RenderOptions): string {
    const optionsString = options ? JSON.stringify(options) : '';
    return `markdown_${markdown.length}_${optionsString}`;
  }
}