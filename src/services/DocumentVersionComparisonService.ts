import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';
import type { AIService } from './AIService';

export interface DocumentVersion {
  fileId: string;
  versionId: string;
  name: string;
  modifiedTime: string;
  content: string;
  author?: string;
  size?: number;
}

export interface VersionComparison {
  fileId: string;
  fileName: string;
  versions: {
    versionId: string;
    modifiedTime: string;
    author?: string;
  }[];
  changes: VersionChange[];
  summary: string;
  keyAdditions: string[];
  keyRemovals: string[];
  overallSentiment: 'positive' | 'negative' | 'neutral';
  generatedAt: Date;
}

export interface VersionChange {
  type: 'addition' | 'removal' | 'modification';
  content: string;
  position?: number;
  context?: string;
}

export interface ComparisonOptions {
  includeSummary?: boolean;
  includeKeyChanges?: boolean;
  language?: string;
  detailLevel?: 'high' | 'medium' | 'low';
}

export class DocumentVersionComparisonService extends BaseService {
  private aiService: AIService | null = null;
  private comparisons: Map<string, VersionComparison> = new Map();
  private readonly MAX_COMPARISON_CACHE = 100;

  constructor(config: BotConfig) {
    super('DocumentVersionComparisonService', config);
  }

  /**
   * Initialize service with required dependencies
   */
  initializeServices(aiService: AIService): void {
    this.aiService = aiService;
  }

  /**
   * Compare two versions of a document
   */
  async compareDocumentVersions(
    file: DriveFile,
    versions: DocumentVersion[],
    options: ComparisonOptions = {}
  ): Promise<VersionComparison> {
    try {
      // Validate input
      if (versions.length < 2) {
        throw new Error('At least two versions are required for comparison');
      }

      // Sort versions by modification time
      const sortedVersions = [...versions].sort((a, b) =>
        new Date(a.modifiedTime).getTime() - new Date(b.modifiedTime).getTime()
      );

      // Create cache key
      const cacheKey = `${file.id}-${sortedVersions.map(v => v.versionId).join('-')}`;
      
      // Check cache
      const cachedComparison = this.comparisons.get(cacheKey);
      if (cachedComparison) {
        logger.debug('Returning cached version comparison', {
          component: 'DocumentVersionComparisonService',
          fileId: file.id,
        });
        return cachedComparison;
      }

      // Perform text-based comparison
      const changes = this.performTextComparison(sortedVersions);
      
      // Generate AI-powered summary if requested
      let summary = '';
      let keyAdditions: string[] = [];
      let keyRemovals: string[] = [];
      let overallSentiment: 'positive' | 'negative' | 'neutral' = 'neutral';
      
      if (options.includeSummary || options.includeKeyChanges) {
        if (this.aiService) {
          try {
            const aiSummary = await this.generateAIComparisonSummary(
              file,
              sortedVersions,
              changes,
              options
            );
            summary = aiSummary.summary;
            keyAdditions = aiSummary.keyAdditions;
            keyRemovals = aiSummary.keyRemovals;
            overallSentiment = aiSummary.sentiment;
          } catch (error) {
            logger.warn('Failed to generate AI summary for version comparison', {
              component: 'DocumentVersionComparisonService',
              fileId: file.id,
              error: error instanceof Error ? error.message : String(error),
            });
            // Fall back to basic summary
            summary = this.generateBasicSummary(changes);
          }
        } else {
          // Fall back to basic summary
          summary = this.generateBasicSummary(changes);
        }
      }

      // Create version comparison object
      const comparison: VersionComparison = {
        fileId: file.id,
        fileName: file.name || 'Untitled',
        versions: sortedVersions.map(v => ({
          versionId: v.versionId,
          modifiedTime: v.modifiedTime,
          ...(v.author !== undefined && { author: v.author })
        })),
        changes,
        summary,
        keyAdditions,
        keyRemovals,
        overallSentiment,
        generatedAt: new Date(),
      };

      // Cache the comparison
      this.cacheComparison(cacheKey, comparison);
      
      logger.info('Document version comparison completed', {
        component: 'DocumentVersionComparisonService',
        fileId: file.id,
        versionCount: versions.length,
        changeCount: changes.length,
      });

      return comparison;
    } catch (error) {
      logger.error('Error comparing document versions', {
        component: 'DocumentVersionComparisonService',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error),
      });
      
      throw error;
    }
  }

  /**
   * Perform basic text comparison between versions
   */
  private performTextComparison(versions: DocumentVersion[]): VersionChange[] {
    const changes: VersionChange[] = [];
    
    // For simplicity, we'll compare the first two versions
    // In a more advanced implementation, we'd compare all versions
    if (versions.length >= 2) {
      const oldVersion = versions[0];
      const newVersion = versions[1];
      
      // Check if versions are defined
      if (oldVersion && newVersion) {
        // Simple line-by-line comparison
        const oldLines = oldVersion.content.split('\n');
        const newLines = newVersion.content.split('\n');
        
        // Find additions and removals
        const oldSet = new Set(oldLines);
        const newSet = new Set(newLines);
        
        // Find added lines
        newLines.forEach((line, index) => {
          if (!oldSet.has(line) && line.trim() !== '') {
            changes.push({
              type: 'addition',
              content: line,
              position: index,
            });
          }
        });
        
        // Find removed lines
        oldLines.forEach((line, index) => {
          if (!newSet.has(line) && line.trim() !== '') {
            changes.push({
              type: 'removal',
              content: line,
              position: index,
            });
          }
        });
      }
    }
    
    return changes;
  }

  /**
   * Generate AI-powered comparison summary
   */
  private async generateAIComparisonSummary(
    file: DriveFile,
    versions: DocumentVersion[],
    changes: VersionChange[],
    options: ComparisonOptions
  ): Promise<{ summary: string; keyAdditions: string[]; keyRemovals: string[]; sentiment: 'positive' | 'negative' | 'neutral' }> {
    if (!this.aiService) {
      throw new Error('AI service not initialized');
    }

    const language = options.language || 'Ukrainian';
    const detailLevel = options.detailLevel || 'medium';
    
    // Create prompt for AI
    let prompt = `Compare the following versions of document "${file.name}" and provide a summary in ${language}.\n\n`;
    
    prompt += `Version history:\n`;
    versions.forEach((version, index) => {
      prompt += `${index + 1}. ${version.modifiedTime} (Version: ${version.versionId})\n`;
    });
    
    prompt += `\nKey changes detected:\n`;
    const additions = changes.filter(c => c.type === 'addition');
    const removals = changes.filter(c => c.type === 'removal');
    
    if (additions.length > 0) {
      prompt += `Additions (${additions.length}):\n`;
      additions.slice(0, 5).forEach((change, index) => {
        prompt += `${index + 1}. ${change.content.substring(0, 100)}...\n`;
      });
    }
    
    if (removals.length > 0) {
      prompt += `\nRemovals (${removals.length}):\n`;
      removals.slice(0, 5).forEach((change, index) => {
        prompt += `${index + 1}. ${change.content.substring(0, 100)}...\n`;
      });
    }
    
    prompt += `\nPlease provide:
    1. A concise summary of the changes
    2. Key additions (up to 5)
    3. Key removals (up to 5)
    4. Overall sentiment of the changes (positive, negative, or neutral)
    
    Format your response as:
    Summary: [your summary]
    Additions: [bullet points]
    Removals: [bullet points]
    Sentiment: [positive|negative|neutral]`;
    
    const aiResponse = await this.aiService.generateResponse(prompt, {
      useCache: true,
      maxTokens: 500,
    });
    
    // Parse AI response
    return this.parseAIComparisonResponse(aiResponse.content);
  }

  /**
   * Parse AI comparison response
   */
  private parseAIComparisonResponse(response: string): { 
    summary: string; 
    keyAdditions: string[]; 
    keyRemovals: string[]; 
    sentiment: 'positive' | 'negative' | 'neutral' 
  } {
    const lines = response.split('\n');
    let summary = '';
    const keyAdditions: string[] = [];
    const keyRemovals: string[] = [];
    let sentiment: 'positive' | 'negative' | 'neutral' = 'neutral';
    
    let currentSection: 'summary' | 'additions' | 'removals' | 'sentiment' | null = null;
    
    for (const line of lines) {
      const trimmedLine = line.trim();
      if (!trimmedLine) continue;
      
      if (trimmedLine.startsWith('Summary:')) {
        currentSection = 'summary';
        summary = trimmedLine.substring(8).trim();
      } else if (trimmedLine.startsWith('Additions:')) {
        currentSection = 'additions';
      } else if (trimmedLine.startsWith('Removals:')) {
        currentSection = 'removals';
      } else if (trimmedLine.startsWith('Sentiment:')) {
        currentSection = 'sentiment';
        const sentimentText = trimmedLine.substring(10).trim().toLowerCase();
        if (sentimentText.includes('positive')) sentiment = 'positive';
        else if (sentimentText.includes('negative')) sentiment = 'negative';
        else sentiment = 'neutral';
      } else {
        // Content lines
        switch (currentSection) {
          case 'summary':
            if (summary) summary += ' ' + trimmedLine;
            else summary = trimmedLine;
            break;
          case 'additions':
            if (trimmedLine.match(/^[\*\-\d]/)) {
              keyAdditions.push(trimmedLine.replace(/^[\*\-\d\.\s]+/, '').trim());
            }
            break;
          case 'removals':
            if (trimmedLine.match(/^[\*\-\d]/)) {
              keyRemovals.push(trimmedLine.replace(/^[\*\-\d\.\s]+/, '').trim());
            }
            break;
        }
      }
    }
    
    return { summary, keyAdditions, keyRemovals, sentiment };
  }

  /**
   * Generate basic summary when AI is not available
   */
  private generateBasicSummary(changes: VersionChange[]): string {
    const additions = changes.filter(c => c.type === 'addition').length;
    const removals = changes.filter(c => c.type === 'removal').length;
    
    return `Document comparison shows ${additions} additions and ${removals} removals.`;
  }

  /**
   * Cache comparison with size management
   */
  private cacheComparison(key: string, comparison: VersionComparison): void {
    // Remove oldest entries if we're at capacity
    if (this.comparisons.size >= this.MAX_COMPARISON_CACHE) {
      const firstKey = this.comparisons.keys().next().value;
      if (firstKey) {
        this.comparisons.delete(firstKey);
      }
    }
    
    this.comparisons.set(key, comparison);
  }

  /**
   * Get cached comparison
   */
  getComparison(fileId: string): VersionComparison | null {
    for (const [key, comparison] of this.comparisons.entries()) {
      if (key.startsWith(fileId)) {
        return comparison;
      }
    }
    return null;
  }

  /**
   * Clear cached comparisons for a document
   */
  clearComparisons(fileId: string): boolean {
    let deleted = false;
    for (const key of this.comparisons.keys()) {
      if (key.startsWith(fileId)) {
        this.comparisons.delete(key);
        deleted = true;
      }
    }
    return deleted;
  }

  /**
   * Get comparison statistics
   */
  getComparisonStats(): {
    totalComparisons: number;
    cacheSize: number;
    averageChanges: number;
  } {
    const comparisons = Array.from(this.comparisons.values());
    const totalChanges = comparisons.reduce((sum, comp) => sum + comp.changes.length, 0);
    const averageChanges = comparisons.length > 0 ? Math.round(totalChanges / comparisons.length) : 0;
    
    return {
      totalComparisons: comparisons.length,
      cacheSize: this.comparisons.size,
      averageChanges,
    };
  }
}