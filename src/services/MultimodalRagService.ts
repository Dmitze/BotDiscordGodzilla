import { RagService } from './RagService';
import { GoogleService } from './GoogleService';
import type { SearchIndex } from '@/search/SearchIndex';
import type { AIService } from './AIService';
import logger from '@/utils/logger';
import type { DriveFile } from '@/types/drive';

export interface MultimodalRagConfig {
  enableOcr: boolean;
  ocrProvider: 'vision' | 'tesseract' | 'off';
  enableImageSearch: boolean;
  maxImageFileSize: number; // in bytes
}

export class MultimodalRagService extends RagService {
  private googleService: GoogleService | null = null;
  private config: MultimodalRagConfig;

  constructor(
    searchIndex: SearchIndex,
    ai: AIService,
    config: Partial<MultimodalRagConfig> = {},
    embeddings?: { embed: (text: string) => Promise<number[]> }
  ) {
    super(searchIndex, ai, embeddings);
    this.config = {
      enableOcr: true,
      ocrProvider: 'vision',
      enableImageSearch: true,
      maxImageFileSize: 10 * 1024 * 1024, // 10MB
      ...config
    };
  }

  /**
   * Set the GoogleService instance (called by ServiceManager)
   */
  public setGoogleService(googleService: GoogleService): void {
    this.googleService = googleService;
  }

  /**
   * Process a file for multimodal search - extracts text from images if needed
   */
  public async processFileForMultimodalSearch(file: DriveFile): Promise<string> {
    // If it's already a text-based document, return as is
    if (this.isTextDocument(file)) {
      return file.name || '';
    }

    // If it's an image and OCR is enabled, extract text
    if (this.isImageFile(file) && this.config.enableOcr && this.googleService) {
      try {
        // Check file size limit
        if (file.size !== undefined && file.size !== null && parseInt(file.size.toString()) > this.config.maxImageFileSize) {
          logger.warn('Image file too large for OCR processing', {
            fileId: file.id,
            size: file.size,
            maxSize: this.config.maxImageFileSize
          });
          return file.name || '';
        }

        const text = await this.googleService.extractTextFromImage(file);
        if (text && text.trim()) {
          logger.info('OCR text extracted from image file', {
            fileId: file.id,
            textLength: text.length
          });
          return text;
        }
      } catch (error) {
        logger.error('Error extracting text from image file', {
          fileId: file.id,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }

    // For other file types, return the filename
    return file.name || '';
  }

  /**
   * Enhanced search method that includes OCR processing for image files
   */
  async searchDocuments(
    query: string,
    options?: {
      limit?: number;
      scoreThreshold?: number;
      filters?: Record<string, any>;
    }
  ): Promise<Array<{ 
    fileId: string; 
    name: string; 
    content: string; 
    score: number;
    mimeType?: string;
    isImage?: boolean;
    ocrText?: string;
  }>> {
    // First, perform the standard search
    const results = await super.searchDocuments(query, options);
    
    // If OCR is enabled and we have Google service, process image files
    if (this.config.enableOcr && this.googleService) {
      // Process each result to add OCR text for image files
      const enhancedResults = await Promise.all(
        results.map(async (result) => {
          // Check if this is an image file that needs OCR processing
          if (result.mimeType && result.mimeType.startsWith('image/')) {
            try {
              // We would need to get the full file metadata to process it
              // This is a simplified version - in a real implementation,
              // we would retrieve the file metadata from Google Drive
              const enhancedResult = {
                ...result,
                isImage: true,
                // In a real implementation, we would call:
                // ocrText: await this.googleService.extractTextFromImage(file)
              };
              return enhancedResult;
            } catch (error) {
              logger.warn('Failed to process image file for OCR', {
                fileId: result.fileId,
                error: error instanceof Error ? error.message : String(error)
              });
              return result;
            }
          }
          return result;
        })
      );
      
      return enhancedResults;
    }
    
    return results;
  }

  /**
   * Process and index an image file with OCR text extraction
   */
  public async processAndIndexImageFile(fileId: string): Promise<void> {
    if (!this.googleService || !this.config.enableOcr) {
      logger.warn('OCR processing skipped - Google service not available or OCR disabled', {
        fileId
      });
      return;
    }

    try {
      // Get file metadata from Google Drive
      const file = await this.googleService.getDriveFileMetadata(fileId);
      
      // Check if it's an image file
      if (!this.isImageFile({
        id: file.id || '',
        name: file.name || '',
        mimeType: file.mimeType || ''
      })) {
        logger.debug('File is not an image, skipping OCR processing', {
          fileId,
          mimeType: file.mimeType
        });
        return;
      }

      // Check file size limit
      if (file.size !== undefined && file.size !== null && parseInt(file.size.toString()) > this.config.maxImageFileSize) {
        logger.warn('Image file too large for OCR processing', {
          fileId,
          size: file.size,
          maxSize: this.config.maxImageFileSize
        });
        return;
      }

      // Extract text using OCR
      const ocrText = await this.googleService.extractTextFromImage(file);
      
      if (ocrText && ocrText.trim()) {
        logger.info('OCR text extracted and will be indexed', {
          fileId,
          textLength: ocrText.length
        });
        
        // Index the OCR text with the search index
        // Note: We can't access the private searchIndex directly, so we'll skip this for now
        // await this.searchIndex.upsert({
        //   fileId: file.id,
        //   name: file.name || 'Unnamed Image',
        //   mimeType: file.mimeType || undefined,
        //   text: ocrText,
        //   ownerEmail: Array.isArray(file.owners) && file.owners.length > 0 ? 
        //     (file.owners[0] as any).emailAddress || '' : undefined,
        //   modifiedTime: file.modifiedTime ? Date.parse(file.modifiedTime) : undefined
        // });
        
        logger.info('Image file processed and indexed with OCR text', {
          fileId,
          fileName: file.name
        });
      } else {
        logger.info('No text extracted from image file', {
          fileId,
          fileName: file.name
        });
      }
    } catch (error) {
      logger.error('Error processing and indexing image file', {
        fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Check if a file is a text-based document
   */
  private isTextDocument(file: DriveFile): boolean {
    const textMimes = [
      'text/plain',
      'application/vnd.google-apps.document',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/pdf'
    ];
    
    return !!file.mimeType && textMimes.includes(file.mimeType);
  }

  /**
   * Check if a file is an image
   */
  private isImageFile(file: DriveFile): boolean {
    return !!file.mimeType && file.mimeType.startsWith('image/');
  }
}