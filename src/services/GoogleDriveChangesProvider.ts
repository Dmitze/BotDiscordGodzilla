import type { IDriveChangesProvider } from './DriveChangesService';
import type { GoogleService } from './GoogleService';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';
import type { drive_v3 } from 'googleapis';

/**
 * Google Drive Changes Provider implementation
 * Implements the IDriveChangesProvider interface for Google Drive API
 */
export class GoogleDriveChangesProvider implements IDriveChangesProvider {
  constructor(private readonly googleService: GoogleService) {}

  /**
   * Get the starting page token for listing future changes
   */
  async getStartPageToken(): Promise<string> {
    try {
      const drive = this.googleService.getDriveClient();
      if (!drive) {
        // If Google Drive client is not yet initialized, we'll return a default token
        // and let the service retry later
        logger.warn('Google Drive client not yet initialized, returning default token', {
          component: 'GoogleDriveChangesProvider'
        });
        // Return a default token that will be updated when the client is ready
        return '1';
      }

      const response = await drive.changes.getStartPageToken({});
      const startPageToken = response.data.startPageToken;
      
      if (!startPageToken) {
        throw new Error('Failed to get start page token from Google Drive API');
      }
      
      logger.debug('Successfully retrieved start page token', {
        component: 'GoogleDriveChangesProvider',
        startPageToken
      });
      
      return startPageToken;
    } catch (error) {
      logger.error('Failed to get start page token', {
        component: 'GoogleDriveChangesProvider',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * List changes since the given page token
   */
  async listChanges(pageToken: string): Promise<{
    changes: Array<{
      removed?: boolean;
      fileId?: string;
      file?: DriveFile;
      time?: string;
    }>;
    nextPageToken?: string;
    newStartPageToken?: string;
  }> {
    try {
      const drive = this.googleService.getDriveClient();
      if (!drive) {
        throw new Error('Google Drive client not initialized');
      }

      // Get changes from Google Drive API
      const response = await drive.changes.list({
        pageToken,
        fields: 'changes(removed,fileId,file),nextPageToken,newStartPageToken'
      });

      const changes = response.data.changes || [];
      const nextPageToken = response.data.nextPageToken ?? undefined;
      const newStartPageToken = response.data.newStartPageToken ?? undefined;

      // Map the changes to our expected format
      const mappedChanges = changes.map((change: drive_v3.Schema$Change) => {
        const result: any = {};
        
        if (change.removed !== undefined && change.removed !== null) {
          result.removed = change.removed;
        }
        if (change.fileId !== undefined && change.fileId !== null) {
          result.fileId = change.fileId;
        }
        if (change.time !== undefined && change.time !== null) {
          result.time = change.time;
        }
        
        if (change.file) {
          result.file = {
            id: change.file.id || '',
            name: change.file.name || '',
            mimeType: change.file.mimeType || '',
            parents: change.file.parents || [],
            webViewLink: change.file.webViewLink || '',
            modifiedTime: change.file.modifiedTime || '',
            owners: change.file.owners ? change.file.owners.map((owner: drive_v3.Schema$User) => ({
              emailAddress: owner.emailAddress || ''
            })) : []
          } as unknown as DriveFile;
        }
        
        return result;
      });

      logger.debug('Successfully retrieved changes', {
        component: 'GoogleDriveChangesProvider',
        changesCount: mappedChanges.length,
        hasNextPage: !!nextPageToken,
        hasNewStartToken: !!newStartPageToken
      });

      // Build the return object with proper typing
      const result: {
        changes: Array<{
          removed?: boolean;
          fileId?: string;
          file?: DriveFile;
          time?: string;
        }>;
        nextPageToken?: string;
        newStartPageToken?: string;
      } = {
        changes: mappedChanges
      };
      
      if (nextPageToken !== undefined) {
        result.nextPageToken = nextPageToken;
      }
      
      if (newStartPageToken !== undefined) {
        result.newStartPageToken = newStartPageToken;
      }
      
      return result;
    } catch (error) {
      logger.error('Failed to list changes', {
        component: 'GoogleDriveChangesProvider',
        pageToken,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }
}