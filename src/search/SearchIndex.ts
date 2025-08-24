export interface SearchFilters {
  fileId?: string[];
  mime?: string[];
  owner?: string[];
  modifiedFrom?: number; // epoch ms
  modifiedTo?: number;   // epoch ms
  sizeFrom?: number;
  sizeTo?: number;
  tags?: string[];
}

export interface SearchQuery {
  text: string;
  filters?: SearchFilters;
  limit?: number;   // default 20
  offset?: number;  // default 0
  changesOnly?: boolean;
}

export interface SearchHit {
  fileId: string;
  name: string;
  mimeType?: string;
  ownerEmail?: string;
  modifiedTime?: number;
  snippet?: string;
  contentHash: string;
  textLen: number;
  score?: number;
  changedMeta?: string[];
}

export interface SearchIndex {
  upsert(doc: {
    fileId: string;
    name: string;
    mimeType?: string;
    ownerEmail?: string;
    sizeBytes?: number;
    modifiedTime?: number;
    createdTime?: number;
    text: string;            // normalized text
    tags?: string[];
    meta?: unknown;
  }): Promise<void>;

  search(q: SearchQuery): Promise<{ hits: SearchHit[]; total: number }>;

  getDiff(fileId: string): Promise<{
    latestHash: string;
    previousHash?: string;
    diffMeta?: Record<string, { old?: unknown; new?: unknown }>;
  }>;
}
