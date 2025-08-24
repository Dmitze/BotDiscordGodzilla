import type { ChatInputCommandInteraction } from 'discord.js';

export interface SearchFilters {
  documentType: string;
  dateFrom?: string;
  dateTo?: string;
  unit?: string;
  priority: string;
  limit: number;
}

export interface SearchResult {
  rows: string[][];
  headers: string[];
  totalCount: number;
  filteredCount: number;
  searchTime: number;
  cacheHit: boolean;
  query: string;
  filters: SearchFilters;
}

export interface PaginationState {
  currentPage: number;
  totalPages: number;
  results: SearchResult;
  timestamp: number;
  userId: string;
  pageSize: number;
  changesOnly: boolean;
}

export type InteractionLike = ChatInputCommandInteraction | { [k: string]: any };
