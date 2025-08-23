export interface PaginationInit {
  filteredCount: number;
  limit: number;
}

export function computePagination({ filteredCount, limit }: PaginationInit) {
  const pageSize = Math.max(1, limit || 10);
  const totalPages = Math.max(1, Math.ceil(filteredCount / pageSize));
  return { pageSize, totalPages };
}
