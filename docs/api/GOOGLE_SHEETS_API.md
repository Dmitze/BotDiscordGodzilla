# Google Sheets API Documentation

## Overview

This document provides detailed information about the Google Sheets API integration through the `GoogleSheetsService` class. The service provides a comprehensive interface for working with Google Sheets and Google Drive files.

## Authentication

The service uses JWT (JSON Web Token) authentication with Google service account credentials. Credentials must be configured in the bot's configuration file.

### Required Scopes

- `https://www.googleapis.com/auth/spreadsheets` - Access to Google Sheets
- `https://www.googleapis.com/auth/drive` - Access to Google Drive

## Core Methods

### listSheets(spreadsheetId: string): Promise<string[]>

Retrieves a list of all sheet names in a spreadsheet.

**Parameters:**
- `spreadsheetId` (string): The ID of the Google Spreadsheet

**Returns:**
- `Promise<string[]>`: Array of sheet names

**Example:**
```typescript
const sheetNames = await googleSheetsService.listSheets('1a2b3c4d5e6f7g8h9i0j');
console.log(sheetNames); // ['Sheet1', 'Data', 'Reports']
```

### getSheetData(spreadsheetId: string, range: string, options?: GoogleServiceOptions): Promise<SheetData>

Retrieves data from a specific range in a Google Spreadsheet.

**Parameters:**
- `spreadsheetId` (string): The ID of the Google Spreadsheet
- `range` (string): The range to retrieve (e.g., 'Sheet1!A1:D10')
- `options` (GoogleServiceOptions, optional): Additional options for the request

**Returns:**
- `Promise<SheetData>`: Object containing the sheet data

**Example:**
```typescript
const data = await googleSheetsService.getSheetData(
  '1a2b3c4d5e6f7g8h9i0j', 
  'Sheet1!A1:D10',
  { useCache: true, cacheTTL: 300 }
);
console.log(data.values); // [['Name', 'Age'], ['John', '30']]
```

### writeSheetData(spreadsheetId: string, range: string, values: string[][], options?: GoogleServiceOptions): Promise<void>

Writes data to a specific range in a Google Spreadsheet.

**Parameters:**
- `spreadsheetId` (string): The ID of the Google Spreadsheet
- `range` (string): The range to write to (e.g., 'Sheet1!A1:D10')
- `values` (string[][]): 2D array of values to write
- `options` (GoogleServiceOptions, optional): Additional options for the request

**Returns:**
- `Promise<void>`: Resolves when the write operation is complete

**Example:**
```typescript
const values = [
  ['Name', 'Age', 'City'],
  ['John', '30', 'New York'],
  ['Jane', '25', 'Los Angeles']
];
await googleSheetsService.writeSheetData(
  '1a2b3c4d5e6f7g8h9i0j',
  'Sheet1!A1:C3',
  values,
  { valueInputOption: 'RAW' }
);
```

### extractTextForChat(fileId: string): Promise<{ text: string; checksum: string; modifiedTime?: string; source: 'export' | 'parser' | 'ocr' | 'raw'; warnings: string[] }>

Extracts text from Google Drive files for chat usage with validation and sanitization.

**Parameters:**
- `fileId` (string): The ID of the Google Drive file

**Returns:**
- `Promise<{ text: string; checksum: string; modifiedTime?: string; source: 'export' | 'parser' | 'ocr' | 'raw'; warnings: string[] }>`: Object containing extracted text and metadata

**Example:**
```typescript
const result = await googleSheetsService.extractTextForChat('1a2b3c4d5e6f7g8h9i0j');
console.log(result.text); // Extracted text content
console.log(result.source); // 'export' | 'parser' | 'ocr' | 'raw'
```

### searchData(query: string, limit: number = 20): Promise<string[][]>

Searches data across spreadsheets.

**Parameters:**
- `query` (string): The search query
- `limit` (number, optional): Maximum number of results (default: 20)

**Returns:**
- `Promise<string[][]>`: 2D array of search results

**Example:**
```typescript
const results = await googleSheetsService.searchData('personnel report', 10);
console.log(results); // Search results as 2D array
```

### readRange(fileId: string, sheetName: string, rangeOrOpts: string | { columnHints?: string[]; headerRow?: number }): Promise<{ headers: string[]; rows: (string | number | null)[][] }>

Reads a specific range of data with normalization.

**Parameters:**
- `fileId` (string): The ID of the Google Spreadsheet
- `sheetName` (string): The name of the sheet to read from
- `rangeOrOpts` (string | { columnHints?: string[]; headerRow?: number }): Range string or options object

**Returns:**
- `Promise<{ headers: string[]; rows: (string | number | null)[][] }>`: Object containing headers and data rows

**Example:**
```typescript
const data = await googleSheetsService.readRange(
  '1a2b3c4d5e6f7g8h9i0j',
  'Sheet1',
  'A1:D10'
);
console.log(data.headers); // ['name', 'age', 'city']
console.log(data.rows); // [['John', 30, 'New York'], ...]
```

### findSheetByName(fileId: string, name: string): Promise<{ title: string; index: number } | null>

Finds a sheet by name (case-insensitive).

**Parameters:**
- `fileId` (string): The ID of the Google Spreadsheet
- `name` (string): The name of the sheet to find

**Returns:**
- `Promise<{ title: string; index: number } | null>`: Object with sheet title and index, or null if not found

**Example:**
```typescript
const sheet = await googleSheetsService.findSheetByName('1a2b3c4d5e6f7g8h9i0j', 'Data');
if (sheet) {
  console.log(`Found sheet: ${sheet.title} at index ${sheet.index}`);
}
```

## Utility Methods

### getDriveFileMetadata(fileId: string): Promise<drive_v3.Schema$File>

Retrieves metadata for a Google Drive file.

**Parameters:**
- `fileId` (string): The ID of the Google Drive file

**Returns:**
- `Promise<drive_v3.Schema$File>`: File metadata object

**Example:**
```typescript
const metadata = await googleSheetsService.getDriveFileMetadata('1a2b3c4d5e6f7g8h9i0j');
console.log(metadata.name); // File name
console.log(metadata.mimeType); // File MIME type
```

### downloadDriveFile(fileId: string): Promise<Buffer>

Downloads a binary file from Google Drive.

**Parameters:**
- `fileId` (string): The ID of the Google Drive file

**Returns:**
- `Promise<Buffer>`: Buffer containing the file data

**Example:**
```typescript
const buffer = await googleSheetsService.downloadDriveFile('1a2b3c4d5e6f7g8h9i0j');
// Use buffer for file processing
```

### exportDriveFile(fileId: string, mimeType: string): Promise<Buffer>

Exports a Google Docs/Sheets/Slides file to a specified MIME type.

**Parameters:**
- `fileId` (string): The ID of the Google Drive file
- `mimeType` (string): The target MIME type for export

**Returns:**
- `Promise<Buffer>`: Buffer containing the exported file data

**Supported MIME Types:**
- `text/csv` - CSV format
- `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` - Excel XLSX format
- `text/plain` - Plain text
- `application/pdf` - PDF format

**Example:**
```typescript
// Export as CSV
const csvBuffer = await googleSheetsService.exportDriveFile('1a2b3c4d5e6f7g8h9i0j', 'text/csv');

// Export as XLSX
const xlsxBuffer = await googleSheetsService.exportDriveFile('1a2b3c4d5e6f7g8h9i0j', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
```

## Error Handling

All methods throw errors that should be properly handled:

```typescript
try {
  const data = await googleSheetsService.getSheetData('invalid-id', 'A1:B2');
} catch (error) {
  if (error instanceof Error) {
    console.error('Google Sheets API Error:', error.message);
    // Handle specific error cases
  }
}
```

## Rate Limiting

The service implements automatic rate limiting to comply with Google API quotas:
- Default: 5 requests per second
- Burst: 10 requests
- Configurable through `drive.rateQps` and `drive.rateBurst` settings

## Caching

Built-in caching is available for improved performance:
- Default TTL: 300 seconds (5 minutes)
- Configurable per request
- Automatic cache invalidation

## Data Types

### SheetData
```typescript
interface SheetData {
  range: string;
  majorDimension: string;
  values: string[][];
}
```

### GoogleServiceOptions
```typescript
interface GoogleServiceOptions {
  useCache?: boolean;
  cacheTTL?: number;
  forceRefresh?: boolean;
  batchSize?: number;
  retryFailed?: boolean;
  cacheResults?: boolean;
  maxRetries?: number;
  valueInputOption?: 'RAW' | 'USER_ENTERED';
  clearCache?: boolean;
}
```

## Best Practices

1. **Always Initialize**: Ensure the service is properly initialized before use
2. **Handle Errors**: Implement proper error handling for all operations
3. **Use Caching**: Take advantage of built-in caching for repeated operations
4. **Clean Shutdown**: Always call shutdown() to properly close connections
5. **Validate Inputs**: Validate all inputs before passing to service methods
6. **Respect Rate Limits**: Be mindful of Google API rate limits in high-volume operations

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Verify service account credentials and permissions
2. **Rate Limiting**: Implement exponential backoff for high-volume operations
3. **File Access**: Ensure proper sharing permissions for Google Drive files
4. **Invalid Range**: Check that range strings are properly formatted

### Debugging

Enable debug logging to troubleshoot issues:

```typescript
// Set environment variable
process.env.LOG_LEVEL = 'debug';
```

Logs will include detailed information about API calls, errors, and performance metrics.