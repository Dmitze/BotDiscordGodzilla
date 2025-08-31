# GoogleSheetsService Documentation

## Overview

The `GoogleSheetsService` is an enhanced service that provides comprehensive integration with Google Sheets and Google Drive APIs. It replaces the previous `GoogleService` implementation and leverages the `google-spreadsheet` library for improved functionality and performance.

## Features

- **Full Google Sheets API Integration**: Complete support for reading, writing, and managing Google Sheets
- **Google Drive Integration**: File management and metadata operations
- **Enhanced Performance**: Optimized with caching and connection pooling
- **Robust Error Handling**: Comprehensive error handling with detailed logging
- **Type Safety**: Strong TypeScript typing for all operations
- **Backward Compatibility**: Maintains compatibility with existing GoogleService interface

## Installation

The service uses the `google-spreadsheet` library which is already included in the project dependencies:

```json
"google-spreadsheet": "^5.0.2"
```

## Configuration

The service requires Google service account credentials configured in the bot's configuration:

```typescript
const config: BotConfig = {
  google: {
    spreadsheetId: 'your-spreadsheet-id',
    driveFolderId: 'your-drive-folder-id',
    credentials: {
      client_email: 'your-service-account-email@project.iam.gserviceaccount.com',
      private_key: '-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n',
      project_id: 'your-project-id',
    },
  },
  // ... other configuration
};
```

## API Reference

### Initialization

```typescript
const googleSheetsService = new GoogleSheetsService(config);
await googleSheetsService.initialize();
```

### Core Methods

#### listSheets(spreadsheetId: string): Promise<string[]>

Lists all sheet names in a spreadsheet.

```typescript
const sheetNames = await googleSheetsService.listSheets('spreadsheet-id');
console.log(sheetNames); // ['Sheet1', 'Sheet2', 'Data']
```

#### getSheetData(spreadsheetId: string, range: string, options?: GoogleServiceOptions): Promise<SheetData>

Retrieves data from a specific range in a spreadsheet.

```typescript
const data = await googleSheetsService.getSheetData('spreadsheet-id', 'Sheet1!A1:D10');
console.log(data.values); // [['Header1', 'Header2'], ['Value1', 'Value2']]
```

#### writeSheetData(spreadsheetId: string, range: string, values: string[][], options?: GoogleServiceOptions): Promise<void>

Writes data to a spreadsheet.

```typescript
const values = [
  ['Name', 'Age', 'City'],
  ['John', '30', 'New York'],
  ['Jane', '25', 'Los Angeles']
];
await googleSheetsService.writeSheetData('spreadsheet-id', 'Sheet1!A1:C3', values);
```

#### extractTextForChat(fileId: string): Promise<{ text: string; checksum: string; modifiedTime?: string; source: 'export' | 'parser' | 'ocr' | 'raw'; warnings: string[] }>

Extracts text from Google Drive files for chat usage with validation and sanitization.

```typescript
const result = await googleSheetsService.extractTextForChat('file-id');
console.log(result.text); // Extracted text content
```

#### searchData(query: string, limit: number = 20): Promise<string[][]>

Searches data across spreadsheets.

```typescript
const results = await googleSheetsService.searchData('search term', 10);
console.log(results); // Search results
```

#### readRange(fileId: string, sheetName: string, rangeOrOpts: string | { columnHints?: string[]; headerRow?: number }): Promise<{ headers: string[]; rows: (string | number | null)[][] }>

Reads a specific range of data with normalization.

```typescript
const data = await googleSheetsService.readRange('spreadsheet-id', 'Sheet1', 'A1:D10');
console.log(data.headers); // Normalized headers
console.log(data.rows); // Data rows
```

#### findSheetByName(fileId: string, name: string): Promise<{ title: string; index: number } | null>

Finds a sheet by name (case-insensitive).

```typescript
const sheet = await googleSheetsService.findSheetByName('spreadsheet-id', 'Sheet1');
if (sheet) {
  console.log(`Found sheet: ${sheet.title} at index ${sheet.index}`);
}
```

### Utility Methods

#### getDriveFileMetadata(fileId: string): Promise<drive_v3.Schema$File>

Retrieves metadata for a Google Drive file.

```typescript
const metadata = await googleSheetsService.getDriveFileMetadata('file-id');
console.log(metadata.name); // File name
```

#### downloadDriveFile(fileId: string): Promise<Buffer>

Downloads a binary file from Google Drive.

```typescript
const buffer = await googleSheetsService.downloadDriveFile('file-id');
// Use buffer for file processing
```

#### exportDriveFile(fileId: string, mimeType: string): Promise<Buffer>

Exports a Google Docs/Sheets/Slides file to a specified MIME type.

```typescript
const csvBuffer = await googleSheetsService.exportDriveFile('file-id', 'text/csv');
const xlsxBuffer = await googleSheetsService.exportDriveFile('file-id', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
```

## Usage Examples

### Basic Spreadsheet Operations

```typescript
import { GoogleSheetsService } from '../services/GoogleSheetsService';
import type { BotConfig } from '../types';

// Initialize service
const service = new GoogleSheetsService(config);
await service.initialize();

// List sheets
const sheets = await service.listSheets('spreadsheet-id');
console.log('Available sheets:', sheets);

// Read data
const data = await service.getSheetData('spreadsheet-id', 'Sheet1!A1:C10');
console.log('Data:', data.values);

// Write data
const newValues = [
  ['Name', 'Email', 'Department'],
  ['John Doe', 'john@example.com', 'Engineering']
];
await service.writeSheetData('spreadsheet-id', 'Sheet1!A1:C2', newValues);

// Shutdown
await service.shutdown();
```

### File Processing

```typescript
// Extract text from various file types
const textResult = await service.extractTextForChat('document-id');
console.log('Extracted text:', textResult.text);

// Download and process file
const fileBuffer = await service.downloadDriveFile('file-id');
// Process buffer as needed
```

## Error Handling

The service provides comprehensive error handling with detailed logging:

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

## Performance Optimization

The service includes several performance optimizations:

1. **Caching**: Built-in caching for frequently accessed data
2. **Connection Pooling**: Efficient management of API connections
3. **Rate Limiting**: Automatic handling of Google API rate limits
4. **Batch Operations**: Support for batch data operations

## Testing

The service includes comprehensive test coverage:

```bash
# Run GoogleSheetsService tests
npm test src/services/__tests__/GoogleSheetsService.test.ts

# Run integration tests
npm test src/services/__tests__/GoogleSheetsService.integration.test.ts
```

## Migration from GoogleService

The GoogleSheetsService maintains backward compatibility with the GoogleService interface. Existing code should work without modifications, but you can take advantage of the enhanced features:

1. Replace imports from `GoogleService` to `GoogleSheetsService`
2. Update service registration in ServiceManager if needed
3. Take advantage of new methods like `findSheetByName` and enhanced `readRange`

## Security

The service implements several security measures:

1. **Input Validation**: All inputs are validated and sanitized
2. **Authentication**: Secure JWT authentication with Google APIs
3. **Rate Limiting**: Automatic rate limiting to prevent abuse
4. **Error Redaction**: Sensitive information is redacted in logs

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Verify service account credentials and permissions
2. **Rate Limiting**: Implement exponential backoff for high-volume operations
3. **File Access**: Ensure proper sharing permissions for Google Drive files

### Debugging

Enable debug logging to troubleshoot issues:

```typescript
// Set environment variable
process.env.LOG_LEVEL = 'debug';
```

## Best Practices

1. **Always Initialize**: Ensure the service is properly initialized before use
2. **Handle Errors**: Implement proper error handling for all operations
3. **Use Caching**: Take advantage of built-in caching for repeated operations
4. **Clean Shutdown**: Always call shutdown() to properly close connections
5. **Validate Inputs**: Validate all inputs before passing to service methods

## API Limits

The service automatically handles Google API limits:
- Sheets API: 5 requests per second by default
- Drive API: 10 requests per second by default
- Customizable through configuration

## Contributing

To contribute to the GoogleSheetsService:

1. Follow the existing code style and patterns
2. Add comprehensive tests for new functionality
3. Update documentation when adding new features
4. Ensure backward compatibility is maintained