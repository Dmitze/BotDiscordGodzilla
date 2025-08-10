# n8n Workflows for Discord Bot

This directory contains n8n workflows that integrate with the Discord bot for automated document processing and monitoring.

## Setup

1. Make sure you have Docker and Docker Compose installed
2. Copy `.env.n8n` to `.env` and adjust the values as needed:
   ```bash
   cp .env.n8n .env
   ```
3. Start the services:
   ```bash
   docker-compose up -d
   ```

## Workflows

### Google Drive Document Monitor

**File:** `google-drive-monitor.json`

This workflow monitors a Google Drive folder for new documents and sends notifications to the Discord bot when new files are added.

#### Setup Instructions:

1. Import the workflow into n8n
2. Configure the Google Drive Trigger node with:
   - Your Google Drive credentials
   - The folder ID to monitor
3. Configure the "Send to Discord Bot" node with:
   - The webhook URL of your Discord bot (typically `http://your-bot-server:3000/webhook/n8n/drive`)
   - Authentication if required

### Document Processing Workflow

**File:** `document-processing-workflow.json`

This workflow processes documents from Google Drive, extracts text, splits it into chunks, creates embeddings, and sends the processed data to the Discord bot for RAG (Retrieval-Augmented Generation).

### Automatic Document Indexing Pipeline

**File:** `automatic-indexing-pipeline.json`

This enhanced workflow provides a complete automatic indexing pipeline that monitors Google Drive for document changes, processes them through text extraction, chunking, and embedding creation, and sends the results to the Discord bot for indexing and RAG capabilities. It includes improved error handling and status reporting.

#### Setup Instructions:

1. Import the workflow into n8n
2. Configure the Google Drive Trigger node with:
   - Your Google Drive credentials
   - The folder ID to monitor
3. Configure the "Send to Discord Bot" nodes with:
   - The webhook URL of your Discord bot (typically `http://your-bot-server:3000/webhook/n8n/drive`)
   - Authentication if required

## Webhook Endpoint

The Discord bot exposes a webhook endpoint at `/webhook/n8n/drive` that receives file update notifications from n8n.

### Endpoint Details:

- **URL:** `/webhook/n8n/drive`
- **Method:** POST
- **Authentication:** None (in current implementation)
- **Payload:**
  ```json
  {
    "fileId": "Google Drive file ID",
    "fileName": "Name of the file",
    "fileContent": "Base64 encoded file content (optional)",
    "mimeType": "MIME type of the file (optional)",
    "chunks": "Array of text chunks (optional)",
    "embeddings": "Array of embeddings (optional)"
  }
  ```

### Response:

```json
{
  "success": true,
  "message": "File update received and processed",
  "fileId": "Google Drive file ID",
  "fileName": "Name of the file"
}
```

## Testing

To test the workflow:

1. Start the n8n service
2. Upload a document to the monitored Google Drive folder
3. Check the Discord bot logs for processing messages
4. Verify that the file is indexed and searchable through the bot

## Troubleshooting

- Ensure the Google Drive credentials are properly configured
- Check that the folder ID is correct
- Verify the webhook URL is accessible from the n8n container
- Check the Discord bot logs for any errors during webhook processing