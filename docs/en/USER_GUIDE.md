# 📚 User Guide
# Discord AI Assistant Bot User Guide

## 🎯 Introduction

This guide will help you effectively use the Discord AI Assistant Bot to automate document work, analyze data, and support operational activities.

## 🔐 Roles and Access Rights

The bot uses a role-based system to control access:

- **Administrator** - full access to all functions
- **Bot User** - basic commands
- **Sheets Access** - working with Google Sheets
- **AI Access** - AI functions
- **Export Access** - data export

Contact your server administrator to obtain the necessary roles.

## 🔍 Search Commands

### Basic Search

The `/search` command allows you to search for information in documents:

```
/search query:"keyword" document_type:"orders"
```

**Parameters:**
- `query` - text to search for (required)
- `document_type` - document type filter
- `date_from` - search from a specific date
- `date_to` - search up to a specific date
- `unit` - unit filter
- `priority` - document priority filter
- `limit` - number of results

### Smart Search

The `/smart-search` command allows you to perform more complex search queries:

```
/smart-search quantity_above:100 price_below:1000
```

**Parameters:**
- `quantity_above` - minimum quantity
- `quantity_below` - maximum quantity
- `price_above` - minimum price
- `price_below` - maximum price

## 🤖 AI Commands

### AI Assistant

The `/ai` command allows you to interact with the AI assistant:

```
/ai query:"analyze personnel"
```

**Parameters:**
- `query` - your query to AI (required)
- `context` - additional context for AI

### AI Usage Examples

1. **Data Analysis:**
   ```
   /ai query:"analyze sales for the last month"
   ```

2. **Report Generation:**
   ```
   /ai query:"create a report on warehouse stock"
   ```

3. **Recommendations:**
   ```
   /ai query:"how to optimize logistics?"
   ```

## 📄 Working with Documents

### /documents Command

Allows you to work with documents:

```
/documents personnel list
```

**Subcommands:**
- `add` - add a new document
- `list` - show list of documents
- `update` - update an existing document
- `delete` - delete a document

### /files Command

Allows you to work with Google Drive files:

```
/files search query:"file name"
```

**Subcommands:**
- `search` - search for files
- `read` - read file content
- `analyze` - AI analyze file
- `report` - create a report from file

## 📊 Analytics and Statistics

### /statistics Command

Shows bot usage statistics:

```
/statistics
```

### /analytics Command

Generates analytical reports:

```
/analytics report type:"general"
```

**Report Types:**
- `general` - general analytics
- `search` - search analytics
- `commands` - command analytics
- `activity` - activity analytics

## ⚡ Operational Commands

### /operations Command

operational processes management:

```
/operations situation sector:"A"
```

**Subcommands:**
- `situation` - get situation information
- `task` - task management
- `coordination` - unit coordination
- `intelligence` - intelligence data

## 🛠️ Administrative Commands

### /status Command

Shows system status:

```
/status
```

### /cache Command

Cache management:

```
/cache clear type:"all"
```

**Cache Types:**
- `search` - search cache
- `ai` - AI cache
- `all` - all cache types

### /help Command

Shows help:

```
/help category:"search"
```

**Categories:**
- `search` - search commands
- `ai` - AI commands
- `documents` - working with documents
- `files` - working with files
- `analytics` - analytical commands
- `operations` - operational commands

## 🔒 Security

### Rate Limiting

The bot has a request limiting system to prevent abuse:

- Search: 10 requests per minute
- AI analysis: 5 requests per 2 minutes
- Export: 3 requests per 5 minutes

### Data Validation

All input data is automatically checked for:
- Length (max. 500 characters for text)
- Presence of forbidden characters
- Date format

## 📤 Data Export

Most commands support result export:

1. Execute a search or analysis command
2. Click the "Export" button in the response
3. Select export format (Excel, CSV, PDF, DOCX)
4. Download the file

## ❓ Frequently Asked Questions

### How to get access to the bot?

Contact your Discord server administrator to obtain the necessary roles.

### What file formats are supported?

The bot supports working with:
- Google Sheets
- Google Docs
- PDF files
- Microsoft Word documents
- Text files

### Is AI usage safe?

Yes, the bot ensures maximum security:
- Data is not transmitted to third parties (when using Ollama)
- All data is masked in logs
- Access control through roles

## 📞 Support

If you encounter problems:

1. Check this guide
2. Execute the `/help` command
3. Contact your server administrator
4. Create an issue in the GitHub repository

© 2025 Dmitry Shivachov (Dmitze). All rights reserved.