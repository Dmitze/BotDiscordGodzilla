# 📚 Discord AI Assistant Bot Usage Guide

## 🎯 Overview

Discord AI Assistant Bot is a powerful tool for working with Google Sheets, Google Drive, and AI data analysis directly in Discord. The bot supports natural language queries, file operations, and report generation.

## 🚀 Quick Start

### 1. Role Setup

Create the following roles on your Discord server:
- Administrator - full access to all functions
- Bot User - basic access
- Sheets Access - access to Google Sheets
- AI Access - AI functions access
- Export Access - data export access

### 2. First Commands

## 🔍 Search Commands

### Field Search

Description: Search data by specific table field

Parameters:
- field to search (name, serial number, counterparty, quantity, price)
- what to search for

Examples:

### Smart Search

Description: Search by multiple criteria simultaneously

Parameters:
- search by product name
- search by counterparty
- search by serial number
- products more expensive than specified price
- products with quantity more than specified

Examples:

### Summary Values

Description: Shows summary values from table

Example:

### Recent Records

Description: Shows last 10 records from table

Example:

## 🤖 AI Functions

### AI Assistant

Description: Natural language query to AI for data work

Parameters:
- what you want to do
- additional context (optional)

Query Examples:

Examples with context:

## 📁 File Operations

### Google Drive Work

Description: Search, read, and analyze files in Google Drive

Parameters:
- what to do (search, read, analyze, report)
- file name or ID
- folder ID for search (optional)
- analysis type (only for "analyze" action)
- report format (only for "report" action)

File Search

File Reading

AI File Analysis

Report Generation

## 📤 Data Export

### Search Results Export

Description: Exports search results to Excel format

Parameters:
- field to search
- what to search for

Examples:

## ⚙️ Administrative Commands

### Bot Statistics

Description: Shows bot usage statistics

Example:

### Cache Clear

Description: Clears user cache

Parameters:
- cache type to clear (search, ai, all)

Examples:

### Help

Description: Shows command help

Parameters:
- command category (basic, search, ai, files, admin)

Examples:

## 🎯 Usage Examples

### Scenario 1: Stock Analysis

### Scenario 2: Document Work

### Scenario 3: Sales Analysis

## 🔧 Setup

### Environment Variables

Main variables for setup:

### Roles and Access Rights

- Administrator - full access
- Bot User - basic commands
- Sheets Access - Google Sheets work
- AI Access - AI functions
- Export Access - data export

## 🚨 Security

### Rate Limiting

Bot has built-in request limiting system:
- Search: 10 requests per minute
- AI analysis: 5 requests per 2 minutes
- Export: 3 requests per 5 minutes

### Input Data Validation

All input data is automatically cleaned
Request length limits
XSS attack protection

### Logging

All user actions are logged for security:
- Commands
- File access
- AI queries
- Errors

## 🆘 Troubleshooting

### Common Issues

Bot not responding:
1. Check if bot is online
2. Check access rights
3. Check logs for errors

Google API errors:
1. Check Google setup
2. Check file access rights
3. Check API quotas

AI not working:
1. Check OpenAI/Ollama setup
2. Check API key balance
3. Check internet connection

Diagnostic Commands

## 📞 Support

If you encounter issues:
1. Check this guide
2. View bot logs
3. Try command
4. Contact administrator

## 🔄 Updates

To update bot:
1. Stop bot
2. Update code
3. Install new dependencies:
4. Start bot:

Last update: Version 2.3.0