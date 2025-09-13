# 🔧 Godzilla Bot Setup

## 📋 Required Environment Variables

Create a `.env` file in the project root using `.env.example` as a base:

## 🔑 Discord Bot Setup

### 1. Obtaining Discord Token

1. Go to Discord Developer Portal
2. Click "New Application" and enter a name
3. In the "Bot" section, create a new bot
4. Copy the token (click "Copy" next to "Token")
5. Enable required Intents:
   ✅ Message Content Intent
   ✅ Server Members Intent
   ✅ Presence Intent
   ✅ Message Content

### 2. Adding Bot to Server

1. In the "OAuth2" "URL Generator" section, select:
   Permissions: Send Messages, Use Slash Commands, Embed Links, Attach Files, Read Message History
2. Generate URL and ask administrator to add bot to server

## ⚙️ Database Setup

1. SQLite3 installs automatically
2. Database will be created at `data/database.sqlite`
3. For migrations, use:

## 🔑 Obtaining API Keys

### OpenAI API Key

1. Go to OpenAI Platform
2. Log in or create an account
3. Go to "API Keys" section
4. Click "Create new secret key"
5. Copy the key

### Google Sheets Setup

#### 1. Creating a Table

1. Create a new Google Sheet
2. Add headers in the first row (for example):
   Product Name, Serial Number, Counterparty, Quantity, Price, Cost

#### 2. Getting Table ID

1. Open your Google Sheet
2. Copy ID from URL:

#### 3. Setting Access

1. Click "Share" in the upper right corner
2. Add your Google API key as editor
3. Or set public access (read-only only)

## 🤖 Discord Bot Setup

### 1. Adding Bot to Server

1. In Discord Developer Portal, go to "OAuth2" "URL Generator" section
2. Select scopes: bot, applications.commands
3. Select permissions: Send Messages, Use Slash Commands, Embed Links, Attach Files, Read Message History
4. Copy the generated URL
5. Open URL in browser and add bot to your server

### 2. Command Setup

Bot automatically registers slash commands on startup.
If you need to update commands:

## 🚀 Bot Launch

### 1. Installing Dependencies

### 2. Launch

Or use PowerShell script:

## 🔍 Testing

Main commands for testing:

1. `/help` - check bot operation
2. `/summary` - check Google Sheets connection
3. `/ai-analyze` - check AI functionality

### Log Checking

Check files in `logs/` folder for diagnostics.

## ❗ Common Issues

### 1. "Invalid token"

- Check Discord Bot Token correctness
- Ensure token doesn't contain extra characters

### 2. "Google Sheets API error"

- Check Google API Key correctness
- Enable Google Sheets API in Google Cloud Console
- Check table access

### 3. "OpenAI API error"

- Check OpenAI API Key correctness
- Ensure you have credits on OpenAI account

### 4. "Command not found"

- Restart bot after changing commands
- Check that bot has rights to use slash commands

## 📞 Support

If you encounter issues:

1. Check logs in `logs/` folder
2. Ensure all environment variables are configured correctly
3. Check internet connection
4. Create Issue in repository with detailed problem description