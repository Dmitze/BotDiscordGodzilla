# 🛡️ Security
# Discord AI Assistant Bot Security Guide

## 🔐 General Security Principles

Discord AI Assistant Bot implements a multi-level security system to protect data and prevent abuse:

### 1. Access Control
- Discord role system
- Rights validation for each command
- Access control to Google documents

### 2. Encryption
- All data is transmitted encrypted
- API keys are stored in environment variables
- Confidential data is masked in logs

### 3. Monitoring
- Logging of all operations
- Alerts for suspicious activity
- User action audit

## 🛡️ Roles and Access Rights

### Discord Role System

The bot uses the following roles to control access:

1. **Administrator** - full access to all functions
2. **Bot User** - basic commands
3. **Sheets Access** - working with Google Sheets
4. **AI Access** - AI functions
5. **Export Access** - data export

### Role Setup

The server administrator must:

1. Create appropriate roles in Discord
2. Configure channel access rights
3. Assign roles to users

## 🔒 Data Protection

### Confidential Information

The bot automatically masks confidential information:

- **Email addresses**: `user***@***.com`
- **Phone numbers**: `+380 *** **** 123`
- **API keys**: not stored in logs

### Encryption

- All connections use HTTPS/TLS
- API keys are stored only in environment variables
- Data in cache (Redis) can be encrypted

## 🚦 Rate Limiting

The bot implements a request limiting system to prevent abuse:

### Request Limits

- **Search**: 10 requests per minute
- **AI analysis**: 5 requests per 2 minutes
- **Export**: 3 requests per 5 minutes
- **General commands**: 20 requests per minute

### Limit Bypass

When limits are exceeded:
1. User receives a limit notification
2. Further requests are ignored until the period ends
3. Administrator receives an alert about repeated attempts

## 🧪 Data Validation

### Input Validation

All input data undergoes validation:

1. **Length**:
   - Search queries: max. 200 characters
   - AI queries: max. 1000 characters
   - File names: max. 100 characters
   - General text: max. 500 characters

2. **Format**:
   - Dates: DD.MM.YYYY format
   - Numbers: only digits and decimal separators
   - Email: format validation

3. **Forbidden Characters**:
   - HTML tags
   - SQL injections
   - System command special characters

## 📊 Security Monitoring

### Logging

The bot maintains detailed logging of all operations:

1. **Security Events**:
   - Access denied
   - Request limit exceeded
   - Invalid input data
   - Security violations

2. **Usage**:
   - Executed commands
   - Processed AI requests
   - File access

### Log Format

Logs contain the following information:
```
[LEVEL] [DATE] [USER_ID] [CHANNEL_ID] MESSAGE
```

Example:
```
[INFO] [2025-01-15 14:30:25] [123456789] [987654321] Executed /search command
```

## ⚠️ Incident Response

### Incident Types

1. **Critical**:
   - Unauthorized access to administrative functions
   - Mass rate limiting bypass attempts
   - Suspicious AI function activity

2. **Medium**:
   - Request limit exceeded
   - Invalid input data
   - Authentication failures

3. **Low**:
   - Failed login attempts
   - Requests from suspicious IP addresses
   - Unusual activity

### Response Plan

1. **Detection**:
   - Log and metrics monitoring
   - Activity pattern analysis

2. **Assessment**:
   - Threat level determination
   - Impact assessment

3. **Response**:
   - User/function blocking
   - Administrator alerts
   - Evidence collection

4. **Analysis**:
   - Incident cause investigation
   - Vulnerability identification

5. **Recovery**:
   - Functionality restoration
   - Identified vulnerability fixes

## 🛠️ Administrator Commands

### Monitoring

```
/status - shows system status
/logs last:10 - shows last 10 log entries
/security incidents - shows recent security incidents
```

### Security Management

```
/security block user:ID - blocks user
/security unblock user:ID - unblocks user
/security limits reset - resets limits for all users
```

## 🔧 Security Updates

### Regular Checks

**Weekly**:
- Security log review
- Metrics analysis
- Dependency updates

**Monthly**:
- Access rights audit
- API key verification
- Security system testing

### Key Rotation

It is recommended to regularly rotate API keys:
1. Create a new key in the appropriate service
2. Update the `.env` file
3. Restart the bot
4. Delete the old key

## 🦙 Local AI and Privacy

### Local AI Benefits (Ollama)

- Data does not leave the infrastructure
- No API key limitations
- Offline operation capability

### Setup

By default, AI processing is performed locally through Ollama:
- Data (queries, RAG context, results) does not leave the host
- External providers (OpenAI/Anthropic) are enabled only with explicit configuration

## 📞 Security Support

### Contacts

- **Administrator**: @admin
- **Technical Support**: support@example.com
- **Emergency Cases**: emergency@example.com

### Resources

- [Discord Developer Portal](https://discord.com/developers/applications)
- [Google Cloud Security](https://cloud.google.com/security)
- [Redis Security](https://redis.io/topics/security)

© 2025 Dmitry Shivachov (Dmitze). All rights reserved.