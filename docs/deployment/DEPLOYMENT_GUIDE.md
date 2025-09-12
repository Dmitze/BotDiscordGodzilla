# Deployment Guide

This guide provides comprehensive instructions for deploying the Discord AI Assistant Bot in various environments.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Environment Configuration](#environment-configuration)
- [Docker Deployment](#docker-deployment)
- [Manual Deployment](#manual-deployment)
- [Cloud Deployment](#cloud-deployment)
- [Service Management](#service-management)
- [Monitoring and Maintenance](#monitoring-and-maintenance)

## Prerequisites

### System Requirements
- Node.js 18+ (for manual deployment)
- Docker and Docker Compose (for Docker deployment)
- At least 2GB RAM
- At least 10GB free disk space

### Required Accounts
- Discord Developer Account
- Google Cloud Platform Account
- OpenAI API Key (optional, for AI features)
- Ollama (optional, for local AI)

## Environment Configuration

### Environment Variables
Create a `.env` file in the project root based on `.env.example`:

```bash
cp .env.example .env
```

Required variables:
```env
# Discord Configuration
DISCORD_TOKEN=your_discord_bot_token
DISCORD_CLIENT_ID=your_discord_client_id
DISCORD_GUILD_ID=your_guild_id

# Google Configuration
GOOGLE_API_KEY=your_google_api_key
GOOGLE_APP_SCRIPT_URL=your_app_script_url

# AI Configuration (optional)
OPENAI_API_KEY=your_openai_api_key
OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=llama3

# Database Configuration
DATABASE_URL=your_database_url

# Redis Configuration
REDIS_URL=redis://localhost:6379

# Security
ENCRYPTION_KEY=your_encryption_key
```

### Configuration Files
- `config/` - Contains configuration files for various services
- `src/config/` - Application-specific configuration

## Docker Deployment

### Quick Start
1. Install Docker and Docker Compose
2. Configure environment variables in `.env`
3. Run:
   ```bash
   docker-compose up -d
   ```

### Docker Services
The Docker Compose configuration includes:
- **discord-bot**: Main bot service
- **n8n**: Workflow automation
- **postgres**: Database for n8n
- **redis**: In-memory cache

### Docker Commands
```bash
# Start all services
docker-compose up -d

# View logs
docker-compose logs -f

# Stop services
docker-compose down

# Update services
docker-compose pull
docker-compose up -d

# Access container shell
docker-compose exec discord-bot sh
```

## Manual Deployment

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/Dmitze/BotDiscordGodzilla.git
   cd BotDiscordGodzilla
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Configure environment variables in `.env`

4. Build the project:
   ```bash
   npm run build
   ```

5. Start the bot:
   ```bash
   npm start
   ```

### Development Mode
```bash
# Start in development mode with auto-reload
npm run dev
```

### Production Mode
```bash
# Build and start in production mode
npm run build
npm start
```

## Cloud Deployment

### Heroku Deployment
1. Create a Heroku account
2. Install Heroku CLI
3. Login to Heroku:
   ```bash
   heroku login
   ```

4. Create a new app:
   ```bash
   heroku create your-app-name
   ```

5. Set environment variables:
   ```bash
   heroku config:set DISCORD_TOKEN=your_token
   heroku config:set GOOGLE_API_KEY=your_key
   # ... other variables
   ```

6. Deploy:
   ```bash
   git push heroku main
   ```

### AWS Deployment
1. Create an EC2 instance
2. Install Docker and Docker Compose
3. Clone the repository
4. Configure environment variables
5. Start services with Docker Compose

### Google Cloud Deployment
1. Create a Compute Engine instance
2. Install Docker and Docker Compose
3. Clone the repository
4. Configure environment variables
5. Start services with Docker Compose

## Service Management

### Process Management
For production deployments, use a process manager like PM2:

1. Install PM2 globally:
   ```bash
   npm install -g pm2
   ```

2. Start the bot with PM2:
   ```bash
   pm2 start dist/index.js --name discord-bot
   ```

3. Set up auto-restart on system boot:
   ```bash
   pm2 startup
   pm2 save
   ```

### Service Monitoring
```bash
# View running processes
pm2 list

# View logs
pm2 logs discord-bot

# Restart service
pm2 restart discord-bot

# Stop service
pm2 stop discord-bot
```

## Monitoring and Maintenance

### Health Checks
The bot includes built-in health checks:
- `/health` endpoint for HTTP health checks
- Docker health checks
- Log monitoring

### Log Management
Logs are stored in:
- `data/logs/` - Application logs
- `data/metrics/` - Performance metrics

### Backup Strategy
Regular backups should include:
- Configuration files
- Database data
- Custom workflows
- User data (if applicable)

### Update Process
1. Pull the latest code:
   ```bash
   git pull origin main
   ```

2. Install updated dependencies:
   ```bash
   npm install
   ```

3. Rebuild the project:
   ```bash
   npm run build
   ```

4. Restart the service:
   ```bash
   # For Docker
   docker-compose down
   docker-compose up -d
   
   # For PM2
   pm2 restart discord-bot
   ```

### Troubleshooting
Common issues and solutions:

1. **Bot not responding**:
   - Check Discord token validity
   - Verify bot is online in Discord
   - Check logs for errors

2. **Google API errors**:
   - Verify API key permissions
   - Check Google Sheets access
   - Validate spreadsheet ID

3. **AI service issues**:
   - Check OpenAI API key
   - Verify Ollama is running
   - Check model availability

4. **Database connection errors**:
   - Verify database credentials
   - Check network connectivity
   - Ensure database service is running

### Performance Optimization
1. Use Redis for caching
2. Implement proper rate limiting
3. Optimize database queries
4. Use connection pooling
5. Monitor memory usage

## Security Best Practices

1. Never commit sensitive data to version control
2. Use environment variables for secrets
3. Regularly rotate API keys
4. Implement proper access controls
5. Keep dependencies updated
6. Use HTTPS for all external communications
7. Implement rate limiting
8. Regular security audits

## Scaling Considerations

For high-traffic deployments:
1. Use load balancers
2. Implement horizontal scaling
3. Use external databases
4. Implement caching strategies
5. Monitor resource usage
6. Set up auto-scaling rules

## Backup and Recovery

### Backup Strategy
1. Daily configuration backups
2. Weekly database backups
3. Monthly full system backups
4. Store backups in multiple locations

### Recovery Process
1. Restore from latest backup
2. Verify configuration files
3. Test service functionality
4. Monitor for issues

## Migration Guide

When migrating between environments:
1. Export configuration
2. Backup data
3. Set up new environment
4. Import configuration
5. Restore data
6. Test functionality
7. Update DNS/URLs if needed