# Docker Deployment Guide

This guide explains how to deploy the Discord AI Assistant Bot using Docker and Docker Compose.

## Overview

The project uses Docker Compose to manage multiple services:
- **discord-bot**: The main Discord bot service
- **n8n**: Workflow automation platform
- **PostgreSQL**: Database for n8n
- **Redis**: In-memory data structure store

## Docker Compose Configuration

The main configuration is in [docker-compose.yml](../../docker-compose.yml) which defines all services:

### Discord Bot Service
- Built from [src/config/docker/Dockerfile](../../src/config/docker/Dockerfile)
- Port: 3000
- Environment variables from `.env` file
- Volumes for persistent data

### n8n Service
- Image: `n8n/n8n:latest`
- Port: 5678
- Environment variables for authentication
- Volumes for persistent data

### PostgreSQL Service
- Image: `postgres:15`
- Port: 5432
- Environment variables for database configuration
- Volume for data persistence

### Redis Service
- Image: `redis:7-alpine`
- Port: 6379
- Volume for data persistence

## Prerequisites

1. Docker Engine installed
2. Docker Compose installed
3. Environment variables configured in `.env` file

## Setup Instructions

### 1. Environment Configuration
Create a `.env` file based on `.env.example`:

```bash
cp .env.example .env
```

Edit the `.env` file and set your values:

```env
# Discord Configuration
DISCORD_TOKEN=your_discord_bot_token
DISCORD_CLIENT_ID=your_discord_client_id
DISCORD_GUILD_ID=your_guild_id

# Google Configuration
GOOGLE_API_KEY=your_google_api_key
GOOGLE_APP_SCRIPT_URL=your_app_script_url

# n8n Configuration
N8N_USER=your_n8n_username
N8N_PASSWORD=your_n8n_password

# PostgreSQL Configuration
POSTGRES_USER=your_postgres_user
POSTGRES_PASSWORD=your_postgres_password
POSTGRES_DB=your_database_name

# AI Configuration (optional)
OPENAI_API_KEY=your_openai_api_key
OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=llama3
```

### 2. Start Services
Start all services using Docker Compose:

```bash
docker-compose up -d
```

This will start:
- Discord bot on port 3000
- n8n on http://localhost:5678
- PostgreSQL on port 5432
- Redis on port 6379

### 3. Access Services
- **Discord Bot**: Runs on port 3000 (internal)
- **n8n**: http://localhost:5678
- **PostgreSQL**: localhost:5432
- **Redis**: localhost:6379

## Management Commands

### View Logs
```bash
# View all logs
docker-compose logs

# View specific service logs
docker-compose logs discord-bot
docker-compose logs n8n
docker-compose logs postgres
docker-compose logs redis
```

### Stop Services
```bash
docker-compose down
```

### Update Services
```bash
# Pull latest images
docker-compose pull

# Restart services
docker-compose up -d
```

## Troubleshooting

### Common Issues
1. **Port conflicts**: Ensure ports 3000, 5678, 5432, and 6379 are free
2. **Permission issues**: Check file permissions for volumes
3. **Environment variables**: Ensure all required variables are set

### Debugging
```bash
# Check service status
docker-compose ps

# View real-time logs
docker-compose logs -f

# Execute commands in containers
docker-compose exec discord-bot sh
docker-compose exec n8n sh
docker-compose exec postgres psql -U $POSTGRES_USER -d $POSTGRES_DB
```

## Security Considerations

1. Change default passwords in `.env`
2. Use secure passwords for all services
3. Restrict access to exposed ports
4. Regularly update Docker images
5. Use Docker secrets for sensitive data in production

## Performance Optimization

1. Adjust resource limits in docker-compose.yml:
   ```yaml
   discord-bot:
     # ... existing config ...
     deploy:
       resources:
         limits:
           cpus: '0.5'
           memory: 512M
   ```

2. Use external volumes for better performance:
   ```yaml
   volumes:
     data_volume:
       driver: local
       driver_opts:
         type: none
         o: bind
         device: /path/to/host/directory
   ```

## Backup and Restore

### Backup
```bash
# Backup PostgreSQL data
docker-compose exec postgres pg_dump -U $POSTGRES_USER -d $POSTGRES_DB > backup.sql
```

### Restore
```bash
# Restore PostgreSQL data
docker-compose exec -T postgres psql -U $POSTGRES_USER -d $POSTGRES_DB < backup.sql
```