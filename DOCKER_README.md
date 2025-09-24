# Docker Setup for Discord AI Assistant Bot

This document explains how to set up and run the Discord AI Assistant Bot using Docker.

## Prerequisites

1. Docker installed on your system
2. Docker Compose installed (usually included with Docker Desktop)
3. Properly configured `.env` file with all required environment variables

## Docker Configuration

The project includes two Docker configurations:

1. **Single Service**: Runs just the Discord bot
2. **Full Stack**: Runs the bot along with all required services (Redis, PostgreSQL, n8n, Prometheus)

## Quick Start

### 1. Prepare Environment Variables

Copy the example environment file and configure it:

```bash
cp .env.example .env
```

Edit the `.env` file and set all required variables:
- Discord bot token
- Google API credentials
- Database credentials
- AI service configuration

### 2. Build and Run with Docker Compose

For the full stack (recommended):

```bash
docker-compose -f docker-compose.improved.yml up -d
```

For just the bot service:

```bash
docker-compose -f docker-compose.improved.yml up -d discord-bot
```

### 3. Check Service Status

```bash
docker-compose -f docker-compose.improved.yml ps
```

### 4. View Logs

```bash
docker-compose -f docker-compose.improved.yml logs -f discord-bot
```

## Services Included

1. **discord-bot**: The main Discord AI Assistant Bot application
2. **n8n**: Workflow automation platform
3. **postgres**: PostgreSQL database for persistent storage
4. **redis**: Redis cache for improved performance
5. **prometheus**: Monitoring and alerting toolkit

## Configuration Files

- `Dockerfile.improved`: Multi-stage Docker image definition
- `docker-compose.improved.yml`: Service definitions and orchestration
- `healthcheck.improved.js`: Health check script for the bot service

## Data Persistence

The following volumes are used for data persistence:

- `postgres_data`: PostgreSQL database files
- `redis_data`: Redis cache data
- `./data`: Application data directory
- `./workspace`: Workspace files
- `./logs`: Application logs
- `./n8n-data`: n8n workflow data

## Health Checks

All services include health checks to ensure they're running properly:

- **discord-bot**: Checks the `/health` endpoint
- **n8n**: Checks the `/healthz` endpoint
- **postgres**: Uses `pg_isready` command
- **redis**: Uses `redis-cli ping` command
- **prometheus**: Checks the `/-/healthy` endpoint

## Troubleshooting

### Common Issues

1. **Port conflicts**: Ensure ports 3000, 5432, 6379, 5678, and 9090 are available
2. **Permission issues**: Make sure Docker has access to the project directory
3. **Environment variables**: Verify all required variables are set in `.env`

### Logs

Check logs for specific services:

```bash
# Bot logs
docker-compose -f docker-compose.improved.yml logs discord-bot

# Database logs
docker-compose -f docker-compose.improved.yml logs postgres

# Redis logs
docker-compose -f docker-compose.improved.yml logs redis
```

### Restart Services

To restart specific services:

```bash
# Restart just the bot
docker-compose -f docker-compose.improved.yml restart discord-bot

# Restart all services
docker-compose -f docker-compose.improved.yml restart
```

### Stop Services

```bash
# Stop all services
docker-compose -f docker-compose.improved.yml down

# Stop and remove volumes (WARNING: This will delete data)
docker-compose -f docker-compose.improved.yml down -v
```

## Updating the Application

To update the application:

1. Pull the latest code:
   ```bash
   git pull
   ```

2. Rebuild the Docker images:
   ```bash
   docker-compose -f docker-compose.improved.yml build
   ```

3. Restart the services:
   ```bash
   docker-compose -f docker-compose.improved.yml up -d
   ```

## Customization

### Resource Limits

You can add resource limits to prevent services from consuming too many resources:

```yaml
services:
  discord-bot:
    # ... other config
    deploy:
      resources:
        limits:
          cpus: '0.5'
          memory: 512M
        reservations:
          cpus: '0.25'
          memory: 256M
```

### Network Configuration

To use a different network configuration, modify the `networks` section in the docker-compose file.

## Security Considerations

1. The bot runs as a non-root user for security
2. Sensitive data should be stored in Docker secrets for production deployments
3. Ensure the `.env` file is not committed to version control
4. Use HTTPS in production environments