# 🏗️ Discord AI Assistant Bot Architecture

## 📋 Contents

## 🗺️ Overview

Architecture Diagram

Main Components

1. **Bot Core** (`src/core/Bot.ts`)
   - Main bot class, initialization, event handling
   - Command registration and management
   - DI container for services
   - Event system

2. **Commands** (`src/commands/`)
   - Parent class for all commands
   - Search in documents
   - Document management
   - Usage statistics
   - File management

3. **Services** (`src/services/`)
   - Google API integration
   - Embedding generation and work
   - RAG pipeline
   - Data caching

4. **Search** (`src/search/`)
   - Search index interface
   - FTS implementation on SQLite
   - Hybrid search (FTS + vector)

5. **RAG** (`src/rag/`)
   - Search for relevant fragments
   - Context preparation
   - Response generation

## 🔄 Data Flows

Document Indexing

Request Processing

## ⚙️ Technology Stack

Main Technologies

- **Language**: TypeScript 5.0+
- **Platform**: Node.js 20.x (LTS)
- **Framework**: Discord.js 14.x
- **Database**: SQLite3 (FTS5), Redis (cache)
- **AI/ML**: Ollama (local), OpenAI API (optional)
- **Integrations**: Google Sheets API
- **Libraries**: DI, Validation, Logging, Testing, Monitoring

## 🔧 Configuration

Main Configuration Parameters (`src/config/`):

More details: Setup Guide

## 🔐 Security

Key Mechanisms

1. **Component Signing**
   - All buttons and selectors are signed via HMAC-SHA256
   - TTL for each component (default 15 minutes)
   - Deny access to expired components

2. **Data Handling**
   - PII masking in logs
   - Encryption of confidential data
   - Limited access to API keys

3. **API Protection**
   - Rate limiting
   - Input data validation
   - Error handling without revealing details

More details: Security Guide

## 📚 Additional Documentation

Quick Start - how to get started quickly

RAG Guide - working with the RAG pipeline

Local AI - Ollama setup

API Documentation - bot API description

Development Guide - developer guide

## 🚀 Future Enhancements

1. Support for additional vector databases (Pinecone, Weaviate)
2. Improved processing of tables and other complex formats
3. Extended usage analytics
4. Graphical interface for settings
5. Additional cloud storage integrations

© 2025 Godzilla Bot Team
License: MIT
Change History