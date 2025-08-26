import type { ChatInputCommandInteraction, InteractionEditReplyOptions } from 'discord.js';

type MinimalEmbed = { data: { title: string; description?: string } };

interface BotLike {
  getService?: (name: string) => unknown;
  handleError?: (err: unknown) => Promise<void> | void;
}

export async function execute(interaction: ChatInputCommandInteraction, bot?: BotLike): Promise<void> {
  const ai = bot?.getService?.('ai') as any;
  const rag = bot?.getService?.('rag') as any;
  const enhancedRag = bot?.getService?.('enhancedRag') as any;
  const cache = bot?.getService?.('cache') as any;
  const responseCache = bot?.getService?.('responseCache') as any;
  const contextMemory = bot?.getService?.('contextMemory') as any;
  const knowledgeBase = bot?.getService?.('knowledgeBase') as any;
  const metrics = bot?.getService?.('metrics') as any;

  const started = performance.now();
  const userId = interaction.user.id;
  const channelId = interaction.channelId;
  
  try {
    const query = interaction.options.getString?.('запит', false)
      ?? interaction.options.getString?.('query', false)
      ?? '';
    const context = interaction.options.getString?.('контекст', false)
      ?? interaction.options.getString?.('context', false)
      ?? null;
    const mode = interaction.options.getString?.('режим', false)
      ?? interaction.options.getString?.('mode', false)
      ?? 'general';

    await interaction.deferReply?.();

    // Add query to context memory
    let queryId: string | undefined;
    if (contextMemory) {
      queryId = contextMemory.addQuery(userId, channelId, query, 'ai', { mode, context });
    }

    // Build contextual prompt using user history
    let contextualQuery = query;
    if (contextMemory && query.trim()) {
      contextualQuery = contextMemory.buildContextualPrompt(userId, query);
    }

    // Enhanced cache key including user context
    const cacheKey = responseCache 
      ? `ai_enhanced:${userId}:${mode}:${query}:${context ?? ''}` 
      : `ai:${mode}:${query}:${context ?? ''}`;
    
    let answer: string | null = null;
    let sources: string[] = []; // Move sources declaration here
    
    // Check response cache first
    if (responseCache?.get) {
      answer = await responseCache.get(cacheKey);
    } else if (cache?.get) {
      answer = await cache.get(cacheKey);
    }

    if (!answer) {
      let searchResults: any[] = [];
      // Remove sources declaration from here since it's now at the top

      // Try Enhanced RAG first (with auto-indexing)
      if (enhancedRag?.search && typeof enhancedRag.search === 'function') {
        try {
          searchResults = await enhancedRag.search(contextualQuery, {
            limit: Number(process.env['RETRIEVER_K'] ?? 6),
            useCache: true
          });
          sources = searchResults.map((r, i) => `[${i + 1}] ${r.fileName || 'Document'}`).filter(Boolean);
          
          if (searchResults.length > 0) {
            // Use AI to generate response based on search results
            const searchContext = searchResults.map(r => r.content).join('\n\n');
            const prompt = `На основі наступної інформації з документів, дайте відповідь на запит користувача.\n\nКонтекст:\n${searchContext}\n\nЗапит: ${query}\n\nВідповідь:`;
            
            if (ai?.generateResponse) {
              answer = await ai.generateResponse(prompt, { maxTokens: 512 });
              answer += `\n\nДжерела: ${sources.join(', ')}`;
            }
          }
        } catch (ragError) {
          console.warn('Enhanced RAG search failed, falling back to standard methods:', ragError);
        }
      }
      
      // Fallback to Knowledge Base search
      if (!answer && knowledgeBase?.search) {
        try {
          const kbResults = await knowledgeBase.search({
            query: contextualQuery,
            limit: 5,
            useSemanticSearch: false
          });
          
          if (kbResults.length > 0) {
            const kbContext = kbResults.map((r: any) => r.entry.content).join('\n\n');
            const kbSources = kbResults.map((r: any, i: number) => `[${i + 1}] ${r.entry.title}`);
            
            if (ai?.generateResponse) {
              const prompt = `На основі наступної інформації з бази знань, дайте відповідь на запит користувача.\n\nБаза знань:\n${kbContext}\n\nЗапит: ${query}\n\nВідповідь:`;
              answer = await ai.generateResponse(prompt, { maxTokens: 512 });
              answer += `\n\nДжерела з бази знань: ${kbSources.join(', ')}`;
              sources = [...sources, ...kbSources];
            }
          }
        } catch (kbError) {
          console.warn('Knowledge Base search failed:', kbError);
        }
      }
      
      // Fallback to standard RAG
      if (!answer && rag?.answer && typeof rag.answer === 'function') {
        try {
          const res = await rag.answer(contextualQuery, {
            k: Number(process.env['RETRIEVER_K'] ?? 6),
            alpha: Number(process.env['RETRIEVER_ALPHA'] ?? 0.5),
          }, {
            maskPII: true,
            maxTokens: Number(process.env['RAG_MAX_CONTEXT_TOKENS'] ?? 1200),
          }, {
            maxTokens: Number(process.env['AI_MAX_TOKENS'] ?? 512),
          });
          const ragSources = res.chunks?.map((c: any, i: number) => `[${i + 1}] ${c.name}`).join(', ');
          answer = `${res.answer}\n\nДжерела: ${ragSources || '—'}`;
          if (ragSources) sources.push(ragSources);
        } catch (ragError) {
          console.warn('Standard RAG failed:', ragError);
        }
      }
      
      // Final fallback to direct AI
      if (!answer) {
        if (!ai?.generateResponse) {
          throw new Error('AI service unavailable');
        }
        const plain = await ai.generateResponse(String(contextualQuery || 'Запит від користувача'), { maxTokens: 512 });
        answer = typeof plain === 'string' ? plain : String(plain?.content ?? '');
      }
      
      // Cache the response
      if (responseCache?.set) {
        responseCache.set(cacheKey, answer, 30, { // 30 minutes
          source: 'ai_assistant',
          tags: ['ai', mode, userId],
          size: answer.length
        });
      } else {
        await cache?.set?.(cacheKey, answer);
      }
    }

    // Update context memory with response
    if (contextMemory && queryId) {
      contextMemory.updateQueryResponse(queryId, answer, {
        responseTime: performance.now() - started,
        sources: sources.length > 0 ? sources : undefined,
        mode
      });
    }

    const embed: MinimalEmbed = { 
      data: { 
        title: '🤖 AI Відповідь', 
        description: String(answer) 
      } 
    };
    
    await interaction.editReply?.({ embeds: [embed] } as unknown as InteractionEditReplyOptions);
    
    metrics?.incrementCommand?.('ai', 'success');
    metrics?.measureCommandDuration?.('ai', performance.now() - started);
  } catch (err) {
    metrics?.incrementCommand?.('ai', 'error');
    metrics?.measureCommandDuration?.('ai', performance.now() - started);
    await bot?.handleError?.(err);
  }
}