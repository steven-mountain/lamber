import type { AiChatMessage, PromptAST, RuntimeEvent, RuntimeEventType } from './types';
import { PromptRenderer } from './PromptRenderer';

type UserMessageContent =
  | string
  | Array<
      | { type: 'text'; text: string }
      | { type: 'image_url'; image_url: { url: string } }
    >;

type ChatCompletionMessage = {
  role: 'system' | 'user' | 'assistant';
  content: UserMessageContent;
};

const MAX_HISTORY_TOKENS = 4000;

function estimateTokens(text: string): number {
  let englishChars = 0;
  let cjkChars = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0x4e00 && code <= 0x9fff) {
      cjkChars++;
    } else {
      englishChars++;
    }
  }
  // English: ~4 chars per token. CJK: ~1.2 chars per token.
  return Math.ceil(englishChars / 4) + Math.ceil(cjkChars / 1.2);
}

/**
 * Enterprise AI Agent Runtime
 * Manages the execution lifecycle, telemetry, and fault tolerance of LLM interactions.
 */
export class AiRuntime {
  private renderer: PromptRenderer;
  private traces: RuntimeEvent[] = [];
  
  constructor() {
    this.renderer = new PromptRenderer();
  }

  public getTraces(): RuntimeEvent[] {
    return this.traces;
  }

  public clearTraces(): void {
    this.traces = [];
  }

  private addTrace(type: RuntimeEventType, payload?: any) {
    this.traces.push({
      type,
      timestamp: Date.now(),
      payload
    });
  }

  /**
   * Execute a prompt cycle with full lifecycle management
   */
  public async execute(
    ast: PromptAST, 
    onChunk: (chunk: string) => void,
    config: { endpoint: string; model: string; apiKey?: string },
    history: AiChatMessage[] = [],
    signal?: AbortSignal
  ): Promise<void> {
    try {
      // 1. Render Prompt
      const compiledPrompt = this.renderer.render(ast);
      this.addTrace('PromptRendered', { length: compiledPrompt.length });

      // 2. Fetch from LLM
      this.addTrace('StreamStarted');
      const userContent = this.buildUserContent(ast);
      const historyMessages = this.buildHistoryMessages(history);
      
      const response = await fetch(config.endpoint, {
        method: 'POST',
        signal,
        headers: {
          'Content-Type': 'application/json',
          ...(config.apiKey ? { 'Authorization': `Bearer ${config.apiKey}` } : {})
        },
        body: JSON.stringify({
          model: config.model,
          messages: [
            { role: 'system', content: compiledPrompt },
            ...historyMessages,
            { role: 'user', content: userContent }
          ],
          stream: true
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder('utf-8');
      if (!reader) throw new Error("No reader available");

      let incompleteLine = '';
      while (true) {
        const { value, done } = await reader.read();
        const chunk = value ? decoder.decode(value, { stream: true }) : '';
        const lines = (incompleteLine + chunk).split('\n');
        incompleteLine = lines.pop() || '';
        
        for (const line of lines) {
          const processed = this.processLine(line);
          if (processed) onChunk(processed);
        }

        if (done) {
          if (incompleteLine) {
            const processed = this.processLine(incompleteLine);
            if (processed) onChunk(processed);
          }
          break;
        }
      }

      this.addTrace('StreamCompleted');
    } catch (error) {
      this.addTrace('ToolError', { error: (error as Error).message });
      console.error("AI Runtime Execution Error:", error);
      throw error;
    }
  }

  private buildHistoryMessages(history: AiChatMessage[]): ChatCompletionMessage[] {
    const validMessages = history.filter(message => message.content.trim());
    const result: ChatCompletionMessage[] = [];
    let totalTokens = 0;

    // Traverse from latest to oldest
    for (let i = validMessages.length - 1; i >= 0; i--) {
      const message = validMessages[i];
      let content = message.content;
      let estTokens = estimateTokens(content);

      // If a single message is excessively large (exceeding half the budget),
      // we truncate it to prevent it from hogging the entire history context
      const singleMessageLimit = Math.floor(MAX_HISTORY_TOKENS / 2);
      if (estTokens > singleMessageLimit) {
        const truncateRatio = singleMessageLimit / estTokens;
        const targetLen = Math.floor(content.length * truncateRatio);
        content = `${content.slice(0, targetLen)}\n[conversation history truncated]`;
        estTokens = estimateTokens(content);
      }

      if (totalTokens + estTokens > MAX_HISTORY_TOKENS) {
        break; // Stop loading older messages once budget is exceeded
      }

      totalTokens += estTokens;
      result.unshift({
        role: message.role,
        content,
      });
    }

    return result;
  }

  private buildUserContent(ast: PromptAST): UserMessageContent {
    const images = (ast.userIntent.images || []).filter(image => Boolean(image.dataUrl));
    if (images.length === 0) {
      return ast.userIntent.raw;
    }

    return [
      {
        type: 'text',
        text: ast.userIntent.raw?.trim() || '请分析图片内容。',
      },
      ...images.map((image) => ({
        type: 'image_url' as const,
        image_url: {
          url: image.dataUrl as string,
        },
      })),
    ];
  }

  private processLine(line: string): string | null {
    const trimmedLine = line.trim();
    if (!trimmedLine || trimmedLine === 'data: [DONE]') return null;
    
    if (trimmedLine.startsWith('data: ')) {
      const jsonStr = trimmedLine.slice(6).trim();
      if (!jsonStr) return null;
      
      try {
        const data = JSON.parse(jsonStr);
        return data.choices?.[0]?.delta?.content ?? 
               data.choices?.[0]?.text ?? 
               data.message?.content ?? 
               null;
      } catch (e) {
        return null;
      }
    }
    return null;
  }

  /**
   * failure isolation for internal tool calls (Placeholder for future tool support)
   */
  public async invokeToolIsolated(toolName: string, args: any): Promise<any> {
    this.addTrace('ToolCallStarted', { toolName, args });
    try {
      // Logic for actual tool execution would go here
      // For now, it's a defensive wrapper
      const result = await Promise.resolve({}); 
      this.addTrace('ToolCallCompleted', { toolName });
      return this.sanitizeToolResult(result);
    } catch (error) {
      this.addTrace('ToolError', { toolName, error: (error as Error).message });
      return "Tool execution failed, please proceed with available context.";
    }
  }

  private sanitizeToolResult(result: any): any {
    if (!result) return {};
    if (typeof result !== 'object') return result;
    
    // Simple sanitization: remove nulls/undefined/empty
    return Object.fromEntries(
      Object.entries(result).filter(([_, v]) => {
        if (v === null || v === undefined) return false;
        if (Array.isArray(v) && v.length === 0) return false;
        if (typeof v === 'object' && Object.keys(v).length === 0) return false;
        return true;
      })
    );
  }
}
