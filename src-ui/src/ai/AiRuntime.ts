import { PromptAST, RuntimeEvent, RuntimeEventType } from './types';
import { PromptRenderer } from './PromptRenderer';

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
    signal?: AbortSignal
  ): Promise<void> {
    try {
      // 1. Render Prompt
      const compiledPrompt = this.renderer.render(ast);
      this.addTrace('PromptRendered', { length: compiledPrompt.length });

      // 2. Fetch from LLM
      this.addTrace('StreamStarted');
      
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
            { role: 'user', content: ast.userIntent.raw }
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
