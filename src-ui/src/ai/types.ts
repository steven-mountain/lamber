/**
 * Prompt Runtime AST Types
 * Establish a formal structure for prompt governance, decoupling UI from LLM instructions.
 */

// 1. System Rule Definition (Supports priority and dynamic pruning)
export interface PromptRule {
  id: string;
  content: string;
  priority: number;
  removable?: boolean;
}

// 2. Unified Context Node Definition (Replaces ad-hoc string/object mixing)
export interface ContextNode {
  type: 'json' | 'markdown' | 'summary';
  title?: string;
  content: any; // Raw data to be rendered by the compiler
  priority?: number;
  metadata?: {
    module?: string;
    updatedAt?: number;
  };
}

export interface AiImageAttachment {
  id: string;
  name: string;
  mimeType: string;
  size: number;
  dataUrl: string;
}

export interface AiChatMessage {
  role: 'user' | 'assistant';
  content: string;
  think?: string;
  images?: AiImageAttachment[];
}

export interface UserIntent {
  raw: string;
  intentType?: string; // Reserved for intent classification
  images?: AiImageAttachment[];
}

// 3. Prompt Abstract Syntax Tree (AST)
export interface PromptAST {
  systemRules: PromptRule[];
  dynamicState: {
    layer1Core: ContextNode[];   // Foundation (Primary calc data)
    layer2Active: ContextNode[]; // Current focus full data
    layer3Context: ContextNode[];// Background associated summaries
  };
  userIntent: UserIntent;
}

// 4. Runtime Telemetry & Events
export type RuntimeEventType = 
  | 'PromptRendered' 
  | 'ToolCallStarted' 
  | 'ToolCallCompleted' 
  | 'ToolError' 
  | 'StreamStarted' 
  | 'StreamCompleted';

export interface RuntimeEvent {
  type: RuntimeEventType;
  timestamp: number;
  payload?: any;
}
