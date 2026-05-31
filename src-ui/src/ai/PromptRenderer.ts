import { PromptAST, ContextNode } from './types';

/**
 * Enterprise-grade Prompt Compiler / Renderer
 * Implements a multi-pass pipeline to transform an AST into a high-quality LLM prompt.
 */
export class PromptRenderer {

  constructor() {
    // model-specific rendering logic can be added here
  }

  /**
   * Main compilation pipeline
   */
  public render(ast: PromptAST): string {
    // Pass 1: Pruning/Cleaning - Remove empty nodes
    const cleanedAST = this.passPruning(ast);

    // Pass 2: Rendering - Convert nodes to text and add metadata markers
    const renderedSections = this.passRendering(cleanedAST);

    // Pass 3: Assembly - Final formatting and joining
    return this.passAssembly(renderedSections, cleanedAST.userIntent.raw);
  }

  /**
   * Pass 1: Cleaning & Pruning
   * Filters out nodes with empty content (null, undefined, {}, [])
   */
  private passPruning(ast: PromptAST): PromptAST {
    const filterEmpty = (nodes: ContextNode[]) => {
      return nodes.filter(node => {
        const c = node.content;
        if (c === null || c === undefined) return false;
        if (typeof c === 'string' && c.trim() === '') return false;
        if (Array.isArray(c) && c.length === 0) return false;
        if (typeof c === 'object' && Object.keys(c).length === 0) return false;
        return true;
      });
    };

    return {
      ...ast,
      dynamicState: {
        layer1Core: filterEmpty(ast.dynamicState.layer1Core),
        layer2Active: filterEmpty(ast.dynamicState.layer2Active),
        layer3Context: filterEmpty(ast.dynamicState.layer3Context),
      }
    };
  }

  /**
   * Pass 2: Rendering
   * Converts typed data into semantic text
   */
  private passRendering(ast: PromptAST): Record<string, string[]> {
    const sections: Record<string, string[]> = {
      rules: ast.systemRules.sort((a, b) => b.priority - a.priority).map(r => r.content),
      core: ast.dynamicState.layer1Core.map(n => this.renderNode(n)),
      active: ast.dynamicState.layer2Active.map(n => this.renderNode(n)),
      context: ast.dynamicState.layer3Context.map(n => this.renderNode(n)),
    };
    return sections;
  }

  /**
   * Pass 3: Assembly
   * Final formatting with HARD_CONSTRAINTS to prevent Context Hijacking.
   */
  private passAssembly(sections: Record<string, string[]>, userRaw: string): string {
    const parts: string[] = [];

    // 1. System Identity & Rules
    if (sections.rules.length > 0) {
      parts.push("# SYSTEM_RULES\n" + sections.rules.join('\n'));
    }

    // 2. Hard Constraints (Attention Management)
    parts.push(`# HARD_CONSTRAINTS
- 当用户进行简单问候（如“你好”、“你是谁”）时，必须直接用自然语言回答，不要输出长篇大论。
- 绝对禁止主动复述、打印或全量总结下方的 CONTEXT_REFERENCE 数据，除非用户明确要求分析！
- 保持回答简洁专业，直击要点。`);

    // 3. Dynamic Context Reference (Passive Data)
    if (sections.core.length > 0 || sections.active.length > 0 || sections.context.length > 0) {
      parts.push("# CONTEXT_REFERENCE (PASSIVE DATA ONLY)\n<reference_data>");
      
      if (sections.core.length > 0) {
        parts.push("## Layer 1: Saved Official Project State (Workspace SQLite)\n" + sections.core.join('\n'));
      }
      if (sections.active.length > 0) {
        parts.push("## Layer 2: Current Page Context and Unsaved Draft Overlay\n" + sections.active.join('\n'));
      }
      if (sections.context.length > 0) {
        parts.push("## Layer 3: Context Loading Notes\n" + sections.context.join('\n'));
      }
      
      parts.push("</reference_data>");
    }

    // 4. User Intent (Highest Priority, placed at the end)
    parts.push("\n# USER_INTENT\n" + userRaw);

    return parts.join('\n\n');
  }

  /**
   * Logic to render an individual ContextNode
   */
  private renderNode(node: ContextNode): string {
    let output = "";
    const title = node.title ? `**[${node.title}]** ` : "";
    
    // Check for recent updates (last 60 seconds)
    const isRecent = node.metadata?.updatedAt && (Date.now() - node.metadata.updatedAt < 60000);
    const updateMarker = isRecent ? " (最新修改)" : "";

    switch (node.type) {
      case 'markdown':
        output = `${title}${node.content}${updateMarker}`;
        break;
      case 'summary':
        output = `${title}${node.content}${updateMarker}`;
        break;
      case 'json':
        output = `${title}${updateMarker}\n${this.serializeJson(node.content)}`;
        break;
    }
    
    return output;
  }

  /**
   * Semantic JSON serialization
   */
  private serializeJson(data: any): string {
    if (!data) return "";
    
    // Flatten if it's a simple KV object
    const lines: string[] = [];
    Object.entries(data).forEach(([key, value]) => {
      if (value === null || value === undefined || value === '') return;
      
      // Basic humanization of keys
      let label = key.replace(/^gen_/, '').replace(/^demand_/, '').replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
      
      const labelMap: Record<string, string> = {
        'Mid Three Code': '产品编号',
        'Mid Three Name': '产品名称',
        'Customer Confirm': '客户确认',
        'Proj Name': '项目名称',
        'Project Name': '项目名称',
        'Self Three': '自主三问',
        'Self Three Selected': '自主三问',
        'Self Three Reminder': '自主三问提醒',
        'Self Three Missing Fees': '自主三问缺失费用提醒',
        'Cashflow Model': '资金收付模型',
        'Cashflow Segment Value Mode': '分板块金额录入方式',
        'Cashflow Segments': '分板块资金计划',
      };

      if (labelMap[label]) label = labelMap[label];

      if (Array.isArray(value)) {
        lines.push(`- ${label}: ${JSON.stringify(value)}`);
      } else if (typeof value === 'object') {
        // Simple one-level deep stringify for objects
        lines.push(`- ${label}: ${JSON.stringify(value)}`);
      } else {
        lines.push(`- ${label}: ${value}`);
      }
    });
    
    return lines.join('\n');
  }
}
