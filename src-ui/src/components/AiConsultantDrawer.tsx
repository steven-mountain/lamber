import React, { useState, useRef, useEffect } from 'react';
import { Bot, X, Settings, Send, Loader2, Sparkles, MessageSquare, Trash2 } from 'lucide-react';
import { SYSTEM_PROMPT_KNOWLEDGE } from '../lib/knowledgeBase';
import { useAiContextStore } from '../store/useAiContextStore';
import { AiRuntime } from '../ai/AiRuntime';
import { PromptAST, ContextNode, PromptRule } from '../ai/types';
import { useStreamingParser } from '../hooks/useStreamingParser';
import MessageBubble from './MessageBubble';
import { clsx } from 'clsx';

interface Message {
  role: 'user' | 'assistant';
  content: string;
  think?: string;
}

interface AiConsultantDrawerProps {
  currentView: string;
}

export default function AiConsultantDrawer({ currentView }: AiConsultantDrawerProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [input, setInput] = useState('');
  const [messages, setMessages] = useState<Message[]>([
    { role: 'assistant', content: '您好！我是 Lamber 智能售前顾问。我可以帮您分析当前页面的项目效益、推荐内置产品。请问有什么可以帮您？' }
  ]);
  const [isTyping, setIsTyping] = useState(false);
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null);
  
  // Settings state with persistence
  const [endpoint, setEndpoint] = useState(() => localStorage.getItem('lamber_ai_endpoint') || 'http://localhost:11434/v1/chat/completions');
  const [model, setModel] = useState(() => localStorage.getItem('lamber_ai_model') || 'gemma:7b');
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('lamber_ai_api_key') || '');

  // AI Context Store
  const businessData = useAiContextStore(state => state.businessData);
  const activeModule = useAiContextStore(state => state.activeModule);
  const lastUpdated = useAiContextStore(state => state.lastUpdated);

  // Runtime Infrastructure
  const runtime = useRef(new AiRuntime());
  const [loadingStatus, setLoadingStatus] = useState("正在分析...");
  const statusTimerRef = useRef<NodeJS.Timeout | null>(null);

  // Streaming Parser
  const { 
    normalText, 
    thinkText, 
    parseChunk, 
    finalize, 
    reset: resetParser,
    stop: stopParser 
  } = useStreamingParser();
  const chatContainerRef = useRef<HTMLDivElement>(null);
  const isAtBottom = useRef(true);
  const abortControllerRef = useRef<AbortController | null>(null);

  useEffect(() => {
    localStorage.setItem('lamber_ai_endpoint', endpoint);
  }, [endpoint]);

  useEffect(() => {
    localStorage.setItem('lamber_ai_model', model);
  }, [model]);

  useEffect(() => {
    localStorage.setItem('lamber_ai_api_key', apiKey);
  }, [apiKey]);

  const messagesEndRef = useRef<HTMLDivElement>(null);

  const handleScroll = () => {
    if (!chatContainerRef.current) return;
    const { scrollTop, scrollHeight, clientHeight } = chatContainerRef.current;
    isAtBottom.current = scrollHeight - scrollTop - clientHeight < 100;
  };

  useEffect(() => {
    if (isAtBottom.current) {
      messagesEndRef.current?.scrollIntoView({ behavior: 'auto' });
    }
  }, [messages, normalText, thinkText]);

  // Sync parser state to the last message
  useEffect(() => {
    if (isTyping) {
      setMessages(prev => {
        if (!prev || prev.length === 0) return prev;
        const newMessages = [...prev];
        const lastIdx = newMessages.length - 1;
        const lastMsg = newMessages[lastIdx];
        
        // Safety check: Only update if the last message is an assistant response
        if (lastMsg?.role === 'assistant') {
          newMessages[lastIdx] = { 
            ...lastMsg, 
            content: normalText || lastMsg.content,
            think: thinkText || lastMsg.think
          };
          return newMessages;
        }
        return prev;
      });
    }
  }, [normalText, thinkText, isTyping]);

  useEffect(() => {
    return () => {
      if (statusTimerRef.current) clearTimeout(statusTimerRef.current);
    };
  }, []);

  const handleSend = async (overrideInput?: string) => {
    const textToSend = overrideInput || input;
    if (!textToSend.trim() || isTyping) return;
    
    const userMessage = textToSend.trim();
    if (!overrideInput) setInput('');
    
    const updatedMessages: Message[] = [...messages, { role: 'user', content: userMessage }];
    setMessages(updatedMessages);
    setIsTyping(true);
    
    setMessages(prev => [...prev, { role: 'assistant', content: '' }]);
    
    // --- Enterprise LLM Infrastructure: AST Construction ---
    const systemRules: PromptRule[] = [
      { id: 'presales_role', content: 'You are a helpful presales AI consultant for Lamber system.', priority: 100 },
      { id: 'knowledge_base', content: SYSTEM_PROMPT_KNOWLEDGE, priority: 90 },
      { id: 'code_priority', content: '优先根据 [产品编号] (如 A302600342) 在知识库中匹配产品。只有当编号缺失时，才根据名称进行模糊匹配。', priority: 85 },
      { id: 'data_awareness', content: 'ALWAYS check the BUSINESS CONTEXT before answering. If data is missing, state it clearly.', priority: 80 }
    ];

    // Layer 1: Core (ICT Main Table)
    const layer1Core: ContextNode[] = businessData['ict'] ? [{
      type: 'json',
      title: '主测算核心指标',
      content: businessData['ict'],
      metadata: { module: 'ict', updatedAt: lastUpdated['ict'] }
    }] : [];

    // Layer 2: Active Workspace (Current template)
    const layer2Active: ContextNode[] = (activeModule !== 'ict' && businessData[activeModule]) ? [{
      type: 'json',
      title: `当前工作空间: ${activeModule.replace('template_', '')}`,
      content: businessData[activeModule],
      metadata: { module: activeModule, updatedAt: lastUpdated[activeModule] }
    }] : [];

    // Layer 3: Context (Other documents)
    const layer3Context: ContextNode[] = Object.keys(businessData)
      .filter(m => m !== 'ict' && m !== activeModule)
      .map(m => ({
        type: 'json',
        title: `关联文档: ${m.replace('template_', '')}`,
        content: businessData[m],
        metadata: { module: m, updatedAt: lastUpdated[m] }
      }));

    const ast: PromptAST = {
      systemRules,
      dynamicState: { layer1Core, layer2Active, layer3Context },
      userIntent: { raw: userMessage }
    };

    // --- Progressive UX: Start Status Timer ---
    setLoadingStatus("正在分析...");
    if (statusTimerRef.current) clearTimeout(statusTimerRef.current);
    statusTimerRef.current = setTimeout(() => {
      setLoadingStatus("正在提取「项目关联文档」数据...");
    }, 1500);

    try {
      resetParser();
      if (abortControllerRef.current) abortControllerRef.current.abort();
      abortControllerRef.current = new AbortController();

      await runtime.current.execute(
        ast, 
        (chunk) => parseChunk(chunk),
        { endpoint, model, apiKey },
        abortControllerRef.current.signal
      );
      finalize();
    } catch (error) {
      stopParser();
      if ((error as Error).name === 'AbortError') {
        console.log("Stream aborted by user");
        return;
      }
      
      console.error("Chat error:", error);
      setMessages(prev => {
        if (!prev || prev.length === 0) return prev;
        const newMessages = [...prev];
        const lastIdx = newMessages.length - 1;
        if (newMessages[lastIdx]?.role === 'assistant') {
          newMessages[lastIdx] = { 
            ...newMessages[lastIdx], 
            content: `**Error:** 连接 AI 服务失败 (${(error as Error).message})` 
          };
        }
        return newMessages;
      });
    } finally {
      setIsTyping(false);
      if (statusTimerRef.current) clearTimeout(statusTimerRef.current);
    }
  };

  const handleStop = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
    stopParser();
    setIsTyping(false);
  };

  const clearMessages = () => {
    if (window.confirm('确定要清除所有聊天记录吗？')) {
      if (abortControllerRef.current) abortControllerRef.current.abort();
      resetParser();
      setMessages([
        { role: 'assistant', content: '聊天记录已清除。请问还有什么可以帮您？' }
      ]);
    }
  };

  const copyToClipboard = (text: string, idx: number) => {
    navigator.clipboard.writeText(text);
    setCopiedIdx(idx);
    setTimeout(() => setCopiedIdx(null), 2000);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && e.ctrlKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const quickActions = [
    { label: '分析当前项目效益', icon: <Sparkles size={14} />, view: 'ict' },
    { label: '推荐合适产品', icon: <Bot size={14} /> },
    { label: '生成立项摘要', icon: <MessageSquare size={14} />, view: 'docfill' }
  ].filter(a => !a.view || a.view === currentView);

  return (
    <>
      {/* FAB */}
      <div 
        className="fixed bottom-10 right-10 w-14 h-14 bg-primary text-primary-foreground rounded-full shadow-lg flex items-center justify-center cursor-pointer z-50 hover:scale-105 transition-transform"
        onClick={() => setIsOpen(true)}
      >
        <Bot size={28} />
      </div>

      {/* Drawer */}
      <div 
        className={clsx(
          "fixed top-0 right-0 h-screen w-[420px] bg-background border-l border-border shadow-2xl z-[60] flex flex-col transition-transform duration-300 ease-in-out",
          isOpen ? 'translate-x-0' : 'translate-x-full'
        )}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-border bg-card">
          <div className="flex items-center gap-2 font-bold text-foreground text-lg">
            <Bot size={24} className="text-primary" />
            智能售前顾问
          </div>
          <div className="flex items-center gap-2">
            <button 
              onClick={clearMessages} 
              title="清除聊天记录"
              className="text-muted-foreground hover:bg-destructive/10 hover:text-destructive p-1.5 rounded-md transition-colors"
            >
              <Trash2 size={18} />
            </button>
            <button onClick={() => setIsOpen(false)} className="text-muted-foreground hover:bg-muted p-1 rounded-md transition-colors">
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Chat Area */}
        <div 
          ref={chatContainerRef}
          onScroll={handleScroll}
          className="flex-1 overflow-y-auto p-6 flex flex-col gap-6"
        >
          {messages.map((msg, idx) => (
            <MessageBubble 
              key={idx}
              msg={msg}
              idx={idx}
              isStreaming={isTyping && idx === messages.length - 1 && msg.role === 'assistant'}
              onCopy={copyToClipboard}
              copiedIdx={copiedIdx}
            />
          ))}
          {messages.length > 0 && isTyping && !messages[messages.length - 1]?.content && !messages[messages.length - 1]?.think && (
            <div className="bg-muted text-foreground self-start rounded-2xl rounded-bl-sm border border-border p-4 flex items-center gap-3 shadow-sm animate-in fade-in duration-300">
               <Loader2 className="h-4 w-4 animate-spin text-primary" />
               <span className="text-xs font-bold text-secondary-foreground animate-pulse">{loadingStatus}</span>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        {/* Settings Panel */}
        {showSettings && (
          <div className="p-4 bg-muted/50 border-t border-border flex flex-col gap-3 text-sm animate-in slide-in-from-bottom-2 duration-200">
            <div className="flex flex-col gap-1">
              <label className="text-muted-foreground font-medium text-xs">API Endpoint</label>
              <input 
                type="text" 
                value={endpoint} 
                onChange={e => setEndpoint(e.target.value)}
                className="bg-background border border-border rounded px-2 py-1.5 outline-none focus:border-primary"
              />
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-muted-foreground font-medium text-xs">Model Name</label>
              <input 
                type="text" 
                value={model} 
                onChange={e => setModel(e.target.value)}
                className="bg-background border border-border rounded px-2 py-1.5 outline-none focus:border-primary"
              />
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-muted-foreground font-medium text-xs">API Key (可选)</label>
              <input 
                type="password" 
                value={apiKey} 
                onChange={e => setApiKey(e.target.value)}
                placeholder="Bearer Token"
                className="bg-background border border-border rounded px-2 py-1.5 outline-none focus:border-primary"
              />
            </div>
          </div>
        )}

        {/* Input Area */}
        <div className="p-4 border-t border-border bg-card">
          {!isTyping && (
            <div className="flex flex-wrap gap-2 mb-3">
              {quickActions.map((action, i) => (
                <button
                  key={i}
                  onClick={() => handleSend(action.label)}
                  className="flex items-center gap-1.5 bg-muted hover:bg-primary/10 hover:text-primary text-secondary-foreground text-[11px] font-semibold px-3 py-1.5 rounded-full border border-border transition-all"
                >
                  {action.icon}
                  {action.label}
                </button>
              ))}
            </div>
          )}

          <div className="flex items-center justify-between mb-2">
            <div 
              className="flex items-center gap-1 text-[11px] text-muted-foreground font-bold cursor-pointer hover:text-primary transition-colors"
              onClick={() => setShowSettings(!showSettings)}
            >
              <Settings size={14} /> 模型设置
            </div>
            {currentView !== 'hub' && (
              <div className="text-[10px] text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-full font-bold flex items-center gap-1 shadow-sm border border-emerald-100">
                <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse" />
                已连接实时业务状态
              </div>
            )}
          </div>
          
          <div className="flex items-end gap-2">
            <textarea 
              value={input}
              onChange={e => setInput(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder={isTyping ? "AI 正在思考中..." : "输入问题，按 Ctrl+Enter 发送..."}
              disabled={isTyping}
              className="flex-1 bg-muted border border-border rounded-xl px-4 py-2.5 text-sm outline-none focus:border-primary resize-none min-h-[44px] max-h-32 shadow-inner disabled:opacity-70"
              rows={1}
            />
            {isTyping && (
              <button
                onClick={handleStop}
                className="p-2 text-muted-foreground hover:text-destructive transition-colors"
                title="停止生成"
              >
                <X size={20} />
              </button>
            )}
            {!isTyping && (
              <button 
                onClick={() => handleSend()}
                disabled={!input.trim() || isTyping}
                className="w-10 h-10 rounded-xl bg-primary text-primary-foreground flex flex-shrink-0 items-center justify-center disabled:opacity-50 disabled:cursor-not-allowed transition-all active:scale-95 shadow-sm hover:shadow-md"
              >
                <Send size={18} className={input.trim() ? 'ml-1' : ''} />
              </button>
            )}
          </div>
        </div>
      </div>
    </>
  );
}
