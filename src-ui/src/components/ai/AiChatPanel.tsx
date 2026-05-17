import { useEffect, useRef, useState } from 'react';
import { Bot, Loader2, MessageSquare, Settings, Sparkles, Trash2 } from 'lucide-react';
import { SYSTEM_PROMPT_KNOWLEDGE } from '../../lib/knowledgeBase';
import { useAiContextStore } from '../../store/useAiContextStore';
import { AiRuntime } from '../../ai/AiRuntime';
import type { AiChatMessage, AiImageAttachment, ContextNode, PromptAST, PromptRule } from '../../ai/types';
import { useStreamingParser } from '../../hooks/useStreamingParser';
import MessageBubble from '../MessageBubble';
import AiInputBox from './AiInputBox';

interface AiChatPanelProps {
  currentView?: string;
}

export default function AiChatPanel({ currentView = 'hub' }: AiChatPanelProps) {
  const [showSettings, setShowSettings] = useState(false);
  const [input, setInput] = useState('');
  const [images, setImages] = useState<AiImageAttachment[]>([]);
  const [messages, setMessages] = useState<AiChatMessage[]>([
    { role: 'assistant', content: '您好！我是 Lamber 智能售前顾问。我可以帮您分析当前页面的项目效益、推荐内置产品。请问有什么可以帮您？' },
  ]);
  const [isTyping, setIsTyping] = useState(false);
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null);

  // Settings state with persistence
  const [endpoint, setEndpoint] = useState(() => localStorage.getItem('lamber_ai_endpoint') || 'http://localhost:11434/v1/chat/completions');
  const [model, setModel] = useState(() => localStorage.getItem('lamber_ai_model') || 'gemma:7b');
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('lamber_ai_api_key') || '');
  const [visionEnabled, setVisionEnabled] = useState(() => localStorage.getItem('lamber_ai_vision_enabled') === 'true');

  // AI Context Store
  const businessData = useAiContextStore(state => state.businessData);
  const activeModule = useAiContextStore(state => state.activeModule);
  const lastUpdated = useAiContextStore(state => state.lastUpdated);

  // Runtime Infrastructure
  const runtime = useRef(new AiRuntime());
  const [loadingStatus, setLoadingStatus] = useState('正在分析...');
  const statusTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Streaming Parser
  const {
    normalText,
    thinkText,
    parseChunk,
    finalize,
    reset: resetParser,
    stop: stopParser,
  } = useStreamingParser();
  const chatContainerRef = useRef<HTMLDivElement>(null);
  const isAtBottom = useRef(true);
  const abortControllerRef = useRef<AbortController | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    localStorage.setItem('lamber_ai_endpoint', endpoint);
  }, [endpoint]);

  useEffect(() => {
    localStorage.setItem('lamber_ai_model', model);
  }, [model]);

  useEffect(() => {
    localStorage.setItem('lamber_ai_api_key', apiKey);
  }, [apiKey]);

  useEffect(() => {
    localStorage.setItem('lamber_ai_vision_enabled', String(visionEnabled));
    if (!visionEnabled) {
      setImages([]);
    }
  }, [visionEnabled]);

  const handleScroll = () => {
    if (!chatContainerRef.current) return;
    const { scrollTop, scrollHeight, clientHeight } = chatContainerRef.current;
    isAtBottom.current = scrollHeight - scrollTop - clientHeight < 100;
  };

  useEffect(() => {
    if (!isAtBottom.current) return;
    const frameId = requestAnimationFrame(() => {
      messagesEndRef.current?.scrollIntoView({ behavior: 'auto' });
    });
    return () => cancelAnimationFrame(frameId);
  }, [messages]);

  useEffect(() => {
    return () => {
      if (statusTimerRef.current) clearTimeout(statusTimerRef.current);
      if (abortControllerRef.current) abortControllerRef.current.abort();
      stopParser();
    };
  }, [stopParser]);

  useEffect(() => {
    if (isAtBottom.current) {
      messagesEndRef.current?.scrollIntoView({ behavior: 'auto' });
    }
  }, []);

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
          const nextContent = normalText || lastMsg.content;
          const nextThink = thinkText || lastMsg.think;
          if (lastMsg.content === nextContent && lastMsg.think === nextThink) {
            return prev;
          }
          newMessages[lastIdx] = {
            ...lastMsg,
            content: nextContent,
            think: nextThink,
          };
          return newMessages;
        }
        return prev;
      });
    }
  }, [normalText, thinkText, isTyping]);

  const handleSend = async (overrideInput?: string) => {
    const textToSend = overrideInput ?? input;
    const imagesToSend = overrideInput ? [] : [...images];
    if ((!textToSend.trim() && imagesToSend.length === 0) || isTyping) return;

    const userMessage = textToSend.trim();
    const promptText = userMessage || '请分析图片内容。';

    if (!overrideInput) {
      setInput('');
      setImages([]);
    }

    const updatedMessages: AiChatMessage[] = [
      ...messages,
      { role: 'user', content: userMessage, images: imagesToSend },
    ];
    setMessages(updatedMessages);
    setIsTyping(true);

    setMessages(prev => [...prev, { role: 'assistant', content: '' }]);

    // --- Enterprise LLM Infrastructure: AST Construction ---
    const systemRules: PromptRule[] = [
      { id: 'presales_role', content: 'You are a helpful presales AI consultant for Lamber system.', priority: 100 },
      { id: 'knowledge_base', content: SYSTEM_PROMPT_KNOWLEDGE, priority: 90 },
      { id: 'code_priority', content: '优先根据 [产品编号] (如 A302600342) 在知识库中匹配产品。只有当编号缺失时，才根据名称进行模糊匹配。', priority: 85 },
      { id: 'data_awareness', content: 'ALWAYS check the BUSINESS CONTEXT before answering. If data is missing, state it clearly.', priority: 80 },
    ];

    const contextView = currentView || 'hub';
    const isIctContext = contextView === 'ict';
    const isDocfillContext = contextView === 'docfill' || contextView.startsWith('template_');
    const activeContextModule = activeModule && activeModule !== 'hub' ? activeModule : '';
    const shouldUseActiveTemplate = (isIctContext || isDocfillContext)
      && activeContextModule
      && activeContextModule !== 'ict'
      && Boolean(businessData[activeContextModule]);

    // Layer 1: Core (ICT Main Table). Never inject stale ICT data while the main window is on the hub.
    const layer1Core: ContextNode[] = isIctContext && businessData['ict'] ? [{
      type: 'json',
      title: '主测算核心指标',
      content: businessData['ict'],
      metadata: { module: 'ict', updatedAt: lastUpdated['ict'] },
    }] : [];

    // Layer 2: Active Workspace (Current template)
    const layer2Active: ContextNode[] = shouldUseActiveTemplate ? [{
      type: 'json',
      title: `当前工作空间: ${activeContextModule.replace('template_', '')}`,
      content: businessData[activeContextModule],
      metadata: { module: activeContextModule, updatedAt: lastUpdated[activeContextModule] },
    }] : [];

    // Layer 3: Context (Other documents)
    const layer3Context: ContextNode[] = Object.keys(businessData)
      .filter((moduleKey) => {
        if (isIctContext) {
          return moduleKey !== 'ict' && moduleKey !== activeContextModule;
        }
        if (isDocfillContext) {
          return moduleKey.startsWith('template_') && moduleKey !== activeContextModule;
        }
        return false;
      })
      .map(m => ({
        type: 'json',
        title: `关联文档: ${m.replace('template_', '')}`,
        content: businessData[m],
        metadata: { module: m, updatedAt: lastUpdated[m] },
      }));

    const ast: PromptAST = {
      systemRules,
      dynamicState: { layer1Core, layer2Active, layer3Context },
      userIntent: {
        raw: promptText,
        images: imagesToSend.length > 0 ? imagesToSend : undefined,
      },
    };

    // --- Progressive UX: Start Status Timer ---
    setLoadingStatus('正在分析...');
    if (statusTimerRef.current) clearTimeout(statusTimerRef.current);
    statusTimerRef.current = setTimeout(() => {
      setLoadingStatus('正在提取「项目关联文档」数据...');
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
        console.log('Stream aborted by user');
        return;
      }

      console.error('Chat error:', error);
      setMessages(prev => {
        if (!prev || prev.length === 0) return prev;
        const newMessages = [...prev];
        const lastIdx = newMessages.length - 1;
        if (newMessages[lastIdx]?.role === 'assistant') {
          newMessages[lastIdx] = {
            ...newMessages[lastIdx],
            content: `**Error:** 连接 AI 服务失败 (${(error as Error).message})`,
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
        { role: 'assistant', content: '聊天记录已清除。请问还有什么可以帮您？' },
      ]);
    }
  };

  const copyToClipboard = (text: string, idx: number) => {
    navigator.clipboard.writeText(text);
    setCopiedIdx(idx);
    setTimeout(() => setCopiedIdx(null), 2000);
  };

  const contextView = currentView && currentView !== 'hub' ? currentView : 'hub';
  const quickActionView = contextView.startsWith('template_') ? 'docfill' : contextView;
  const quickActions = [
    { label: '分析当前项目效益', icon: <Sparkles size={14} />, view: 'ict' },
    { label: '推荐合适产品', icon: <Bot size={14} /> },
    { label: '生成立项摘要', icon: <MessageSquare size={14} />, view: 'docfill' },
  ].filter(action => !action.view || action.view === quickActionView);

  return (
    <div className="flex min-h-0 flex-1 flex-col bg-background">
      <div
        ref={chatContainerRef}
        onScroll={handleScroll}
        className="flex flex-1 flex-col gap-6 overflow-y-auto p-5"
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
          <div className="flex animate-in items-center gap-3 self-start rounded-2xl rounded-bl-sm border border-border bg-muted p-4 text-foreground shadow-sm fade-in duration-300">
            <Loader2 className="h-4 w-4 animate-spin text-primary" />
            <span className="animate-pulse text-xs font-bold text-secondary-foreground">{loadingStatus}</span>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {showSettings && (
        <div className="flex flex-col gap-3 border-t border-border bg-muted/50 p-4 text-sm animate-in slide-in-from-bottom-2 duration-200">
          <div className="flex flex-col gap-1">
            <label className="text-xs font-medium text-muted-foreground">API Endpoint</label>
            <input
              type="text"
              value={endpoint}
              onChange={event => setEndpoint(event.target.value)}
              className="rounded border border-border bg-background px-2 py-1.5 outline-none focus:border-primary"
            />
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-medium text-muted-foreground">Model Name</label>
            <input
              type="text"
              value={model}
              onChange={event => setModel(event.target.value)}
              className="rounded border border-border bg-background px-2 py-1.5 outline-none focus:border-primary"
            />
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-medium text-muted-foreground">API Key (可选)</label>
            <input
              type="password"
              value={apiKey}
              onChange={event => setApiKey(event.target.value)}
              placeholder="Bearer Token"
              className="rounded border border-border bg-background px-2 py-1.5 outline-none focus:border-primary"
            />
          </div>
          <label className="flex items-center gap-2 text-xs font-medium text-muted-foreground">
            <input
              type="checkbox"
              checked={visionEnabled}
              onChange={event => setVisionEnabled(event.target.checked)}
              className="h-4 w-4 accent-primary"
            />
            启用图片输入
          </label>
          {visionEnabled && (
            <p className="text-[11px] leading-5 text-muted-foreground">
              当前模型需支持 vision / image_url 格式，否则图片可能被忽略或请求失败。
            </p>
          )}
        </div>
      )}

      <div className="border-t border-border bg-card p-4">
        {!isTyping && (
          <div className="mb-3 flex flex-wrap gap-2">
            {quickActions.map((action, index) => (
              <button
                key={index}
                type="button"
                onClick={() => handleSend(action.label)}
                className="flex items-center gap-1.5 rounded-full border border-border bg-muted px-3 py-1.5 text-[11px] font-semibold text-secondary-foreground transition-all hover:bg-primary/10 hover:text-primary"
              >
                {action.icon}
                {action.label}
              </button>
            ))}
          </div>
        )}

        <div className="mb-2 flex items-center justify-between">
          <button
            type="button"
            className="flex items-center gap-1 text-[11px] font-bold text-muted-foreground transition-colors hover:text-primary"
            onClick={() => setShowSettings(!showSettings)}
          >
            <Settings size={14} /> 模型设置
          </button>
          <div className="flex items-center gap-2">
            {quickActionView !== 'hub' && (
              <div className="flex items-center gap-1 rounded-full border border-emerald-100 bg-emerald-50 px-2 py-0.5 text-[10px] font-bold text-emerald-600 shadow-sm">
                <span className="h-1.5 w-1.5 animate-pulse rounded-full bg-emerald-500" />
                已连接实时业务状态
              </div>
            )}
            <button
              type="button"
              onClick={clearMessages}
              title="清除聊天记录"
              className="rounded-md p-1.5 text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
            >
              <Trash2 size={16} />
            </button>
          </div>
        </div>

        <AiInputBox
          input={input}
          images={images}
          isTyping={isTyping}
          visionEnabled={visionEnabled}
          onInputChange={setInput}
          onImagesChange={setImages}
          onSend={() => handleSend()}
          onStop={handleStop}
        />
      </div>
    </div>
  );
}
