import { useEffect, useRef, useState } from 'react';
import { listen } from '@tauri-apps/api/event';
import { SYSTEM_PROMPT_KNOWLEDGE } from '../../lib/knowledgeBase';
import {
  AI_CONTEXT_REFRESH_REQUEST_EVENT,
  AI_CONTEXT_UPDATED_EVENT,
  useAiContextStore,
} from '../../store/useAiContextStore';
import { AiRuntime } from '../../ai/AiRuntime';
import type { AiChatMessage, AiImageAttachment, ContextNode, PromptAST, PromptRule } from '../../ai/types';
import { useStreamingParser } from '../../hooks/useStreamingParser';
import MessageBubble from '../MessageBubble';
import AiInputBox from './AiInputBox';
import { AI_CONTEXT_KEY, getAiContextDisplayName, getAiContextScope, isAiContextKeyForView } from '../../utils/aiContextKeys';
import AppIcon, { type AppIconName } from '../icons/AppIcon';

interface AiChatPanelProps {
  currentView?: string;
}

function isTauriRuntime() {
  return typeof window !== 'undefined' && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

function getCoreContextKey(view: string) {
  if (view === 'ict') return AI_CONTEXT_KEY.ICT_CORE;
  if (view === 'benefit') return AI_CONTEXT_KEY.BENEFIT_CORE;
  if (view === 'docfill') return AI_CONTEXT_KEY.DOCFILL_CORE;
  return view;
}

function formatLastUpdated(timestamp?: number) {
  if (!timestamp) return '--';
  return new Date(timestamp).toLocaleString();
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
  const activeModule = useAiContextStore(state => state.activeModule);
  const businessData = useAiContextStore(state => state.businessData);
  const lastUpdated = useAiContextStore(state => state.lastUpdated);

  useEffect(() => {
    const hydrateAiContext = () => {
      useAiContextStore.getState().hydrateFromStorage();
    };

    hydrateAiContext();

    if (!isTauriRuntime()) return;

    let disposed = false;
    const unlistenFns: Array<() => void> = [];

    Promise.all([
      listen(AI_CONTEXT_UPDATED_EVENT, hydrateAiContext),
      listen(AI_CONTEXT_REFRESH_REQUEST_EVENT, hydrateAiContext),
    ]).then((handlers) => {
      if (disposed) {
        handlers.forEach(handler => handler());
        return;
      }
      unlistenFns.push(...handlers);
    }).catch((error) => {
      console.warn('Failed to listen for AI context events:', error);
    });

    return () => {
      disposed = true;
      unlistenFns.forEach(handler => handler());
    };
  }, []);

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
    const conversationHistory = messages.filter(message => message.content.trim());
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
      { id: 'currency_unit_policy', content: 'Currency unit policy: all financial amount fields from BUSINESS CONTEXT are CNY yuan (元) unless the context explicitly says otherwise. Never label those raw values as ten-thousand yuan / 万元. If the user explicitly asks for 万元, divide the yuan value by 10,000 and state that conversion.', priority: 84 },
      { id: 'data_awareness', content: 'ALWAYS check the BUSINESS CONTEXT before answering. If data is missing, state it clearly.', priority: 80 },
    ];

    useAiContextStore.getState().hydrateFromStorage();
    const latestState = useAiContextStore.getState();
    const {
      businessData: latestBusinessData,
      activeModule: latestActiveModule,
      lastUpdated: latestLastUpdated,
    } = latestState;

    const contextView = currentView || 'hub';
    const coreContextKey = contextView === 'hub' ? '' : getCoreContextKey(contextView);
    const activeContextModule = latestActiveModule
      && latestActiveModule !== AI_CONTEXT_KEY.HUB
      && isAiContextKeyForView(latestActiveModule, contextView)
      ? latestActiveModule
      : '';
    const shouldUseActiveContext = activeContextModule
      && activeContextModule
      && activeContextModule !== coreContextKey
      && Boolean(latestBusinessData[activeContextModule]);

    // Layer 1: Core (ICT Main Table). Never inject stale ICT data while the main window is on the hub.
    const layer1Core: ContextNode[] = coreContextKey && latestBusinessData[coreContextKey] ? [{
      type: 'json',
      title: getAiContextDisplayName(coreContextKey),
      content: latestBusinessData[coreContextKey],
      metadata: { module: coreContextKey, updatedAt: latestLastUpdated[coreContextKey] },
    }] : [];

    // Layer 2: Active Workspace (Current template)
    const layer2Active: ContextNode[] = shouldUseActiveContext ? [{
      type: 'json',
      title: getAiContextDisplayName(activeContextModule),
      content: latestBusinessData[activeContextModule],
      metadata: { module: activeContextModule, updatedAt: latestLastUpdated[activeContextModule] },
    }] : [];

    // Layer 3: Context (Other documents)
    const layer3Context: ContextNode[] = Object.keys(latestBusinessData)
      .filter(moduleKey => isAiContextKeyForView(moduleKey, contextView))
      .filter(moduleKey => moduleKey !== coreContextKey && moduleKey !== activeContextModule)
      .map(m => ({
        type: 'json',
        title: getAiContextDisplayName(m),
        content: latestBusinessData[m],
        metadata: { module: m, updatedAt: latestLastUpdated[m] },
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
        conversationHistory,
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
  const quickActionView = contextView;
  const quickActionItems = [
    { label: '分析当前项目效益', icon: 'ai', view: 'ict' },
    { label: '推荐合适产品', icon: 'aiThinking' },
    { label: '生成立项摘要', icon: 'document', view: 'docfill' },
  ] satisfies Array<{ label: string; icon: AppIconName; view?: string }>;
  const quickActions = quickActionItems.filter(action => !action.view || action.view === quickActionView);
  const quickActionContextKey = quickActionView === 'hub' ? '' : getCoreContextKey(quickActionView);
  const activeContextData = activeModule && activeModule !== AI_CONTEXT_KEY.HUB
    ? businessData[activeModule]
    : undefined;
  const quickActionContextData = quickActionContextKey
    ? businessData[quickActionContextKey] ?? businessData[quickActionView]
    : undefined;
  const connectedContextModule = activeContextData
    ? activeModule
    : quickActionContextData
      ? (businessData[quickActionContextKey] ? quickActionContextKey : quickActionView)
      : '';
  const connectedScope = connectedContextModule
    ? getAiContextScope(connectedContextModule) ?? connectedContextModule
    : '';
  const statusLastUpdated = connectedContextModule
    ? lastUpdated[connectedContextModule] ?? lastUpdated[quickActionView]
    : lastUpdated[activeModule];
  const connectionStatusText = connectedContextModule
    ? `已连接：${connectedScope}`
    : '未检测到业务状态';
  const connectionStatusClassName = connectedContextModule
    ? 'border-emerald-100 bg-emerald-50 text-emerald-600'
    : 'border-border bg-muted text-muted-foreground';
  const connectionDotClassName = connectedContextModule
    ? 'bg-emerald-500'
    : 'bg-muted-foreground/50';

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
            <AppIcon name="loading" size={16} className="animate-spin text-primary" />
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
                <AppIcon name={action.icon} size={14} />
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
            <AppIcon name="settings" size={14} /> 模型设置
          </button>
          <div className="flex items-center gap-2">
            <div
              className={`flex items-center gap-1 rounded-full border px-2 py-0.5 text-[10px] font-bold shadow-sm ${connectionStatusClassName}`}
              title={`activeModule: ${activeModule || '--'} · lastUpdated: ${formatLastUpdated(statusLastUpdated)}`}
            >
              <span className={`h-1.5 w-1.5 rounded-full ${connectedContextModule ? 'animate-pulse' : ''} ${connectionDotClassName}`} />
              <span>{connectionStatusText}</span>
            </div>
            <button
              type="button"
              onClick={clearMessages}
              title="清除聊天记录"
              className="rounded-md p-1.5 text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
            >
              <AppIcon name="delete" size={16} />
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
