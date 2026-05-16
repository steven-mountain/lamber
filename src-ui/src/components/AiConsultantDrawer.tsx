import React, { useState, useRef, useEffect } from 'react';
import { fetch } from '@tauri-apps/plugin-http';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Bot, X, Settings, Send, Loader2 } from 'lucide-react';

interface Message {
  role: 'user' | 'assistant';
  content: string;
}

export default function AiConsultantDrawer() {
  const [isOpen, setIsOpen] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [input, setInput] = useState('');
  const [messages, setMessages] = useState<Message[]>([
    { role: 'assistant', content: '您好！我是 Lamber 智能售前顾问。我可以帮您分析当前页面的项目效益、推荐内置产品。请问有什么可以帮您？' }
  ]);
  const [isTyping, setIsTyping] = useState(false);
  
  // Settings state with persistence
  const [endpoint, setEndpoint] = useState(() => localStorage.getItem('lamber_ai_endpoint') || 'http://localhost:11434/v1/chat/completions');
  const [model, setModel] = useState(() => localStorage.getItem('lamber_ai_model') || 'gemma:7b');
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('lamber_ai_api_key') || '');

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

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isTyping]);

  const handleSend = async () => {
    if (!input.trim() || isTyping) return;
    
    const userMessage = input.trim();
    setInput('');
    setMessages(prev => [...prev, { role: 'user', content: userMessage }]);
    setIsTyping(true);
    
    // Create an empty assistant message to append stream to
    setMessages(prev => [...prev, { role: 'assistant', content: '' }]);

    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          ...(apiKey ? { 'Authorization': `Bearer ${apiKey}` } : {})
        },
        body: JSON.stringify({
          model: model,
          messages: [
            // Future: Inject context here via System Prompt
            { role: 'system', content: 'You are a helpful presales AI consultant.' },
            ...messages.map(m => ({ role: m.role, content: m.content })),
            { role: 'user', content: userMessage }
          ],
          stream: true
        })
      });

      if (!response.ok) {
        throw new Error(`API returned ${response.status} ${response.statusText}`);
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
          processLine(line);
        }

        if (done) {
          if (incompleteLine) processLine(incompleteLine);
          break;
        }
      }

      function processLine(line: string) {
        const trimmedLine = line.trim();
        if (!trimmedLine || trimmedLine === 'data: [DONE]') return;
        
        if (trimmedLine.startsWith('data: ')) {
          const jsonStr = trimmedLine.slice(6).trim();
          if (!jsonStr) return;
          
          try {
            const data = JSON.parse(jsonStr);
            const deltaContent = data.choices?.[0]?.delta?.content ?? 
                                data.choices?.[0]?.text ?? 
                                data.message?.content ?? 
                                '';
            
            if (deltaContent) {
              setMessages(prev => {
                const newMessages = [...prev];
                const lastIdx = newMessages.length - 1;
                const lastMsg = { ...newMessages[lastIdx] };
                const currentText = lastMsg.content;

                // 寻找最大重叠部分 (Suffix-Prefix Overlap)
                // 解决“你好！欢迎你好！欢迎来到”这种重复拼接问题
                let overlap = 0;
                const maxPossibleOverlap = Math.min(currentText.length, deltaContent.length);
                for (let i = maxPossibleOverlap; i > 0; i--) {
                  if (currentText.endsWith(deltaContent.substring(0, i))) {
                    overlap = i;
                    break;
                  }
                }

                // 拼接非重叠部分
                lastMsg.content = currentText + deltaContent.substring(overlap);
                newMessages[lastIdx] = lastMsg;
                return newMessages;
              });
            }
          } catch (e) {
            console.error("Error parsing JSON:", jsonStr, e);
          }
        }
      }
    } catch (error) {
      console.error("Chat error:", error);
      setMessages(prev => {
        const newMessages = [...prev];
        newMessages[newMessages.length - 1].content = `**Error:** 连接 AI 服务失败 (${(error as Error).message})`;
        return newMessages;
      });
    } finally {
      setIsTyping(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && e.ctrlKey) {
      e.preventDefault();
      handleSend();
    }
  };

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
        className={`fixed top-0 right-0 h-screen w-96 bg-background border-l border-border shadow-2xl z-[60] flex flex-col transition-transform duration-300 ease-in-out ${isOpen ? 'translate-x-0' : 'translate-x-full'}`}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-border bg-card">
          <div className="flex items-center gap-2 font-bold text-foreground text-lg">
            <Bot size={24} className="text-primary" />
            智能售前顾问
          </div>
          <button onClick={() => setIsOpen(false)} className="text-muted-foreground hover:bg-muted p-1 rounded-md transition-colors">
            <X size={20} />
          </button>
        </div>

        {/* Chat Area */}
        <div className="flex-1 overflow-y-auto p-6 flex flex-col gap-4">
          {messages.map((msg, idx) => (
            <div 
              key={idx} 
              className={`max-w-[85%] p-3 rounded-2xl ${msg.role === 'user' ? 'bg-primary text-primary-foreground self-end rounded-br-sm' : 'bg-muted text-foreground self-start rounded-bl-sm border border-border'}`}
            >
              {msg.role === 'user' ? (
                <div className="whitespace-pre-wrap text-sm">{msg.content}</div>
              ) : (
                <div className="prose prose-sm dark:prose-invert prose-p:leading-relaxed prose-pre:p-0 max-w-none text-sm break-words ai-markdown">
                  <ReactMarkdown 
                    remarkPlugins={[remarkGfm]}
                    components={{
                      p: ({node, ...props}) => <p className="mb-2 last:mb-0" {...props} />
                    }}
                  >
                    {msg.content}
                  </ReactMarkdown>
                </div>
              )}
            </div>
          ))}
          {isTyping && (
            <div className="bg-muted text-foreground self-start rounded-2xl rounded-bl-sm border border-border p-3 flex items-center gap-1">
               <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        {/* Settings Panel */}
        {showSettings && (
          <div className="p-4 bg-muted/50 border-t border-border flex flex-col gap-3 text-sm">
            <div className="flex flex-col gap-1">
              <label className="text-muted-foreground font-medium text-xs">API Endpoint</label>
              <input 
                type="text" 
                value={endpoint} 
                onChange={e => setEndpoint(e.target.value)}
                className="bg-background border border-border rounded px-2 py-1 outline-none focus:border-primary"
              />
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-muted-foreground font-medium text-xs">Model Name</label>
              <input 
                type="text" 
                value={model} 
                onChange={e => setModel(e.target.value)}
                className="bg-background border border-border rounded px-2 py-1 outline-none focus:border-primary"
              />
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-muted-foreground font-medium text-xs">API Key (Optional)</label>
              <input 
                type="password" 
                value={apiKey} 
                onChange={e => setApiKey(e.target.value)}
                placeholder="Bearer Token"
                className="bg-background border border-border rounded px-2 py-1 outline-none focus:border-primary"
              />
            </div>
          </div>
        )}

        {/* Input Area */}
        <div className="p-4 border-t border-border bg-card">
          <div 
            className="flex items-center gap-1 text-xs text-muted-foreground font-medium cursor-pointer mb-2 hover:text-primary transition-colors w-max"
            onClick={() => setShowSettings(!showSettings)}
          >
            <Settings size={14} /> 设置
          </div>
          <div className="flex items-end gap-2">
            <textarea 
              value={input}
              onChange={e => setInput(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="输入问题，按 Ctrl+Enter 发送..."
              className="flex-1 bg-muted border border-border rounded-xl px-3 py-2 text-sm outline-none focus:border-primary resize-none min-h-[44px] max-h-32"
              rows={1}
            />
            <button 
              onClick={handleSend}
              disabled={!input.trim() || isTyping}
              className="w-10 h-10 rounded-full bg-primary text-primary-foreground flex flex-shrink-0 items-center justify-center disabled:opacity-50 disabled:cursor-not-allowed transition-opacity"
            >
              <Send size={18} className={input.trim() ? 'ml-1' : ''} />
            </button>
          </div>
        </div>
      </div>
    </>
  );
}
