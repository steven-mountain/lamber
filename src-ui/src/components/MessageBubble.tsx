import React, { useState, useRef, useEffect } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Sparkles, Check, Copy, ChevronDown, ChevronUp, Bot } from 'lucide-react';
import { clsx } from 'clsx';

interface Message {
  role: 'user' | 'assistant';
  content: string;
  think?: string;
}

interface MessageBubbleProps {
  msg: Message;
  idx: number;
  isStreaming: boolean;
  onCopy: (text: string, idx: number) => void;
  copiedIdx: number | null;
}

const MessageBubble = ({ msg, idx, isStreaming, onCopy, copiedIdx }: MessageBubbleProps) => {
  const [isExpanded, setIsExpanded] = useState(true);
  const hasUserTouched = useRef(false);

  // 1. Parser Recovery / Empty Ghost Defense
  if (!msg.content?.trim() && !msg.think?.trim() && msg.role === 'assistant') {
    return null;
  }

  // 2. Auto-collapse Logic
  useEffect(() => {
    // When streaming stops, if user hasn't interacted, auto-collapse the think block
    if (!isStreaming && msg.think && !hasUserTouched.current) {
      setIsExpanded(false);
    }
  }, [isStreaming, msg.think]);

  const handleHeaderClick = () => {
    hasUserTouched.current = true;
    setIsExpanded(!isExpanded);
  };

  return (
    <div 
      className={clsx(
        "flex flex-col gap-2 max-w-[85%] group animate-in fade-in slide-in-from-bottom-2 duration-300",
        msg.role === 'user' ? "self-end" : "self-start"
      )}
    >
      <div className={clsx(
        "px-4 py-3 rounded-2xl shadow-sm border",
        msg.role === 'user' 
          ? "bg-primary text-primary-foreground rounded-br-sm border-primary/20" 
          : "bg-card text-card-foreground rounded-bl-sm border-border"
      )}>
        {/* Thinking Block */}
        {msg.think && (
          <div className="mb-3 overflow-hidden rounded-lg border border-border/50 bg-muted/30">
            <button 
              onClick={handleHeaderClick}
              className="flex w-full items-center justify-between px-3 py-2 hover:bg-muted/50 transition-colors"
            >
              <div className="flex items-center gap-2">
                {isStreaming && !msg.content ? (
                  <Sparkles className="h-3 w-3 text-primary animate-pulse" />
                ) : (
                  <Bot className="h-3 w-3 text-muted-foreground" />
                )}
                <span className="text-[11px] font-bold uppercase tracking-wider text-muted-foreground">
                  {isStreaming ? "✨ 思考中..." : "🧠 思考已完成"}
                </span>
              </div>
              {isExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>
            
            {isExpanded && (
              <div className="px-3 pb-3 pt-1 text-xs leading-relaxed text-muted-foreground/80 italic border-t border-border/30 whitespace-pre-wrap">
                {msg.think}
              </div>
            )}
          </div>
        )}

        {/* Content Block */}
        {msg.content && (
          <div className="relative">
            {isStreaming ? (
              /* Phase 1: Throttled Plain Text Rendering to avoid Markdown parsing overhead during streaming */
              <div className="whitespace-pre-wrap text-sm leading-relaxed break-words overflow-hidden">
                {String(msg.content)}
              </div>
            ) : (
              <ReactMarkdown 
                remarkPlugins={[remarkGfm]}
                components={{
                  p: ({node, children, ...props}) => {
                    return (
                      <p className="mb-3 last:mb-0" {...props}>
                        {React.Children.map(children, (child) => {
                          if (typeof child === 'string') {
                            if (child.includes('[系统内置]')) {
                              return child.split('[系统内置]').map((part, i, arr) => (
                                <React.Fragment key={i}>
                                  {part}
                                  {i < arr.length - 1 && (
                                    <span className="inline-flex items-center gap-1 bg-primary/10 text-primary px-1.5 py-0.5 rounded-md font-bold border border-primary/20 mx-1">
                                      <Check size={12} strokeWidth={3} /> 系统内置
                                    </span>
                                  )}
                                </React.Fragment>
                              ));
                            }
                            if (child.includes('【系统外扩展】')) {
                               return child.split('【系统外扩展】').map((part, i, arr) => (
                                <React.Fragment key={i}>
                                  {part}
                                  {i < arr.length - 1 && (
                                    <span className="inline-flex items-center gap-1 bg-muted-foreground/10 text-muted-foreground px-1.5 py-0.5 rounded-md font-medium border border-border mx-1 italic text-[11px]">
                                      系统外扩展
                                    </span>
                                  )}
                                </React.Fragment>
                              ));
                            }
                          }
                          return child;
                        })}
                      </p>
                    );
                  }
                }}
              >
                {String(msg.content || '')}
              </ReactMarkdown>
            )}
            
            {/* Copy Button */}
            {!isStreaming && (
              <button
                onClick={() => onCopy(msg.content, idx)}
                className="absolute -right-2 -bottom-2 p-1.5 rounded-md bg-background border border-border opacity-0 group-hover:opacity-100 transition-opacity shadow-sm hover:bg-muted"
                title="复制内容"
              >
                {copiedIdx === idx ? <Check size={12} className="text-green-500" /> : <Copy size={12} className="text-muted-foreground" />}
              </button>
            )}
          </div>
        )}
      </div>
      
      {/* Timestamp or Status */}
      <span className="text-[10px] text-muted-foreground px-1">
        {msg.role === 'user' ? "您" : "Lamber 智能顾问"}
      </span>
    </div>
  );
};

export default React.memo(MessageBubble, (prev, next) => {
  return (
    prev.msg === next.msg &&
    prev.isStreaming === next.isStreaming &&
    prev.copiedIdx === next.copiedIdx
  );
});
