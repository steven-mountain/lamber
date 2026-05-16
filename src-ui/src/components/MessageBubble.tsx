import React, { useState, useRef, useEffect } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Sparkles, Check, Copy, ChevronDown, ChevronUp } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

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

export default function MessageBubble({ msg, idx, isStreaming, onCopy, copiedIdx }: MessageBubbleProps) {
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

  return (
    <div 
      className={cn(
        "group relative max-w-[90%] p-4 rounded-2xl animate-in fade-in slide-in-from-bottom-2 duration-300",
        msg.role === 'user' ? 'bg-primary text-primary-foreground self-end rounded-br-sm' : 'bg-muted text-foreground self-start rounded-bl-sm border border-border'
      )}
    >
      {msg.role === 'user' ? (
        <div className="whitespace-pre-wrap text-sm leading-relaxed">{msg.content}</div>
      ) : (
        <div className="prose prose-sm dark:prose-invert prose-p:leading-relaxed prose-pre:p-0 max-w-none text-sm break-words ai-markdown">
          {/* Thinking Process Block */}
          {msg.think && (
            <div className="mb-4 bg-background/40 border border-border/50 rounded-lg overflow-hidden transition-all duration-300">
              <div 
                onClick={() => {
                  setIsExpanded(!isExpanded);
                  hasUserTouched.current = true;
                }}
                className="flex items-center justify-between px-3 py-2 bg-background/20 cursor-pointer hover:bg-background/40 transition-colors select-none"
              >
                <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-wider opacity-70">
                  <Sparkles size={12} className={cn("text-primary", isStreaming && "animate-pulse")} />
                  {isStreaming ? '✨ 思考中...' : '🧠 思考已完成'}
                </div>
                {isExpanded ? <ChevronUp size={12} className="opacity-50" /> : <ChevronDown size={12} className="opacity-50" />}
              </div>
              
              {isExpanded && (
                <div className="p-3 text-[11px] text-muted-foreground italic leading-relaxed border-t border-border/30 animate-in slide-in-from-top-1 duration-200">
                  {msg.think}
                </div>
              )}
            </div>
          )}

          {/* Main Content */}
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
          
          {/* Copy Button */}
          {msg.content && (
            <button 
              onClick={() => onCopy(msg.content, idx)}
              className="absolute -bottom-6 left-0 opacity-0 group-hover:opacity-100 transition-opacity flex items-center gap-1 text-[10px] text-muted-foreground hover:text-primary p-1"
            >
              {copiedIdx === idx ? <><Check size={10} /> 已复制</> : <><Copy size={10} /> 复制内容</>}
            </button>
          )}
        </div>
      )}
    </div>
  );
}
