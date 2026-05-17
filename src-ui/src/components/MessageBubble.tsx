import React, { useState, useRef, useEffect } from 'react';
import ReactMarkdown, { type Components } from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Sparkles, Check, Copy, ChevronDown, ChevronUp, Bot } from 'lucide-react';
import { clsx } from 'clsx';
import type { AiChatMessage } from '../ai/types';

interface MessageBubbleProps {
  msg: AiChatMessage;
  idx: number;
  isStreaming: boolean;
  onCopy: (text: string, idx: number) => void;
  copiedIdx: number | null;
}

const renderInlineBadges = (children: React.ReactNode) => {
  return React.Children.map(children, (child) => {
    if (typeof child !== 'string') return child;

    return child.split(/(\[系统内置\]|【系统外扩展】)/g).map((segment, segmentIndex) => {
      if (segment === '[系统内置]') {
        return (
          <span className="mx-1 inline-flex items-center gap-1 rounded-md border border-primary/20 bg-primary/10 px-1.5 py-0.5 text-xs font-bold text-primary">
            <Check size={12} strokeWidth={3} /> 系统内置
          </span>
        );
      }
      if (segment === '【系统外扩展】') {
        return (
          <span className="mx-1 inline-flex items-center gap-1 rounded-md border border-border bg-muted-foreground/10 px-1.5 py-0.5 text-[11px] font-medium italic text-muted-foreground">
            系统外扩展
          </span>
        );
      }
      return <React.Fragment key={`text-${segmentIndex}`}>{segment}</React.Fragment>;
    });
  });
};

const markdownComponents: Components = {
  h1: ({ children, ...props }) => (
    <h1 className="mb-4 mt-1 border-b border-border pb-2 text-xl font-bold leading-tight tracking-tight text-foreground" {...props}>
      {children}
    </h1>
  ),
  h2: ({ children, ...props }) => (
    <h2 className="mb-3 mt-5 text-lg font-bold leading-snug tracking-tight text-foreground first:mt-0" {...props}>
      {children}
    </h2>
  ),
  h3: ({ children, ...props }) => (
    <h3 className="mb-2 mt-4 text-base font-bold leading-snug text-foreground first:mt-0" {...props}>
      {children}
    </h3>
  ),
  p: ({ children, ...props }) => (
    <p className="my-3 leading-7 first:mt-0 last:mb-0" {...props}>
      {renderInlineBadges(children)}
    </p>
  ),
  ul: ({ children, ...props }) => (
    <ul className="my-3 list-disc space-y-1.5 pl-5 marker:text-muted-foreground" {...props}>
      {children}
    </ul>
  ),
  ol: ({ children, ...props }) => (
    <ol className="my-3 list-decimal space-y-1.5 pl-5 marker:font-semibold marker:text-muted-foreground" {...props}>
      {children}
    </ol>
  ),
  li: ({ children, ...props }) => (
    <li className="pl-1 leading-7" {...props}>
      {children}
    </li>
  ),
  blockquote: ({ children, ...props }) => (
    <blockquote className="my-4 border-l-4 border-primary/30 bg-muted/40 py-2 pl-4 pr-3 text-muted-foreground" {...props}>
      {children}
    </blockquote>
  ),
  table: ({ children }) => (
    <div className="my-4 max-w-full overflow-x-auto rounded-xl border border-border bg-background shadow-sm">
      <table className="w-max min-w-full border-collapse text-left text-sm">
        {children}
      </table>
    </div>
  ),
  thead: ({ children, ...props }) => (
    <thead className="bg-muted/80" {...props}>
      {children}
    </thead>
  ),
  tr: ({ children, ...props }) => (
    <tr className="border-b border-border last:border-b-0" {...props}>
      {children}
    </tr>
  ),
  th: ({ children, ...props }) => (
    <th className="whitespace-nowrap px-3 py-2.5 text-xs font-bold uppercase tracking-wide text-foreground" {...props}>
      {children}
    </th>
  ),
  td: ({ children, ...props }) => (
    <td className="whitespace-nowrap px-3 py-2.5 align-top leading-6 text-foreground/90" {...props}>
      {children}
    </td>
  ),
  code: ({ children, className, ...props }) => {
    const isCodeBlock = typeof className === 'string' && className.includes('language-');
    return (
      <code
        className={clsx(
          isCodeBlock
            ? 'block overflow-x-auto whitespace-pre rounded-none bg-transparent p-0 font-mono text-[12px] leading-6 text-slate-100'
            : 'rounded-md border border-border bg-muted px-1.5 py-0.5 font-mono text-[0.86em] text-primary',
          className
        )}
        {...props}
      >
        {children}
      </code>
    );
  },
  pre: ({ children, ...props }) => (
    <pre className="my-4 overflow-x-auto rounded-xl border border-slate-700 bg-slate-950 p-4 shadow-sm" {...props}>
      {children}
    </pre>
  ),
  a: ({ children, ...props }) => (
    <a className="font-medium text-primary underline underline-offset-4 hover:text-primary/80" target="_blank" rel="noreferrer" {...props}>
      {children}
    </a>
  ),
  hr: (props) => (
    <hr className="my-5 border-border" {...props} />
  ),
  strong: ({ children, ...props }) => (
    <strong className="font-bold text-foreground" {...props}>
      {children}
    </strong>
  )
};

const MessageBubble = ({ msg, idx, isStreaming, onCopy, copiedIdx }: MessageBubbleProps) => {
  const [isExpanded, setIsExpanded] = useState(true);
  const hasUserTouched = useRef(false);

  useEffect(() => {
    if (!isStreaming && msg.think && !hasUserTouched.current) {
      setIsExpanded(false);
    }
  }, [isStreaming, msg.think]);

  if (!msg.content?.trim() && !msg.think?.trim() && msg.role === 'assistant') {
    return null;
  }

  const handleHeaderClick = () => {
    hasUserTouched.current = true;
    setIsExpanded(!isExpanded);
  };

  return (
    <div
      className={clsx(
        'group flex flex-col gap-2 animate-in fade-in slide-in-from-bottom-2 duration-300',
        msg.role === 'user' ? 'max-w-[85%] self-end' : 'w-full max-w-full self-start'
      )}
    >
      <div className={clsx(
        'rounded-2xl border px-4 py-3 shadow-sm',
        msg.role === 'user'
          ? 'rounded-br-sm border-primary/20 bg-primary text-primary-foreground'
          : 'rounded-bl-sm border-border bg-card text-card-foreground'
      )}>
        {msg.think && (
          <div className="mb-3 overflow-hidden rounded-lg border border-border/50 bg-muted/30">
            <button
              onClick={handleHeaderClick}
              className="flex w-full items-center justify-between px-3 py-2 transition-colors hover:bg-muted/50"
            >
              <div className="flex items-center gap-2">
                {isStreaming && !msg.content ? (
                  <Sparkles className="h-3 w-3 animate-pulse text-primary" />
                ) : (
                  <Bot className="h-3 w-3 text-muted-foreground" />
                )}
                <span className="text-[11px] font-bold uppercase tracking-wider text-muted-foreground">
                  {isStreaming ? '思考中...' : '思考已完成'}
                </span>
              </div>
              {isExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>

            {isExpanded && (
              <div className="whitespace-pre-wrap border-t border-border/30 px-3 pb-3 pt-1 text-xs italic leading-relaxed text-muted-foreground/80">
                {msg.think}
              </div>
            )}
          </div>
        )}

        {msg.images && msg.images.length > 0 && (
          <div className={clsx('mb-3 grid gap-2', msg.images.length === 1 ? 'grid-cols-1' : 'grid-cols-2')}>
            {msg.images.map((image) => (
              <a
                key={image.id}
                href={image.dataUrl}
                target="_blank"
                rel="noreferrer"
                className="block overflow-hidden rounded-lg border border-primary-foreground/30 bg-black/10"
                title={image.name}
              >
                <img src={image.dataUrl} alt={image.name} className="max-h-44 w-full object-cover" />
              </a>
            ))}
          </div>
        )}

        {msg.content && (
          <div className="relative">
            {isStreaming ? (
              <div className="whitespace-pre-wrap break-words overflow-hidden text-sm leading-7">
                {String(msg.content)}
              </div>
            ) : msg.role === 'user' ? (
              <div className="whitespace-pre-wrap break-words text-sm leading-7 text-primary-foreground">
                {String(msg.content)}
              </div>
            ) : (
              <div className="min-w-0 max-w-none break-words text-[15px] leading-7 text-foreground">
                <ReactMarkdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
                  {String(msg.content || '')}
                </ReactMarkdown>
              </div>
            )}

            {!isStreaming && msg.content.trim() && (
              <button
                onClick={() => onCopy(msg.content, idx)}
                className="absolute -bottom-2 -right-2 rounded-md border border-border bg-background p-1.5 opacity-0 shadow-sm transition-opacity hover:bg-muted group-hover:opacity-100"
                title="复制内容"
              >
                {copiedIdx === idx ? <Check size={12} className="text-green-500" /> : <Copy size={12} className="text-muted-foreground" />}
              </button>
            )}
          </div>
        )}
      </div>

      <span className="px-1 text-[10px] text-muted-foreground">
        {msg.role === 'user' ? '您' : 'Lamber 智能顾问'}
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
