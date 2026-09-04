import { useEffect, useRef, useState } from 'react';
import type { AiSession } from '../../ai/sessionTypes';
import { cn } from '../../lib/utils';
import AppIcon from '../icons/AppIcon';

interface AiSessionItemProps {
  session: AiSession;
  isActive: boolean;
  isGenerating: boolean;
  contextLabel?: string;
  onSelect: (sessionId: string) => void;
  onRename: (sessionId: string, title: string) => void;
  onDelete: (sessionId: string) => void;
}

function formatSessionTime(timestamp: number) {
  const date = new Date(timestamp);
  const now = new Date();
  const isToday = date.toDateString() === now.toDateString();
  return isToday
    ? date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    : date.toLocaleDateString([], { month: 'numeric', day: 'numeric' });
}

export default function AiSessionItem({
  session,
  isActive,
  isGenerating,
  contextLabel,
  onSelect,
  onRename,
  onDelete,
}: AiSessionItemProps) {
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [isRenaming, setIsRenaming] = useState(false);
  const [draftTitle, setDraftTitle] = useState(session.title);
  const itemRef = useRef<HTMLDivElement | null>(null);
  const renameInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    setDraftTitle(session.title);
  }, [session.title]);

  useEffect(() => {
    if (!isMenuOpen) return;

    const handlePointerDown = (event: PointerEvent) => {
      if (!itemRef.current?.contains(event.target as Node)) {
        setIsMenuOpen(false);
      }
    };
    document.addEventListener('pointerdown', handlePointerDown);
    return () => document.removeEventListener('pointerdown', handlePointerDown);
  }, [isMenuOpen]);

  useEffect(() => {
    if (!isRenaming) return;
    renameInputRef.current?.focus();
    renameInputRef.current?.select();
  }, [isRenaming]);

  const commitRename = () => {
    const title = draftTitle.trim();
    if (title && title !== session.title) {
      onRename(session.id, title);
    } else {
      setDraftTitle(session.title);
    }
    setIsRenaming(false);
  };

  const runMenuAction = (action: () => void) => {
    setIsMenuOpen(false);
    action();
  };

  return (
    <div
      ref={itemRef}
      className={cn(
        'group relative w-full rounded-xl transition-colors',
        isActive
          ? 'bg-card text-foreground shadow-sm'
          : 'text-secondary-foreground hover:bg-card/60 hover:text-foreground',
      )}
    >
      <div
        role={isRenaming ? undefined : 'button'}
        tabIndex={isRenaming ? -1 : 0}
        aria-current={isActive ? 'page' : undefined}
        onClick={() => {
          if (!isRenaming) onSelect(session.id);
        }}
        onKeyDown={event => {
          if (!isRenaming && (event.key === 'Enter' || event.key === ' ')) {
            event.preventDefault();
            onSelect(session.id);
          }
        }}
        className="flex w-full items-start gap-2.5 rounded-xl px-3 py-2.5 pr-9 text-left"
        title={session.title}
      >
        <span
          className={cn(
            'mt-0.5 flex h-7 w-7 shrink-0 items-center justify-center rounded-lg transition-colors',
            isActive ? 'bg-primary-soft text-primary' : 'bg-background/70 text-muted-foreground',
          )}
        >
          <AppIcon name="aiMessage" size={14} />
        </span>

        <span className="min-w-0 flex-1">
          {isRenaming ? (
            <input
              ref={renameInputRef}
              value={draftTitle}
              maxLength={80}
              onClick={event => event.stopPropagation()}
              onChange={event => setDraftTitle(event.target.value)}
              onBlur={commitRename}
              onKeyDown={event => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  commitRename();
                } else if (event.key === 'Escape') {
                  event.preventDefault();
                  setDraftTitle(session.title);
                  setIsRenaming(false);
                }
              }}
              className="block h-5 w-full rounded-md bg-muted/70 px-1.5 text-[13px] font-semibold leading-5 text-foreground outline-none ring-2 ring-primary/20"
              aria-label="会话名称"
            />
          ) : (
            <span className="block truncate text-[13px] font-semibold leading-5">
              {session.title}
            </span>
          )}
          <span className="mt-0.5 flex items-center gap-1.5 text-[10px] leading-4 text-muted-foreground">
            {isGenerating && (
              <span className="h-1.5 w-1.5 shrink-0 animate-pulse rounded-full bg-primary" />
            )}
            <span className="truncate">{isGenerating ? '正在生成' : contextLabel}</span>
            <span aria-hidden="true">·</span>
            <span className="numeric-value shrink-0">{formatSessionTime(session.updatedAt)}</span>
          </span>
        </span>
      </div>

      {!isRenaming && (
        <button
          type="button"
          onClick={event => {
            event.stopPropagation();
            setIsMenuOpen(open => !open);
          }}
          className={cn(
            'absolute right-1.5 top-2 flex h-7 w-7 items-center justify-center rounded-lg text-muted-foreground transition-all hover:bg-muted hover:text-foreground focus:opacity-100',
            isMenuOpen || isActive ? 'opacity-100' : 'opacity-0 group-hover:opacity-100',
          )}
          aria-label={`管理会话 ${session.title}`}
          title="会话操作"
        >
          <AppIcon name="more" size={16} />
        </button>
      )}

      {isMenuOpen && (
        <div className="absolute right-1.5 top-9 z-30 w-44 rounded-xl bg-card p-1.5 text-[12px] font-medium text-foreground shadow-xl">
          <button
            type="button"
            onClick={() => runMenuAction(() => setIsRenaming(true))}
            className="flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-left transition-colors hover:bg-muted"
          >
            <AppIcon name="edit" size={14} />
            重命名
          </button>

          <button
            type="button"
            onClick={() => runMenuAction(() => onDelete(session.id))}
            className="flex w-full items-center gap-2 rounded-lg px-2.5 py-2 text-left text-destructive transition-colors hover:bg-destructive-soft"
          >
            <AppIcon name="delete" size={14} />
            删除会话
          </button>
        </div>
      )}
    </div>
  );
}
