import { useMemo } from 'react';
import type { AiSession } from '../../ai/sessionTypes';
import { cn } from '../../lib/utils';
import AppIcon from '../icons/AppIcon';
import AiSessionItem from './AiSessionItem';

interface AiSessionSidebarProps {
  sessions: AiSession[];
  currentSessionId: string | null;
  generatingSessionId: string | null;
  currentProjectId?: string;
  currentProjectName?: string;
  compact: boolean;
  onCreate: () => void;
  onSelect: (sessionId: string) => void;
  onRename: (sessionId: string, title: string) => void;
  onDelete: (sessionId: string) => void;
  onClose: () => void;
}

interface SessionGroupProps {
  title: string;
  sessions: AiSession[];
  currentSessionId: string | null;
  generatingSessionId: string | null;
  currentProjectId?: string;
  onSelect: (sessionId: string) => void;
  onRename: (sessionId: string, title: string) => void;
  onDelete: (sessionId: string) => void;
  emptyText?: string;
}

function SessionGroup({
  title,
  sessions,
  currentSessionId,
  generatingSessionId,
  currentProjectId,
  onSelect,
  onRename,
  onDelete,
  emptyText,
}: SessionGroupProps) {
  return (
    <section className="space-y-1.5">
      <div className="flex items-center justify-between px-2 text-[10px] font-bold uppercase tracking-[0.12em] text-muted-foreground">
        <span className="truncate">{title}</span>
        <span className="numeric-value ml-2 shrink-0 font-medium">{sessions.length}</span>
      </div>

      {sessions.length > 0 ? sessions.map(session => (
        <AiSessionItem
          key={session.id}
          session={session}
          isActive={session.id === currentSessionId}
          isGenerating={session.id === generatingSessionId}
          contextLabel={session.projectId
            ? session.projectId === currentProjectId ? '当前项目' : '其他项目'
            : '通用'}
          onSelect={onSelect}
          onRename={onRename}
          onDelete={onDelete}
        />
      )) : emptyText ? (
        <p className="px-3 py-2 text-[11px] leading-5 text-muted-foreground">
          {emptyText}
        </p>
      ) : null}
    </section>
  );
}

export default function AiSessionSidebar({
  sessions,
  currentSessionId,
  generatingSessionId,
  currentProjectId,
  currentProjectName,
  compact,
  onCreate,
  onSelect,
  onRename,
  onDelete,
  onClose,
}: AiSessionSidebarProps) {
  const sortedSessions = useMemo(
    () => [...sessions].sort((a, b) => b.updatedAt - a.updatedAt),
    [sessions],
  );
  const currentProjectSessions = currentProjectId
    ? sortedSessions.filter(session => session.projectId === currentProjectId)
    : [];
  const otherSessions = currentProjectId
    ? sortedSessions.filter(session => session.projectId !== currentProjectId)
    : sortedSessions;

  return (
    <aside
      className={cn(
        'flex h-full w-[216px] shrink-0 flex-col overflow-hidden bg-muted/40',
        compact && 'w-[min(82vw,260px)] bg-background shadow-xl',
      )}
      aria-label="AI 会话列表"
    >
      <div className="flex items-center gap-2 px-3 pb-2 pt-3">
        <button
          type="button"
          onClick={onCreate}
          className="flex h-10 min-w-0 flex-1 items-center justify-center gap-2 rounded-xl bg-card px-3 text-[13px] font-semibold text-foreground shadow-sm transition-colors hover:bg-primary-soft hover:text-primary"
        >
          <AppIcon name="add" size={16} />
          <span>新建会话</span>
        </button>
        {compact && (
          <button
            type="button"
            onClick={onClose}
            className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
            title="收起会话列表"
          >
            <AppIcon name="close" size={17} />
          </button>
        )}
      </div>

      <div className="min-h-0 flex-1 space-y-5 overflow-y-auto px-2 pb-4 pt-2">
        {currentProjectId && (
          <SessionGroup
            title={currentProjectName || '当前项目'}
            sessions={currentProjectSessions}
            currentSessionId={currentSessionId}
            generatingSessionId={generatingSessionId}
            currentProjectId={currentProjectId}
            onSelect={onSelect}
            onRename={onRename}
            onDelete={onDelete}
            emptyText="新建会话会自动关联当前项目。"
          />
        )}

        <SessionGroup
          title={currentProjectId ? '其他 / 通用会话' : '会话'}
          sessions={otherSessions}
          currentSessionId={currentSessionId}
          generatingSessionId={generatingSessionId}
          currentProjectId={currentProjectId}
          onSelect={onSelect}
          onRename={onRename}
          onDelete={onDelete}
        />
      </div>

      <div className="bg-background/50 px-4 py-3 text-[10px] leading-4 text-muted-foreground">
        会话记录保存在本机
      </div>
    </aside>
  );
}
