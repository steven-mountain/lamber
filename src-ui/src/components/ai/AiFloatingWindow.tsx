import { useCallback, useEffect, useMemo, useState } from 'react';
import { Bot, X } from 'lucide-react';
import { getCurrentWindow, type Window as TauriWindow } from '@tauri-apps/api/window';
import AiChatPanel from './AiChatPanel';
import { useAiContextStore } from '../../store/useAiContextStore';

const AI_WINDOW_POSITION_KEY = 'lamber_ai_window_position';
const AI_CURRENT_VIEW_KEY = 'lamber_ai_current_view';

interface AiFloatingWindowProps {
  currentView?: string;
}

function isTauriRuntime() {
  return typeof window !== 'undefined' && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

export default function AiFloatingWindow({ currentView }: AiFloatingWindowProps) {
  const appWindow = useMemo<TauriWindow | null>(() => (
    isTauriRuntime() ? getCurrentWindow() : null
  ), []);
  const [assistantView, setAssistantView] = useState(() => (
    currentView || localStorage.getItem(AI_CURRENT_VIEW_KEY) || 'hub'
  ));

  const saveWindowPosition = useCallback(async () => {
    if (!appWindow) return;

    try {
      const position = await appWindow.outerPosition();
      const scaleFactor = await appWindow.scaleFactor();
      const logicalPosition = position.toLogical(scaleFactor);
      localStorage.setItem(AI_WINDOW_POSITION_KEY, JSON.stringify({
        x: logicalPosition.x,
        y: logicalPosition.y,
      }));
    } catch (error) {
      console.warn('Failed to save AI window position:', error);
    }
  }, [appWindow]);

  useEffect(() => {
    const previousHtmlBackground = document.documentElement.style.background;
    const previousBodyBackground = document.body.style.background;
    document.documentElement.style.background = 'transparent';
    document.body.style.background = 'transparent';

    let unlisten: (() => void) | undefined;
    if (appWindow) {
      appWindow.onCloseRequested(async () => {
        await saveWindowPosition();
      }).then((handler) => {
        unlisten = handler;
      });
    }

    return () => {
      document.documentElement.style.background = previousHtmlBackground;
      document.body.style.background = previousBodyBackground;
      unlisten?.();
    };
  }, [appWindow, saveWindowPosition]);

  useEffect(() => {
    if (currentView) {
      setAssistantView(currentView);
    }
  }, [currentView]);

  useEffect(() => {
    const handleStorage = (event: StorageEvent) => {
      if (event.key === AI_CURRENT_VIEW_KEY && event.newValue) {
        setAssistantView(event.newValue);
        useAiContextStore.getState().hydrateFromStorage();
      }
    };

    window.addEventListener('storage', handleStorage);
    return () => window.removeEventListener('storage', handleStorage);
  }, []);

  useEffect(() => {
    if (!appWindow) return;

    let unlisten: (() => void) | undefined;
    appWindow.listen<{ view?: string }>('lamber-ai-view-changed', (event) => {
      if (event.payload?.view) {
        setAssistantView(event.payload.view);
        useAiContextStore.getState().hydrateFromStorage();
      }
    }).then((handler) => {
      unlisten = handler;
    });

    return () => unlisten?.();
  }, [appWindow]);

  const closeWindow = async () => {
    await saveWindowPosition();
    if (appWindow) {
      try {
        await appWindow.destroy();
      } catch (error) {
        console.warn('Failed to destroy AI window, closing instead:', error);
        await appWindow.close();
      }
    } else {
      window.location.hash = '';
    }
  };

  return (
    <div className="h-screen w-screen overflow-hidden bg-transparent text-foreground">
      <div className="flex h-full w-full flex-col overflow-hidden rounded-[18px] border border-border bg-background">
        <div className="flex h-12 flex-shrink-0 items-center justify-between border-b border-border bg-card px-4 select-none">
          <div data-tauri-drag-region className="flex h-full flex-1 cursor-move items-center gap-2 text-sm font-semibold">
            <Bot size={18} className="text-primary" />
            Lamber AI 助手
          </div>

          <button
            type="button"
            onClick={closeWindow}
            className="flex h-8 w-8 items-center justify-center rounded-lg text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
            title="关闭"
          >
            <X size={18} />
          </button>
        </div>

        <AiChatPanel currentView={assistantView} />
      </div>
    </div>
  );
}
