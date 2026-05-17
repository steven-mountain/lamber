import { useEffect, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import { Bot } from 'lucide-react';
import { emit, emitTo } from '@tauri-apps/api/event';
import { WebviewWindow } from '@tauri-apps/api/webviewWindow';
import { AI_CONTEXT_REFRESH_REQUEST_EVENT } from '../../store/useAiContextStore';

const AI_ASSISTANT_LABEL = 'ai-assistant';
const AI_LAUNCHER_POSITION_KEY = 'lamber_ai_launcher_position';
const AI_WINDOW_POSITION_KEY = 'lamber_ai_window_position';
const AI_CURRENT_VIEW_KEY = 'lamber_ai_current_view';
const BUTTON_SIZE = 56;
const SCREEN_MARGIN = 16;

interface AiFloatingLauncherProps {
  currentView: string;
}

interface FloatingPosition {
  x: number;
  y: number;
}

interface DragState {
  pointerId: number;
  startX: number;
  startY: number;
  originX: number;
  originY: number;
  lastX: number;
  lastY: number;
  moved: boolean;
}

function clampLauncherPosition(position: FloatingPosition): FloatingPosition {
  const maxX = Math.max(window.innerWidth - BUTTON_SIZE - SCREEN_MARGIN, SCREEN_MARGIN);
  const maxY = Math.max(window.innerHeight - BUTTON_SIZE - SCREEN_MARGIN, SCREEN_MARGIN);

  return {
    x: Math.min(Math.max(position.x, SCREEN_MARGIN), maxX),
    y: Math.min(Math.max(position.y, SCREEN_MARGIN), maxY),
  };
}

function getDefaultLauncherPosition(): FloatingPosition {
  return {
    x: Math.max(window.innerWidth - BUTTON_SIZE - 40, SCREEN_MARGIN),
    y: Math.max(window.innerHeight - BUTTON_SIZE - 40, SCREEN_MARGIN),
  };
}

function readPosition(key: string): FloatingPosition | null {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<FloatingPosition>;
    if (typeof parsed.x !== 'number' || typeof parsed.y !== 'number') return null;
    return parsed as FloatingPosition;
  } catch (error) {
    console.warn(`Failed to read ${key}:`, error);
    return null;
  }
}

function savePosition(key: string, position: FloatingPosition) {
  localStorage.setItem(key, JSON.stringify(position));
}

function isTauriRuntime() {
  return typeof window !== 'undefined' && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

export default function AiFloatingLauncher({ currentView }: AiFloatingLauncherProps) {
  const [position, setPosition] = useState<FloatingPosition>(() => (
    clampLauncherPosition(readPosition(AI_LAUNCHER_POSITION_KEY) || getDefaultLauncherPosition())
  ));
  const dragStateRef = useRef<DragState | null>(null);

  useEffect(() => {
    const handleResize = () => {
      setPosition((current) => {
        const nextPosition = clampLauncherPosition(current);
        savePosition(AI_LAUNCHER_POSITION_KEY, nextPosition);
        return nextPosition;
      });
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const openAiWindow = async () => {
    localStorage.setItem(AI_CURRENT_VIEW_KEY, currentView);

    if (!isTauriRuntime()) {
      window.location.hash = `#/ai-assistant?view=${encodeURIComponent(currentView)}`;
      return;
    }

    const existing = await WebviewWindow.getByLabel(AI_ASSISTANT_LABEL);
    if (existing) {
      await existing.show();
      await existing.setFocus();
      await emitTo(AI_ASSISTANT_LABEL, 'lamber-ai-view-changed', { view: currentView });
      await emit(AI_CONTEXT_REFRESH_REQUEST_EVENT, { view: currentView });
      return;
    }

    const savedWindowPosition = readPosition(AI_WINDOW_POSITION_KEY);
    const aiWindow = new WebviewWindow(AI_ASSISTANT_LABEL, {
      url: `/#/ai-assistant?view=${encodeURIComponent(currentView)}`,
      title: 'Lamber AI 助手',
      width: 420,
      height: 680,
      minWidth: 360,
      minHeight: 480,
      decorations: false,
      transparent: true,
      backgroundColor: [0, 0, 0, 0],
      alwaysOnTop: true,
      resizable: true,
      shadow: false,
      skipTaskbar: false,
      center: false,
      preventOverflow: true,
      ...(savedWindowPosition ? { x: savedWindowPosition.x, y: savedWindowPosition.y } : {}),
    });

    aiWindow.once('tauri://created', () => {
      console.log('AI assistant window created');
      emit(AI_CONTEXT_REFRESH_REQUEST_EVENT, { view: currentView }).catch((error) => {
        console.warn('Failed to request AI context refresh:', error);
      });
    });

    aiWindow.once('tauri://error', (event) => {
      console.error('AI assistant window error', event);
    });
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLButtonElement>) => {
    if (event.button !== 0) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    dragStateRef.current = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      originX: position.x,
      originY: position.y,
      lastX: position.x,
      lastY: position.y,
      moved: false,
    };
  };

  const handlePointerMove = (event: ReactPointerEvent<HTMLButtonElement>) => {
    const dragState = dragStateRef.current;
    if (!dragState || dragState.pointerId !== event.pointerId) return;

    const deltaX = event.clientX - dragState.startX;
    const deltaY = event.clientY - dragState.startY;
    if (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3) {
      dragState.moved = true;
    }

    const nextPosition = clampLauncherPosition({
      x: dragState.originX + deltaX,
      y: dragState.originY + deltaY,
    });
    dragState.lastX = nextPosition.x;
    dragState.lastY = nextPosition.y;
    setPosition(nextPosition);
  };

  const handlePointerUp = async (event: ReactPointerEvent<HTMLButtonElement>) => {
    const dragState = dragStateRef.current;
    if (!dragState || dragState.pointerId !== event.pointerId) return;

    event.currentTarget.releasePointerCapture(event.pointerId);
    dragStateRef.current = null;

    const nextPosition = { x: dragState.lastX, y: dragState.lastY };
    savePosition(AI_LAUNCHER_POSITION_KEY, nextPosition);

    if (!dragState.moved) {
      await openAiWindow();
    }
  };

  return (
    <button
      type="button"
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onKeyDown={(event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          openAiWindow();
        }
      }}
      className="fixed z-50 flex h-14 w-14 touch-none items-center justify-center rounded-full bg-primary text-primary-foreground shadow-xl transition-transform hover:scale-105 active:scale-95"
      style={{ left: `${position.x}px`, top: `${position.y}px` }}
      title="打开 AI 助手"
      aria-label="打开 AI 助手"
    >
      <Bot size={28} />
    </button>
  );
}
