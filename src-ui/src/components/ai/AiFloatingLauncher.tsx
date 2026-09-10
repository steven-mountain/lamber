import { listen } from '@tauri-apps/api/event';
import { useEffect, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import { openAiAssistantWindow } from '../../services/aiAssistantWindow';
import AppIcon from '../icons/AppIcon';

const AI_LAUNCHER_POSITION_KEY = 'lamber_ai_launcher_position';
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

function getLauncherTransform(position: FloatingPosition) {
  return `translate3d(${position.x}px, ${position.y}px, 0)`;
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

export default function AiFloatingLauncher({ currentView }: AiFloatingLauncherProps) {
  const [openError, setOpenError] = useState('');
  const [opening, setOpening] = useState(false);
  const [position, setPosition] = useState<FloatingPosition>(() => (
    clampLauncherPosition(readPosition(AI_LAUNCHER_POSITION_KEY) || getDefaultLauncherPosition())
  ));
  const launcherRef = useRef<HTMLButtonElement | null>(null);
  const positionRef = useRef(position);
  const dragStateRef = useRef<DragState | null>(null);
  const suppressClickRef = useRef(false);
  const queuedPositionRef = useRef<FloatingPosition | null>(null);
  const animationFrameRef = useRef<number | null>(null);

  const renderLauncherPosition = (nextPosition: FloatingPosition) => {
    if (!launcherRef.current) return;
    launcherRef.current.style.transform = getLauncherTransform(nextPosition);
  };

  const renderLauncherPositionImmediately = (nextPosition: FloatingPosition) => {
    if (animationFrameRef.current !== null) {
      window.cancelAnimationFrame(animationFrameRef.current);
      animationFrameRef.current = null;
    }
    queuedPositionRef.current = null;
    renderLauncherPosition(nextPosition);
  };

  const scheduleLauncherPositionRender = (nextPosition: FloatingPosition) => {
    queuedPositionRef.current = nextPosition;
    if (animationFrameRef.current !== null) return;

    animationFrameRef.current = window.requestAnimationFrame(() => {
      animationFrameRef.current = null;
      const queuedPosition = queuedPositionRef.current;
      queuedPositionRef.current = null;
      if (queuedPosition) {
        renderLauncherPosition(queuedPosition);
      }
    });
  };

  useEffect(() => {
    if (!('__TAURI_INTERNALS__' in window)) return;
    let disposed = false; let stop: (() => void) | undefined;
    void listen<string>('lamber-ai-startup-error', event => setOpenError(`AI 重启失败：${event.payload}`))
      .then(unlisten => { if (disposed) unlisten(); else stop = unlisten; }).catch(console.error);
    return () => { disposed = true; stop?.(); };
  }, []);

  useEffect(() => {
    const handleResize = () => {
      const nextPosition = clampLauncherPosition(positionRef.current);
      positionRef.current = nextPosition;
      setPosition(nextPosition);
      if (animationFrameRef.current !== null) {
        window.cancelAnimationFrame(animationFrameRef.current);
        animationFrameRef.current = null;
      }
      queuedPositionRef.current = null;
      renderLauncherPosition(nextPosition);
      savePosition(AI_LAUNCHER_POSITION_KEY, nextPosition);
    };

    window.addEventListener('resize', handleResize);
    return () => {
      window.removeEventListener('resize', handleResize);
      if (animationFrameRef.current !== null) {
        window.cancelAnimationFrame(animationFrameRef.current);
      }
    };
  }, []);

  const openAiWindow = async () => {
    setOpenError(''); setOpening(true);
    try { await openAiAssistantWindow(currentView); }
    catch (error) { setOpenError(`AI 窗口打开失败：${String(error)}`); }
    finally { setOpening(false); }
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLButtonElement>) => {
    if (event.button !== 0) return;
    suppressClickRef.current = false;
    event.currentTarget.setPointerCapture(event.pointerId);
    const currentPosition = positionRef.current;
    dragStateRef.current = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      originX: currentPosition.x,
      originY: currentPosition.y,
      lastX: currentPosition.x,
      lastY: currentPosition.y,
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
    positionRef.current = nextPosition;
    scheduleLauncherPositionRender(nextPosition);
  };

  const finishPointerInteraction = (
    event: ReactPointerEvent<HTMLButtonElement>,
    cancelled: boolean
  ) => {
    const dragState = dragStateRef.current;
    if (!dragState || dragState.pointerId !== event.pointerId) return;

    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
    dragStateRef.current = null;

    const nextPosition = { x: dragState.lastX, y: dragState.lastY };
    positionRef.current = nextPosition;
    setPosition(nextPosition);
    renderLauncherPositionImmediately(nextPosition);
    savePosition(AI_LAUNCHER_POSITION_KEY, nextPosition);

    suppressClickRef.current = cancelled || dragState.moved;
  };

  return (
    <>
    {openError && <div role="alert" className="fixed bottom-24 right-4 z-50 max-w-sm rounded-lg bg-card p-4 text-caption text-destructive shadow-md">
      {openError}
      <button type="button" className="ml-2 rounded-md bg-muted px-3 py-2 text-foreground" onClick={() => void openAiWindow()}>重试</button>
    </div>}
    <button
      ref={launcherRef}
      type="button"
      aria-busy={opening}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={(event) => {
        finishPointerInteraction(event, false);
      }}
      onPointerCancel={(event) => {
        finishPointerInteraction(event, true);
      }}
      onClick={() => {
        if (suppressClickRef.current) { suppressClickRef.current = false; return; }
        void openAiWindow();
      }}
      className="group fixed left-0 top-0 z-50 h-14 w-14 touch-none rounded-full focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-200"
      style={{ transform: getLauncherTransform(positionRef.current || position), willChange: 'transform' }}
      title="打开 AI 助手"
      aria-label="打开 AI 助手"
    >
      <span className="flex h-full w-full items-center justify-center rounded-full border border-slate-200 bg-white text-blue-600 shadow-md transition-[background-color,color,box-shadow,transform] duration-150 group-hover:scale-105 group-hover:bg-blue-600 group-hover:text-white group-hover:shadow-lg group-active:scale-95">
        <AppIcon name={opening ? 'loading' : 'ai'} size={28} />
      </span>
    </button>
    </>
  );
}
