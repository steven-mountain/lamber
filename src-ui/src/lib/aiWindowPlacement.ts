export const AI_WINDOW_POSITION_KEY = 'lamber_ai_window_position';
export interface Point { x: number; y: number }
export interface Size { width: number; height: number }
export interface ScreenArea { workArea: { position: Point; size: Size } }
export interface SavedWindowPosition extends Point { version?: 2 }

export function parseWindowPosition(raw: string | null): SavedWindowPosition | null {
  try {
    const value = JSON.parse(raw || 'null');
    if (!value || !Number.isFinite(value.x) || !Number.isFinite(value.y)
      || (value.version !== undefined && value.version !== 2)) return null;
    return value;
  } catch { return null; }
}

/** All geometry is physical pixels. Negative coordinates are valid on connected monitors. */
export function placeAiWindow(position: Point, size: Size, screens: ScreenArea[], preferred: ScreenArea | null) {
  if (!screens.length) throw new Error('无法读取显示器范围，请重新点击 AI 按钮。');
  const fits = screens.some(({ workArea: { position: origin, size: area } }) =>
    position.x >= origin.x && position.y >= origin.y
    && position.x + size.width <= origin.x + area.width
    && position.y + size.height <= origin.y + area.height);
  if (fits) return { position, size };
  const { workArea } = preferred || screens[0];
  const nextSize = { width: Math.min(size.width, workArea.size.width), height: Math.min(size.height, workArea.size.height) };
  return {
    position: {
      x: Math.round(workArea.position.x + (workArea.size.width - nextSize.width) / 2),
      y: Math.round(workArea.position.y + (workArea.size.height - nextSize.height) / 2),
    },
    size: nextSize,
  };
}
