import { useState, useEffect, useCallback } from 'react';
import { invoke } from '@tauri-apps/api/core';

export function useModulePath(moduleId: string) {
  const [path, setPath] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  const refreshPath = useCallback(async () => {
    try {
      const p = await invoke<string | null>('get_module_path', { moduleId });
      setPath(p);
    } catch (e) {
      console.error(`Failed to get path for ${moduleId}:`, e);
    } finally {
      setIsLoading(false);
    }
  }, [moduleId]);

  const updatePath = async () => {
    try {
      const newPath = await invoke<string>('set_module_path', { moduleId });
      if (newPath) {
        setPath(newPath);
        return newPath;
      }
    } catch (e) {
      if (e !== "用户取消了选择") {
        console.error(`Failed to set path for ${moduleId}:`, e);
        throw e;
      }
    }
    return null;
  };

  useEffect(() => {
    refreshPath();
  }, [refreshPath]);

  return { path, isLoading, updatePath, refreshPath };
}
