import { useState, useCallback, useRef, useEffect } from 'react';

/**
 * Industrial-grade Streaming Cognitive Parser Hook
 * Implements a state machine to separate "Thinking" (DeepSeek-style) from "Content".
 * Includes recovery logic for malformed tags and cleanup mechanisms.
 */
export function useStreamingParser() {
  const [normalText, setNormalText] = useState("");
  const [thinkText, setThinkText] = useState("");
  const [isInsideThink, setIsInsideThink] = useState(false);
  const [isStreaming, setIsStreaming] = useState(false);

  // Internal semantic buffers (The "Reservoir")
  const normalBufferRef = useRef("");
  const thinkBufferRef = useRef("");
  const internalRawBufferRef = useRef(""); 
  const stateRef = useRef<{ isInsideThink: boolean }>({ isInsideThink: false });
  const isStreamingRef = useRef(false);
  
  // Render Scheduler (The "Gating")
  const renderTimerRef = useRef<NodeJS.Timeout | null>(null);

  const flushToState = useCallback(() => {
    setNormalText(normalBufferRef.current);
    setThinkText(thinkBufferRef.current);
  }, []);

  const reset = useCallback(() => {
    setNormalText("");
    setThinkText("");
    setIsInsideThink(false);
    setIsStreaming(false);
    isStreamingRef.current = false;
    normalBufferRef.current = "";
    thinkBufferRef.current = "";
    internalRawBufferRef.current = "";
    stateRef.current.isInsideThink = false;
    
    if (renderTimerRef.current) {
      clearInterval(renderTimerRef.current);
      renderTimerRef.current = null;
    }
  }, []);

  /**
   * Process a new chunk of text from the stream
   * High-speed accumulation without triggering re-render
   */
  const parseChunk = useCallback((chunk: string) => {
    if (!chunk) return;
    if (!isStreamingRef.current) {
      isStreamingRef.current = true;
      setIsStreaming(true);
    }
    
    // 1. Initialize Scheduler if not running (80ms = ~12FPS)
    if (!renderTimerRef.current) {
      renderTimerRef.current = setInterval(() => {
        flushToState();
      }, 80);
    }

    // 2. Accumulate in Raw Buffer
    let current = internalRawBufferRef.current + chunk;
    
    // 3. Process tags incrementally into semantic buffers
    if (!stateRef.current.isInsideThink) {
      const startIdx = current.indexOf("<think>");
      if (startIdx !== -1) {
        normalBufferRef.current += current.substring(0, startIdx);
        stateRef.current.isInsideThink = true;
        setIsInsideThink(true);
        internalRawBufferRef.current = current.substring(startIdx + 7);
      } else {
        const lastBracket = current.lastIndexOf("<");
        if (lastBracket !== -1 && lastBracket > current.length - 7) {
          normalBufferRef.current += current.substring(0, lastBracket);
          internalRawBufferRef.current = current.substring(lastBracket);
        } else {
          normalBufferRef.current += current;
          internalRawBufferRef.current = "";
        }
      }
    } else {
      const endIdx = current.indexOf("</think>");
      if (endIdx !== -1) {
        thinkBufferRef.current += current.substring(0, endIdx);
        stateRef.current.isInsideThink = false;
        setIsInsideThink(false);
        internalRawBufferRef.current = current.substring(endIdx + 8);
      } else {
        const lastBracket = current.lastIndexOf("<");
        if (lastBracket !== -1 && lastBracket > current.length - 8) {
          thinkBufferRef.current += current.substring(0, lastBracket);
          internalRawBufferRef.current = current.substring(lastBracket);
        } else {
          thinkBufferRef.current += current;
          internalRawBufferRef.current = "";
        }
      }
    }
  }, [flushToState]);

  /**
   * Parser Recovery & Final Flush
   */
  const finalize = useCallback(() => {
    // 1. Stop scheduler
    if (renderTimerRef.current) {
      clearInterval(renderTimerRef.current);
      renderTimerRef.current = null;
    }

    // 2. Clear remaining raw buffer
    if (internalRawBufferRef.current) {
      if (stateRef.current.isInsideThink) {
        thinkBufferRef.current += internalRawBufferRef.current;
      } else {
        normalBufferRef.current += internalRawBufferRef.current;
      }
      internalRawBufferRef.current = "";
    }
    
    // 3. Recovery: Force close any open tags
    if (stateRef.current.isInsideThink) {
      console.warn("Parser Recovery: Stream ended without closing </think> tag.");
      stateRef.current.isInsideThink = false;
      setIsInsideThink(false);
    }
    
    // 4. Critical: Final state sync and status update
    flushToState();
    isStreamingRef.current = false;
    setIsStreaming(false);
  }, [flushToState]);

  /**
   * Emergency Stop: Just kill the timer and status
   * Used for errors/aborts where we don't want to flush residual data
   */
  const stop = useCallback(() => {
    if (renderTimerRef.current) {
      clearInterval(renderTimerRef.current);
      renderTimerRef.current = null;
    }
    isStreamingRef.current = false;
    setIsStreaming(false);
  }, []);

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      if (renderTimerRef.current) clearInterval(renderTimerRef.current);
    };
  }, []);

  return {
    normalText,
    thinkText,
    isInsideThink,
    isStreaming,
    parseChunk,
    finalize,
    reset,
    stop
  };
}
