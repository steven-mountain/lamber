import { ApprovalReview } from "./ApprovalReview";
import { useEffect, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { getCurrentWebviewWindow } from "@tauri-apps/api/webviewWindow";
import { invoke } from "@tauri-apps/api/core";
import { remainingSeconds, type ApprovalPrompt } from '../../ai/approvalReview';

export const AI_APPROVAL_REQUEST_EVENT = "ai://approval-request";

export default function AgentApprovalDialog() {
  const [queue, setQueue] = useState<ApprovalPrompt[]>([]);
  useEffect(() => {
    let unlisten: (() => void) | undefined;
    let unlistenSettled: (() => void) | undefined;
    let disposed = false;
    void listen<{requestId:string}>('ai://approval-settled', event => {
      setQueue(items => items.filter(item => item.requestId !== event.payload.requestId));
    }).then(handler => { if (disposed) handler(); else unlistenSettled = handler; })
      .catch(error => console.warn('Failed to listen for approval outcomes:', error));
    listen<ApprovalPrompt>(AI_APPROVAL_REQUEST_EVENT, event => {
      if (remainingSeconds(event.payload.expiresAt) > 0) setQueue(prev => prev.some(item => item.requestId === event.payload.requestId) ? prev : [...prev,event.payload]);
    }, { target: getCurrentWebviewWindow().label }).then(handler => { if (disposed) handler(); else unlisten = handler; })
      .catch(error => console.warn('Failed to listen for approval requests:',error));
    return () => { disposed = true; unlisten?.(); unlistenSettled?.(); };
  }, []);
  const current = queue[0];
  return current ? <ApprovalReview key={current.requestId} current={current} resolve={args => invoke("ai_resolve_approval", args)} onSettled={id => setQueue(items => items.filter(item => item.requestId !== id))} /> : null;
}
