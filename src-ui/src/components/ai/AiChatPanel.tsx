import TechItemsCard from './TechItemsCard';
import StructureReverseCard from './StructureReverseCard';
import { structureReverseIntent, structureReversePrompt } from '../../ai/structureReverseIntent';
import { appReceiptContext } from '../../ai/appReceiptContext';
import InquiryCard from './InquiryCard';
import { templateListIntent, templateListPrompt, techProposal } from '../../ai/templateListIntent';
import DocumentGenerationCards from './DocumentGenerationCards';
import { documentTemplateRequests, documentGenerationPrompt } from '../../ai/documentGenerationIntent';
import { GENERAL_SESSION_LABEL, sessionScopePrompt } from '../../ai/sessionScopePolicy';
import { demandImageCompletionPrompt, isDemandFormRequest, wantsDemandImageCompletion } from '../../ai/demandFormIntent';
import AiSessionProjectPicker from './AiSessionProjectPicker';
import DemandImageCompletionCards from "./DemandImageCompletionCards";
import TemplateImageCard from './TemplateImageCard';
import { wantsTemplateImages, templateImagePrompt } from '../../ai/templateImageIntent';
import { useEffect, useRef, useState } from 'react';
import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
import { SYSTEM_PROMPT_KNOWLEDGE } from '../../lib/knowledgeBase';
import {
  AI_CONTEXT_REFRESH_REQUEST_EVENT,
  AI_CONTEXT_UPDATED_EVENT,
  useAiContextStore,
} from '../../store/useAiContextStore';
import AiAgentSettingsCard from '../settings/AiAgentSettingsCard';
import { getAiAgentSettings, type AiAgentSettings } from '../../ai/agentSettings';
import { PromptRenderer } from '../../ai/PromptRenderer';
import { DshRuntime } from '../../ai/DshRuntime';
import type { AiChatMessage, AiImageAttachment, PromptAST, PromptRule } from '../../ai/types';
import { buildAiChatContext } from '../../ai/context/buildAiChatContext';
import { loadAiTemplateAsset } from '../../services/aiProjectContextService';
import {
  AI_TEMPLATE_ASSET_SELECTED_EVENT,
  AI_TEMPLATE_ASSET_SELECTED_STORAGE_KEY,
  parseTemplateAssetSelection,
  type AiTemplateAssetSelection,
} from '../../ai/templateAssetSelection';
import MessageBubble from '../MessageBubble';
import AiInputBox from './AiInputBox';
import AiSessionSidebar from './AiSessionSidebar';
import { AI_CONTEXT_KEY, getAiContextScope } from '../../utils/aiContextKeys';
import AppIcon, { type AppIconName } from '../icons/AppIcon';
import { useAiSessionStore } from '../../store/useAiSessionStore';
import { readStoredCurrentProject, useProjectStore } from '../../store/useProjectStore';

interface AiChatPanelProps {
  currentView?: string;
}

const EMPTY_MESSAGES: AiChatMessage[] = [];
const SESSION_SIDEBAR_BREAKPOINT = 680;

function isTauriRuntime() {
  return typeof window !== 'undefined' && Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);
}

function getCoreContextKey(view: string) {
  if (view === 'ict' || view === 'ict_lifecycle') return AI_CONTEXT_KEY.ICT_CORE;
  return view;
}

function getAiContextView(view: string) {
  if (view === 'ict_lifecycle') return 'ict';
  return view;
}

function formatLastUpdated(timestamp?: number) {
  if (!timestamp) return '--';
  return new Date(timestamp).toLocaleString();
}

export default function AiChatPanel({ currentView = 'hub' }: AiChatPanelProps) {
  const [pendingReceipts, setPendingReceipts] = useState<{ sessionId: string; content: string }[]>([]);
  const [showSettings, setShowSettings] = useState(false);
  const [input, setInput] = useState('');
  const [images, setImages] = useState<AiImageAttachment[]>([]);
  const [isTyping, setIsTyping] = useState(false);
  const [streamingSessionId, setStreamingSessionId] = useState<string | null>(null);
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null);
  const [isCompactLayout, setIsCompactLayout] = useState(() => (
    typeof window !== 'undefined' ? window.innerWidth < SESSION_SIDEBAR_BREAKPOINT : false
  ));
  const [isSidebarOpen, setIsSidebarOpen] = useState(() => (
    typeof window !== 'undefined' ? window.innerWidth >= SESSION_SIDEBAR_BREAKPOINT : true
  ));

  const sessions = useAiSessionStore(state => state.sessions);
  const currentSessionId = useAiSessionStore(state => state.currentSessionId);
  const createSession = useAiSessionStore(state => state.createSession);
  const ensureActiveSession = useAiSessionStore(state => state.ensureActiveSession);
  const selectSession = useAiSessionStore(state => state.selectSession);
  const deleteSession = useAiSessionStore(state => state.deleteSession);
  const appendMessages = useAiSessionStore(state => state.appendMessages);
  const updateLastAssistantMessage = useAiSessionStore(state => state.updateLastAssistantMessage);
  const resetSessionMessages = useAiSessionStore(state => state.resetSessionMessages);
  const setSessionTitle = useAiSessionStore(state => state.setSessionTitle);
  const flushSessionPersistence = useAiSessionStore(state => state.flushPersistence);
  const currentProject = useProjectStore(state => state.currentProject);
  const currentSession = sessions.find(session => session.id === currentSessionId);
  const messages = currentSession?.messages ?? EMPTY_MESSAGES;
  const listIntent = templateListIntent(messages);
  const reverseIntent = structureReverseIntent(messages);
  const listProposal = techProposal(messages);
  const demandImagesRequested = wantsDemandImageCompletion(messages);
  const latestImageRequest = [...messages].reverse().find(message => message.role === 'user')?.content || '';
  const [imageCardSession, setImageCardSession] = useState<string | null>(null);
  const savedImagesRequested = wantsTemplateImages(latestImageRequest) || imageCardSession === currentSessionId;
  const documentTemplateIds = documentTemplateRequests(messages);

  const dshRuntime = useRef(new DshRuntime());
  const [bindingState, setBindingState] = useState<{ sessionId: string; binding: { projectId: string | null; workspaceId: string; projectName?: string | null } | null; error?: string } | null>(null);
  const [bindingVersion, setBindingVersion] = useState(0);
  const bindingReady = bindingState?.sessionId === currentSessionId && Boolean(bindingState?.binding);
  useEffect(() => {
    if (!currentSessionId) return;
    let active = true;
    invoke<{ projectId: string | null; workspaceId: string; projectName?: string | null } | null>('ai_get_session_binding', { sessionId: currentSessionId })
      .then(binding => { if (active) setBindingState({ sessionId: currentSessionId, binding }); })
      .catch(error => { if (active) setBindingState({ sessionId: currentSessionId, binding: null, error: String(error) }); });
    return () => { active = false; };
  }, [currentSessionId, bindingVersion]);
  const chooseBinding = async (projectId: string | null) => {
    const fresh = currentSession && !currentSession.harnessSessionId && !currentSession.messages.some(message => message.role === 'user');
    // Historical contexts always get a fresh identity; unused placeholders can be completed.
    const sessionId = fresh ? currentSession.id : crypto.randomUUID();
    const binding = await invoke<{ projectId: string | null; workspaceId: string; projectName?: string | null }>('ai_bind_session_to_project', { sessionId, projectId });
    if (fresh) {
      useAiSessionStore.getState().setSessionProject(sessionId, projectId ?? undefined);
      selectSession(sessionId);
    } else createSession(projectId ?? undefined, sessionId);
    setBindingVersion(version => version + 1);
    setBindingState({ sessionId, binding });
    setImages([]);
    if (isCompactLayout) setIsSidebarOpen(false);
  };
  const [dshSettings, setDshSettings] = useState<AiAgentSettings | null>(null);
  const [dshSettingsError, setDshSettingsError] = useState('');
  const dshSupportsImages = dshSettings?.models.find(item => item.id === dshSettings.model)?.supportsImages ?? false;
  useEffect(() => {
    let active = true;
    getAiAgentSettings().then(settings => {
      if (active) { setDshSettings(settings); setDshSettingsError(''); }
    }).catch(error => { if (active) setDshSettingsError(String(error)); });
    return () => { active = false; };
  }, [showSettings]);
  const activeTurnRef = useRef<{
    sessionId: string; requestId: string; controller: AbortController;
    finished: Promise<void>; finish: () => void;
  } | null>(null);
  const [loadingStatus, setLoadingStatus] = useState('正在分析...');
  const chatContainerRef = useRef<HTMLDivElement>(null);
  const isAtBottom = useRef(true);
  const abortControllerRef = useRef<AbortController | null>(null);
  const activeModule = useAiContextStore(state => state.activeModule);
  const businessData = useAiContextStore(state => state.businessData);
  const lastUpdated = useAiContextStore(state => state.lastUpdated);
  const handledTemplateAssetRequestsRef = useRef<Set<string>>(new Set());

  useEffect(() => {
    ensureActiveSession(currentProject?.id);
  }, [currentProject?.id, ensureActiveSession]);

  useEffect(() => {
    const handleResize = () => {
      const compact = window.innerWidth < SESSION_SIDEBAR_BREAKPOINT;
      setIsCompactLayout(compact);
      setIsSidebarOpen(!compact);
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  useEffect(() => {
    const hydrateAiContext = () => {
      useAiContextStore.getState().hydrateFromStorage();
      useProjectStore.setState({ currentProject: readStoredCurrentProject() });
    };

    hydrateAiContext();

    if (!isTauriRuntime()) return;

    let disposed = false;
    const unlistenFns: Array<() => void> = [];

    Promise.all([
      listen(AI_CONTEXT_UPDATED_EVENT, hydrateAiContext),
      listen(AI_CONTEXT_REFRESH_REQUEST_EVENT, hydrateAiContext),
    ]).then((handlers) => {
      if (disposed) {
        handlers.forEach(handler => handler());
        return;
      }
      unlistenFns.push(...handlers);
    }).catch((error) => {
      console.warn('Failed to listen for AI context events:', error);
    });

    return () => {
      disposed = true;
      unlistenFns.forEach(handler => handler());
    };
  }, []);

  useEffect(() => {
    const appendTemplateAsset = (selection: AiTemplateAssetSelection | null) => {
      if (!selection || handledTemplateAssetRequestsRef.current.has(selection.requestId)) return;
      handledTemplateAssetRequestsRef.current.add(selection.requestId);
      setImages(prev => {
        const withoutDuplicate = prev.filter(image => image.assetId !== selection.assetId);
        const next: AiImageAttachment = {
          id: selection.requestId,
          name: selection.fileName || selection.fieldKey || selection.assetId,
          mimeType: selection.mimeType || 'image/png',
          size: selection.size || 0,
          source: 'template_asset',
          projectId: selection.projectId,
          templateId: selection.templateId,
          assetId: selection.assetId,
          fieldKey: selection.fieldKey || undefined,
        };
        return [...withoutDuplicate, next].slice(-4);
      });
      if (!input.trim()) {
        setInput('请分析这张模板图片的内容和明显问题。');
      }
    };

    const handleWindowEvent = (event: Event) => {
      appendTemplateAsset(parseTemplateAssetSelection((event as CustomEvent).detail));
    };
    const handleStorage = (event: StorageEvent) => {
      if (event.key === AI_TEMPLATE_ASSET_SELECTED_STORAGE_KEY) {
        appendTemplateAsset(parseTemplateAssetSelection(event.newValue));
      }
    };

    window.addEventListener(AI_TEMPLATE_ASSET_SELECTED_EVENT, handleWindowEvent);
    window.addEventListener('storage', handleStorage);
    appendTemplateAsset(parseTemplateAssetSelection(localStorage.getItem(AI_TEMPLATE_ASSET_SELECTED_STORAGE_KEY)));

    let disposed = false;
    let unlisten: (() => void) | undefined;
    if (isTauriRuntime()) {
      listen<AiTemplateAssetSelection>(AI_TEMPLATE_ASSET_SELECTED_EVENT, event => {
        appendTemplateAsset(parseTemplateAssetSelection(event.payload));
      }).then(handler => {
        if (disposed) {
          handler();
          return;
        }
        unlisten = handler;
      }).catch(error => {
        console.warn('Failed to listen for template asset selections:', error);
      });
    }

    return () => {
      disposed = true;
      unlisten?.();
      window.removeEventListener(AI_TEMPLATE_ASSET_SELECTED_EVENT, handleWindowEvent);
      window.removeEventListener('storage', handleStorage);
    };
  }, [input]);

  const resolveImagesForSend = async (pendingImages: AiImageAttachment[]) => {
    return Promise.all(pendingImages.map(async (image) => {
      if (image.source !== 'template_asset' || image.dataUrl) return image;
      if (!image.projectId || !image.assetId) {
        throw new Error('模板图片附件缺少 projectId 或 assetId');
      }
      const loaded = await loadAiTemplateAsset(image.projectId, image.assetId);
      return {
        ...image,
        id: image.id || loaded.id,
        name: loaded.name || image.name,
        mimeType: loaded.mimeType,
        size: loaded.size,
        dataUrl: loaded.dataUrl,
      };
    }));
  };

  const handleScroll = () => {
    if (!chatContainerRef.current) return;
    const { scrollTop, scrollHeight, clientHeight } = chatContainerRef.current;
    isAtBottom.current = scrollHeight - scrollTop - clientHeight < 100;
  };

  useEffect(() => {
    if (!isAtBottom.current) return;
    const frameId = requestAnimationFrame(() => {
      const container = chatContainerRef.current;
      if (container) container.scrollTop = container.scrollHeight;
    });
    return () => cancelAnimationFrame(frameId);
  }, [messages]);

  useEffect(() => {
    return () => {
      if (abortControllerRef.current) abortControllerRef.current.abort();
      flushSessionPersistence();
    };
  }, [flushSessionPersistence]);

  useEffect(() => {
    if (isAtBottom.current) {
      const container = chatContainerRef.current;
      if (container) container.scrollTop = container.scrollHeight;
    }
  }, []);

  useEffect(() => {
    if (isTyping || !pendingReceipts.length) return;
    for (const receipt of pendingReceipts) appendMessages(receipt.sessionId, [{ role: 'assistant', content: receipt.content, appReceipt: true }]);
    setPendingReceipts([]);
  }, [isTyping, pendingReceipts, appendMessages]);

  const handleSend = async (overrideInput?: string) => {
    const textToSend = overrideInput ?? input;
    const imagesToSend = overrideInput ? [] : [...images];
    if ((!textToSend.trim() && imagesToSend.length === 0) || isTyping || activeTurnRef.current) return;

    if (imagesToSend.length && !dshSupportsImages) {
      setDshSettingsError('当前模型不支持图片，请在下方 AI 模型与服务中选择视觉模型。');
      setShowSettings(true);
      return;
    }
    if (!bindingReady) return;
    const sessionId = ensureActiveSession(currentProject?.id);
    const userMessage = textToSend.trim();
    const promptText = userMessage || '请分析图片内容。';

    if (!overrideInput) {
      setInput('');
      setImages([]);
    }

    if (abortControllerRef.current) abortControllerRef.current.abort();
    const controller = new AbortController();
    abortControllerRef.current = controller;
    let finish!: () => void;
    const finished = new Promise<void>(resolve => { finish = resolve; });
    const turn = { sessionId, requestId: crypto.randomUUID(), controller, finished, finish };
    activeTurnRef.current = turn;

    const pendingMessages: AiChatMessage[] = [
      { role: 'user', content: userMessage, images: imagesToSend },
      { role: 'assistant', content: '' },
    ];
    appendMessages(sessionId, pendingMessages);
    setStreamingSessionId(sessionId);
    setIsTyping(true);

    try {
      // --- Enterprise LLM Infrastructure: AST Construction ---
      const systemRules: PromptRule[] = [
        { id: 'presales_role', content: 'You are a helpful presales AI consultant for Lamber system.', priority: 100 },
        { id: 'knowledge_base', content: SYSTEM_PROMPT_KNOWLEDGE, priority: 90 },
        { id: 'code_priority', content: '优先根据 [产品编号] (如 A302600342) 在知识库中匹配产品。只有当编号缺失时，才根据名称进行模糊匹配。', priority: 85 },
        { id: 'currency_unit_policy', content: 'Currency unit policy: all financial amount fields from BUSINESS CONTEXT are CNY yuan (元) unless the context explicitly says otherwise. Never label those raw values as ten-thousand yuan / 万元. If the user explicitly asks for 万元, divide the yuan value by 10,000 and state that conversion.', priority: 84 },
        { id: 'data_awareness', content: 'ALWAYS check the BUSINESS CONTEXT before answering. If data is missing, state it clearly.', priority: 80 },
        {
          id: 'saved_vs_draft_boundary',
          content: [
            'Project context may contain two clearly separated sources.',
            'Saved official state comes from the current Workspace SQLite database and represents persisted project data.',
            'Unsaved draft overlay comes from the current editing page and only represents temporary changes that have not been saved.',
            'When answering about current saved project status, prioritize saved official state.',
            'If using draft overlay content, explicitly call it "current unsaved changes" and do not claim it has been saved, submitted, recalculated, or written to the project.',
            'If saved state and draft overlay differ, point out the difference instead of silently merging them.',
            'Do not trigger or imply project writes, template saves, file operations, recalculations, NPV/IRR/margin/tax-rule changes, or reverse-calculation changes.',
          ].join('\n'),
          priority: 88,
        },
        {
          id: 'template_context_boundary',
          content: [
            'Template context rules:',
            'Specified template saved content comes from the current Workspace SQLite database and represents official persisted template data.',
            'Current template-page edits are unsaved draft overlay only; distinguish them from saved template content.',
            'Template images are metadata-only by default. Only images explicitly selected by the user for this turn are provided as vision input.',
            'Do not claim you modified, completed, saved, or generated template content.',
            'Do not change project/template data based on image analysis unless the user explicitly approves a supported text write or uses the image upload card in chat. Images are never written by the model; selecting or pasting into a labeled upload card is a user action. Merely attaching an image to a chat message does not save it to the template.',
          ].join('\n'),
          priority: 87,
        },
        {
          id: 'workspace_specified_project_boundary',
          content: [
            'Workspace specified project context rules:',
            'When context is marked as "Specified project saved official state", it was resolved from an explicit project name in the current user message and loaded from the current Workspace SQLite database by real projectId.',
            'If the user explicitly names a project, answer from that specified project context instead of defaulting to the currently opened project.',
            'If multiple specified project contexts are provided, keep each project source separate and do not merge fields across projects.',
            'If project matching is ambiguous or unavailable, do not guess project data; ask the user to specify the exact project.',
            'Current unsaved draft overlay belongs only to its marked projectId and must not override or contaminate another specified project.',
            'Project names are only routing hints for this turn; persisted reads must be treated as projectId-based Workspace SQLite reads.',
          ].join('\n'),
          priority: 89,
        },
      ];

      const binding = await invoke<{ projectId: string | null; workspaceId: string; projectName?: string | null } | null>('ai_get_session_binding', { sessionId });
      if (!binding) throw new Error('会话未绑定，请新建会话并选择项目');
      if (binding && imagesToSend.some(image => image.source === 'template_asset' && image.projectId !== binding.projectId)) {
        throw new Error('模板图片不属于会话绑定项目，请移除附件或新建对应项目会话');
      }
      if (binding) systemRules.push({ id: 'session_project_scope', priority: 100, content: sessionScopePrompt(binding.projectId) });
      systemRules.push({ id: 'demand_image_invitation', priority: 100,
        content: demandImageCompletionPrompt(Boolean(binding.projectId) && isDemandFormRequest(userMessage)) });
      systemRules.push({ id: 'saved_template_images', priority: 100,
        content: templateImagePrompt(Boolean(binding.projectId) && (wantsTemplateImages(userMessage) || imageCardSession === sessionId)) });
      systemRules.push({ id: 'document_generation_invitation', priority: 100,
        content: documentGenerationPrompt(binding.projectId ? documentTemplateRequests([{ role: 'user', content: userMessage }]) : []) });
      systemRules.push({id:'template_list_invitation',priority:100, content:templateListPrompt(binding.projectId ? templateListIntent([{role:'user',content:userMessage}]) : {tech:false,inquiry:false})});
      systemRules.push({ id: 'structure_reverse_invitation', priority: 100,
        content: structureReversePrompt(Boolean(binding.projectId) && structureReverseIntent([{ role: 'user', content: userMessage }]).requested) });
      const contextView = currentView || 'hub';
      const composedContext = await buildAiChatContext({
        currentView: contextView,
        userMessage: promptText,
        boundProjectId: binding.projectId,
      });

      const resolvedImagesToSend = await resolveImagesForSend(imagesToSend);
      const imageSourceNotes = resolvedImagesToSend
        .filter(image => image.source === 'template_asset')
        .map(image => `${image.name} (projectId=${image.projectId}, templateId=${image.templateId}, assetId=${image.assetId}, field=${image.fieldKey || '--'})`);

      const ast: PromptAST = {
        systemRules,
        dynamicState: {
          layer1Core: composedContext.contextNodes.savedOfficial,
          layer2Active: [
            ...composedContext.contextNodes.pageContext,
            ...composedContext.contextNodes.draftOverlay,
            ...(imageSourceNotes.length > 0 ? [{
              type: 'summary' as const,
              title: 'Explicit template image attachments for this turn',
              content: imageSourceNotes.map(note => `- ${note}`).join('\n'),
              metadata: { module: 'template_asset_vision_input' },
            }] : []),
          ],
          layer3Context: [...composedContext.contextNodes.warnings,
            ...appReceiptContext(useAiSessionStore.getState().sessions.find(item => item.id === sessionId)?.messages ?? [])],
        },
        userIntent: {
          raw: promptText,
          images: resolvedImagesToSend.length > 0 ? resolvedImagesToSend : undefined,
        },
      };

      setLoadingStatus('正在等待模型回复…');
      if (controller.signal.aborted) return;
      await dshRuntime.current.execute({
        sessionId, requestId: turn.requestId, text: new PromptRenderer().render(ast),
        signal: controller.signal, images: resolvedImagesToSend,
        onUpdate: message => updateLastAssistantMessage(sessionId, message),
        harnessSessionId: useAiSessionStore.getState().sessions.find(item => item.id === sessionId)?.harnessSessionId,
        onSession: id => useAiSessionStore.getState().setHarnessSessionId(sessionId, id),
      });
    } catch (error) {
      if ((error as Error).name === 'AbortError') {
        console.log('Stream aborted by user');
        return;
      }

      console.error('Chat error:', error);
      const partial = useAiSessionStore.getState().sessions.find(item => item.id === sessionId)?.messages.at(-1);
      updateLastAssistantMessage(sessionId, {
        content: `${partial?.content || ''}\n\n**Error:** 连接 AI 服务失败 (${String(error instanceof Error ? error.message : error)})`,
        think: partial?.think,
        toolCalls: partial?.toolCalls,
      });
    } finally {
      activeTurnRef.current = null;
      turn.finish();
      setIsTyping(false);
      flushSessionPersistence();
      setStreamingSessionId(null);
    }
  };

  const handleStop = async () => {
    const turn = activeTurnRef.current;
    if (!turn) return;
    turn.controller.abort();
    setLoadingStatus('正在停止生成…');
    await turn.finished;
    flushSessionPersistence();
  };

  const clearMessages = async () => {
    if (!currentSessionId) return;
    if (window.confirm('确定要清除当前会话的聊天记录吗？')) {
      if (activeTurnRef.current?.sessionId === currentSessionId) await handleStop();
      try {
        if (isTauriRuntime()) await invoke('ai_reset_session', { sessionId: currentSessionId });
      } catch (error) { window.alert(String(error)); return; }
      setBindingState(null);
      setBindingVersion(version => version + 1);
      resetSessionMessages(currentSessionId, {
        role: 'assistant',
        content: '聊天记录已清除。请问还有什么可以帮您？',
      });
    }
  };

  const handleCreateSession = () => {
    createSession(currentProject?.id);
    setInput('');
    setImages([]);
    setCopiedIdx(null);
    if (isCompactLayout) setIsSidebarOpen(false);
  };

  const handleSelectSession = (sessionId: string) => {
    selectSession(sessionId);
    setInput('');
    setImages([]);
    setCopiedIdx(null);
    isAtBottom.current = true;
    if (isCompactLayout) setIsSidebarOpen(false);
  };

  const handleRenameSession = (sessionId: string, title: string) => {
    setSessionTitle(sessionId, title, 'manual');
  };

  const handleDeleteSession = async (sessionId: string) => {
    const session = useAiSessionStore.getState().sessions.find(item => item.id === sessionId);
    if (!session || !window.confirm(`确定删除会话「${session.title}」吗？此操作无法撤销。`)) return;
    const wasCurrentSession = currentSessionId === sessionId;

    if (activeTurnRef.current?.sessionId === sessionId) await handleStop();

    try {
      if (isTauriRuntime()) await invoke('ai_reset_session', { sessionId });
    } catch (error) { window.alert(String(error)); return; }
    deleteSession(sessionId);
    if (useAiSessionStore.getState().sessions.length === 0) {
      createSession(currentProject?.id);
    }
    if (wasCurrentSession) {
      setInput('');
      setImages([]);
      setCopiedIdx(null);
      isAtBottom.current = true;
    }
  };

  const copyToClipboard = (text: string, idx: number) => {
    navigator.clipboard.writeText(text);
    setCopiedIdx(idx);
    setTimeout(() => setCopiedIdx(null), 2000);
  };

  const contextView = currentView && currentView !== 'hub' ? currentView : 'hub';
  const quickActionView = getAiContextView(contextView);
  const quickActionItems = [
    { label: '分析当前项目效益', icon: 'ai', view: 'ict' },
    { label: '推荐合适产品', icon: 'aiThinking' },
  ] satisfies Array<{ label: string; icon: AppIconName; view?: string }>;
  const quickActions = quickActionItems.filter(action => !action.view || action.view === quickActionView);
  const quickActionContextKey = quickActionView === 'hub' ? '' : getCoreContextKey(quickActionView);
  const activeContextData = activeModule && activeModule !== AI_CONTEXT_KEY.HUB
    ? businessData[activeModule]
    : undefined;
  const quickActionContextData = quickActionContextKey
    ? businessData[quickActionContextKey] ?? businessData[quickActionView]
    : undefined;
  const connectedContextModule = activeContextData
    ? activeModule
    : quickActionContextData
      ? (businessData[quickActionContextKey] ? quickActionContextKey : quickActionView)
      : '';
  const connectedScope = connectedContextModule
    ? getAiContextScope(connectedContextModule) ?? quickActionView
    : '';
  const statusLastUpdated = connectedContextModule
    ? lastUpdated[connectedContextModule] ?? lastUpdated[quickActionView]
    : lastUpdated[activeModule];
  const connectionStatusText = connectedContextModule
    ? `已连接：${connectedScope}`
    : '未检测到业务状态';
  const connectionStatusClassName = connectedContextModule
    ? 'bg-success-soft text-success'
    : 'bg-muted text-muted-foreground';
  const connectionDotClassName = connectedContextModule
    ? 'bg-success'
    : 'bg-muted-foreground/50';
  const isCurrentSessionStreaming = isTyping && currentSessionId === streamingSessionId;
  const isOtherSessionStreaming = isTyping && Boolean(streamingSessionId) && !isCurrentSessionStreaming;
  const sessionContextLabel = !bindingReady ? '未建立项目权限' : bindingState?.binding?.projectId
      ? `已绑定：${bindingState.binding.projectName || bindingState.binding.projectId}`
      : GENERAL_SESSION_LABEL;

  return (
    <div className="relative flex min-h-0 flex-1 overflow-hidden bg-background">
      {isCompactLayout && isSidebarOpen && (
        <button
          type="button"
          className="absolute inset-0 z-20 bg-foreground/10 backdrop-blur-[1px]"
          aria-label="收起会话列表"
          onClick={() => setIsSidebarOpen(false)}
        />
      )}

      {(!isCompactLayout || isSidebarOpen) && (
        <div className={isCompactLayout ? 'absolute inset-y-0 left-0 z-30' : 'relative'}>
          <AiSessionSidebar
            sessions={sessions}
            currentSessionId={currentSessionId}
            generatingSessionId={isTyping ? streamingSessionId : null}
            currentProjectId={currentProject?.id}
            currentProjectName={currentProject?.name}
            compact={isCompactLayout}
            onCreate={handleCreateSession}
            onSelect={handleSelectSession}
            onRename={handleRenameSession}
            onDelete={handleDeleteSession}
            onClose={() => setIsSidebarOpen(false)}
          />
        </div>
      )}

      <div className="flex min-h-0 min-w-0 flex-1 flex-col bg-background">
        <div className="flex h-12 shrink-0 items-center justify-between gap-3 bg-card/70 px-4">
          <div className="flex min-w-0 items-center gap-2.5">
            {isCompactLayout && (
              <button
                type="button"
                onClick={() => setIsSidebarOpen(true)}
                className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
                title="展开会话列表"
              >
                <AppIcon name="panelLeft" size={17} />
              </button>
            )}
            <div className="min-w-0">
              <div className="truncate text-[13px] font-semibold text-foreground">
                {currentSession?.title || '新会话'}
              </div>
              <div className="truncate text-[10px] text-muted-foreground">
                {sessionContextLabel}
              </div>
            </div>
          </div>

          {isOtherSessionStreaming && (
            <div className="flex shrink-0 items-center gap-1.5 rounded-full bg-primary-soft px-2.5 py-1 text-[10px] font-semibold text-primary">
              <span className="h-1.5 w-1.5 animate-pulse rounded-full bg-primary" />
              其他会话正在生成
            </div>
          )}
        </div>

      <div
        ref={chatContainerRef}
        onScroll={handleScroll}
        className="flex min-h-0 flex-1 flex-col gap-6 overflow-y-auto p-5 [&>*]:shrink-0"
      >
        {messages.map((msg, idx) => (
          <MessageBubble
            key={`${currentSessionId || 'session'}-${idx}`}
            msg={msg}
            idx={idx}
            isStreaming={isCurrentSessionStreaming && idx === messages.length - 1 && msg.role === 'assistant'}
            onCopy={copyToClipboard}
            copiedIdx={copiedIdx}
          />
        ))}
        {messages.length > 0 && isCurrentSessionStreaming && !messages[messages.length - 1]?.content && !messages[messages.length - 1]?.think && (
          <div className="flex animate-in items-center gap-3 self-start rounded-2xl rounded-bl-sm border border-border bg-muted p-4 text-foreground shadow-sm fade-in duration-300">
            <AppIcon name="loading" size={16} className="animate-spin text-primary" />
            <span className="animate-pulse text-xs font-bold text-secondary-foreground">{loadingStatus}</span>
          </div>
        )}
        {demandImagesRequested && bindingReady && bindingState?.binding?.projectId && currentSessionId && (
          <DemandImageCompletionCards key={`${currentSessionId}-${bindingState.binding.workspaceId}`}
            sessionId={currentSessionId} refreshToken={messages.length} disabled={isTyping}
            onReceipt={content => setPendingReceipts(previous => [...previous, { sessionId: currentSessionId, content }])} />
        )}
        {savedImagesRequested && bindingReady && bindingState?.binding?.projectId && currentSessionId && (
          <TemplateImageCard key={`images-${currentSessionId}-${bindingState.binding.workspaceId}`}
            sessionId={currentSessionId} disabled={isTyping}
            onReceipt={content => setPendingReceipts(previous => [...previous, { sessionId: currentSessionId, content }])}
            onAnalyze={image => { setImages(previous => [...previous.filter(item => item.assetId !== image.assetId), image].slice(-4));
              setInput('请分析这张模板图片的内容和明显问题。'); }} />
        )}
        {documentTemplateIds.length > 0 && bindingReady && bindingState?.binding?.projectId && currentSessionId && (
          <DocumentGenerationCards key={`${currentSessionId}-${bindingState.binding.workspaceId}-${documentTemplateIds.join(',')}`}
            sessionId={currentSessionId} templateIds={documentTemplateIds} refreshToken={messages.length} disabled={isTyping}
            onReceipt={content => setPendingReceipts(previous => [...previous, { sessionId: currentSessionId, content }])} />
        )}
        {listIntent.tech && bindingReady && bindingState?.binding?.projectId && currentSessionId && <TechItemsCard
          key={`tech-${currentSessionId}-${bindingState.binding.workspaceId}`} sessionId={currentSessionId} proposal={isTyping?[]:listProposal} disabled={isTyping}
          onReceipt={content=>setPendingReceipts(previous=>[...previous,{sessionId:currentSessionId,content}])}/>}
        {listIntent.inquiry && bindingReady && bindingState?.binding?.projectId && currentSessionId && <InquiryCard
          key={`inquiry-${currentSessionId}-${bindingState.binding.workspaceId}`} sessionId={currentSessionId} disabled={isTyping}
          onReceipt={content=>setPendingReceipts(previous=>[...previous,{sessionId:currentSessionId,content}])}/>}
        {reverseIntent.requested && bindingReady && bindingState?.binding?.projectId && currentSessionId && <StructureReverseCard
          key={`structure-${currentSessionId}-${bindingState.binding.workspaceId}-${reverseIntent.key}`}
          sessionId={currentSessionId} initialMetric={reverseIntent.metricType} initialTarget={reverseIntent.targetPercent}
          initialScenario={reverseIntent.scenario} disabled={isTyping}
          onReceipt={content => setPendingReceipts(previous => [...previous, { sessionId: currentSessionId, content }])} />}
      </div>

      {showSettings && <div className="max-h-80 overflow-auto p-4"><AiAgentSettingsCard onSaved={settings => { setDshSettings(settings); setDshSettingsError(''); }} /></div>}
      <div className="bg-muted/40 px-4 py-2 text-caption text-secondary-foreground">
        {dshSettingsError || (dshSupportsImages ? '当前视觉模型支持图片输入。' : '当前模型不支持图片，请在模型设置中选择视觉模型。')}
      </div>
      {!bindingReady && <div className="px-4 py-3">
        {bindingState?.sessionId !== currentSessionId ? <p className="text-caption text-muted-foreground">正在核对会话权限…</p>
          : bindingState.error ? <p role="alert" className="text-caption text-destructive">{bindingState.error}。请打开原工作区或新建会话。</p>
          : <AiSessionProjectPicker key={currentSessionId} legacy={Boolean(currentSession?.messages.some(message => message.role === 'user'))} onChoose={chooseBinding} />}
      </div>}
      <div className="bg-card p-4 shadow-[0_-8px_24px_hsl(var(--foreground)/0.025)]" hidden={!bindingReady}>
        {!isTyping && (
          <div className="mb-3 flex flex-wrap gap-2">
            {quickActions.map((action, index) => (
              <button
                key={index}
                type="button"
                onClick={() => handleSend(action.label)}
                className="flex items-center gap-1.5 rounded-full border border-border bg-muted px-3 py-1.5 text-[11px] font-semibold text-secondary-foreground transition-all hover:bg-primary/10 hover:text-primary"
              >
                <AppIcon name={action.icon} size={14} />
                {action.label}
              </button>
            ))}
          </div>
        )}

        <div className="mb-2 flex items-center justify-between">
          {bindingReady && bindingState?.binding?.projectId && currentSessionId && <button type="button"
            className="rounded-md bg-muted px-3 py-2 text-caption" onClick={() => setImageCardSession(currentSessionId)}>项目图片</button>}
          <button
            type="button"
            className="flex items-center gap-1 text-[11px] font-bold text-muted-foreground transition-colors hover:text-primary"
            onClick={() => setShowSettings(!showSettings)}
          >
            <AppIcon name="settings" size={14} /> 模型设置
          </button>
          <div className="flex items-center gap-2">
            <div
              className={`flex items-center gap-1 rounded-full px-2 py-0.5 text-[10px] font-bold ${connectionStatusClassName}`}
              title={`activeModule: ${activeModule || '--'} · lastUpdated: ${formatLastUpdated(statusLastUpdated)}`}
            >
              <span className={`h-1.5 w-1.5 rounded-full ${connectedContextModule ? 'animate-pulse' : ''} ${connectionDotClassName}`} />
              <span>{connectionStatusText}</span>
            </div>
            <button
              type="button"
              onClick={clearMessages}
              title="清除当前会话记录"
              className="rounded-md p-1.5 text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
            >
              <AppIcon name="delete" size={16} />
            </button>
          </div>
        </div>

        <AiInputBox
          input={input}
          images={images}
          isTyping={isTyping}
          visionEnabled={dshSupportsImages}
          onInputChange={setInput}
          onImagesChange={setImages}
          onSend={() => handleSend()}
          onStop={handleStop}
        />
      </div>
      </div>
    </div>
  );
}
