import { invoke } from "@tauri-apps/api/core";
import { getDemandCompletion } from "../../lib/templateCompletion/demand";
import {
  buildAiProjectContext,
  listAiWorkspaceProjects,
  type AiProjectContextSource,
  type AiWorkspaceProjectIndexItem,
} from "../../services/aiProjectContextService";
import { useAiContextStore, type AiContextSnapshot } from "../../store/useAiContextStore";
import { readStoredNavigationState } from "../../store/useNavigationStore";
import { readStoredCurrentProject } from "../../store/useProjectStore";
import { useSaveStore } from "../../store/useSaveStore";
import type { ContextNode } from "../types";
import { buildDraftOverlay } from "./buildDraftOverlay";
import type { AiComposedChatContext, AiSavedProjectContext } from "./types";
import {
  resolveTemplateFromMessage,
  resolveWorkspaceProjectsFromMessage,
} from "./workspaceProjectRouter";

interface BuildAiChatContextInput {
  currentView: string;
  userMessage: string;
  /** null is explicit general chat; missing identity always fails closed. */
  boundProjectId: string | null;
}

function isProjectAwareView(view: string) {
  return view === "project_board" || view === "ict_lifecycle" || view === "ict";
}

function readStoredIctProjectId() {
  if (typeof window === "undefined") return null;
  const projectId = window.localStorage.getItem("lamber_active_project_id")?.trim();
  return projectId || null;
}

function getCurrentProjectId(currentView: string) {
  if (!isProjectAwareView(currentView)) return null;

  const storedNavigation = readStoredNavigationState();
  const storedProject = readStoredCurrentProject();

  if (currentView === "ict_lifecycle" || currentView === "ict") {
    return (
      storedNavigation.activeProjectId ||
      readStoredIctProjectId() ||
      storedProject?.id ||
      null
    );
  }

  return (
    storedProject?.id ||
    null
  );
}

function buildRequestedSources(currentView: string, userMessage: string, activeTemplateId?: string | null): AiProjectContextSource[] | undefined {
  if (activeTemplateId && (currentView === "ict_lifecycle" || currentView === "ict")) {
    void userMessage;
    return ["templates", "template_detail"];
  }
  // Undefined uses the backend default summary mode and avoids requesting detailed JSON.
  return undefined;
}

function getActiveTemplateId(aiState: Pick<AiContextSnapshot, "activeModule" | "businessData">) {
  const activeModule = aiState.activeModule;
  if (!activeModule?.startsWith("ict.template.")) return null;
  const data = aiState.businessData[activeModule];
  if (data && typeof data === "object" && !Array.isArray(data)) {
    const selectedTemplate = (data as Record<string, unknown>).selectedTemplate;
    if (typeof selectedTemplate === "string" && selectedTemplate.trim()) {
      return selectedTemplate.trim();
    }
  }
  const suffix = activeModule.slice("ict.template.".length).trim();
  return suffix || null;
}

function buildSavedNode(saved: AiSavedProjectContext): ContextNode {
  const data = saved.data;
  const isSpecified = saved.resolution?.reason === "specified_project";
  return {
    type: "json",
    title: isSpecified
      ? `Specified project saved official state (Workspace SQLite): ${data.projectName}`
      : "Saved official project state (Workspace SQLite)",
    content: {
      source: saved.source,
      projectId: saved.projectId,
      projectName: data.projectName,
      resolution: saved.resolution,
      overview: data.overview,
      lifecycle: data.lifecycle,
      cashflow: data.cashflow,
      benefit: data.benefit,
      templates: data.templates,
      templateDetail: data.templateDetail ? {
        source: data.templateDetail.source,
        status: data.templateDetail.hasSavedState
          ? "Saved template detail loaded from the current Workspace SQLite database."
          : "No saved official template detail exists for this template.",
        projectId: data.templateDetail.projectId,
        templateId: data.templateDetail.templateId,
        templateName: data.templateDetail.templateName,
        updatedAt: data.templateDetail.updatedAt,
        fields: data.templateDetail.fields,
        fieldMapping: data.templateDetail.fieldMapping,
        outputConfig: data.templateDetail.outputConfig,
        assets: data.templateDetail.assets,
        completion: data.templateDetail.templateId.includes("需求导入表") ? (() => {
          const items = getDemandCompletion(data.templateDetail.templateId, data.templateDetail.fields, data.templateDetail.assets);
          return { items, missingFields: items.filter(item => item.kind === 'field' && !item.filled), missingImages: items.filter(item => item.kind === 'image' && !item.filled) };
        })() : undefined,
        warnings: data.templateDetail.warnings,
      } : undefined,
      files: data.files,
      sourceRows: data.sources,
      warnings: data.warnings,
    },
    metadata: { module: "workspace_sqlite" },
  };
}

function buildDraftNode(draft: NonNullable<AiComposedChatContext["draftOverlay"]>): ContextNode {
  return {
    type: "json",
    title: "Current unsaved draft overlay",
    content: {
      source: draft.source,
      status: "The following state is not saved to SQLite and must not be treated as official project data.",
      projectId: draft.projectId,
      view: draft.view,
      dirtyScopes: draft.dirtyScopes,
      data: draft.data,
    },
    metadata: { module: "unsaved_frontend_draft" },
  };
}

function buildWorkspaceProjectIndexNode(projectIndex: AiWorkspaceProjectIndexItem[]): ContextNode[] {
  if (projectIndex.length === 0) return [];
  return [{
    type: "json",
    title: "Current Workspace lightweight project index",
    content: {
      source: "workspace_sqlite_project_index",
      status: "Read-only lightweight index from the current Workspace SQLite database. It contains project identity and saved-state existence metadata only.",
      projectCount: projectIndex.length,
      projects: projectIndex.slice(0, 100).map(project => ({
        projectId: project.projectId,
        projectName: project.projectName,
        customerName: project.customerName,
        status: project.status,
        updatedAt: project.updatedAt,
        hasLifecycleState: project.hasLifecycleState,
        hasCashflowState: project.hasCashflowState,
        hasTemplateState: project.hasTemplateState,
        templateNames: project.templateNames,
        hasBenefitSchemes: project.hasBenefitSchemes,
      })),
      truncatedProjectCount: Math.max(projectIndex.length - 100, 0),
    },
    metadata: { module: "workspace_sqlite_project_index" },
  }];
}

function buildWarningNode(warnings: string[]): ContextNode[] {
  if (warnings.length === 0) return [];
  return [{
    type: "summary",
    title: "Context warnings",
    content: warnings.map(item => `- ${item}`).join("\n"),
    metadata: { module: "context_warnings" },
  }];
}

function isWorkspaceProjectListQuery(userMessage: string) {
  return /(哪些|所有|全部|工作区|Workspace|workspace).{0,12}项目|项目.{0,12}(哪些|所有|全部|列表|填写|完成)/.test(userMessage);
}

function hasPotentialNamedProjectReference(userMessage: string) {
  if (/当前项目|这个项目|该项目|本项目|此项目/.test(userMessage)) return false;
  return /[A-Za-z0-9_\-\u4e00-\u9fa5（）()]{2,80}(?:项目)?的/.test(userMessage);
}

export async function buildAiChatContext(input: BuildAiChatContextInput): Promise<AiComposedChatContext> {
  if (input.boundProjectId === undefined) throw new Error('缺少会话项目权限，不能读取业务上下文');
  const warnings: string[] = [];
  if (input.boundProjectId === null) return {
    savedProjectContexts: [], warnings: ['通用聊天不自动提供业务上下文；跨项目汇总须通过 query_projects 读取；另允许甄选费纯计算，项目明细和写工具禁用。'],
    contextNodes: { savedOfficial: [], pageContext: [], draftOverlay: [], warnings: [] },
  };
  const pageProjectId = getCurrentProjectId(input.currentView);
  const projectId = input.boundProjectId;
  const latestAiState = useAiContextStore.getState();
  const aiSnapshot: AiContextSnapshot = {
    activeModule: latestAiState.activeModule,
    businessData: latestAiState.businessData,
    lastUpdated: latestAiState.lastUpdated,
  };
  const dirtyScopes = useSaveStore.getState().dirtyScopes;
  const activeTemplateId = pageProjectId === projectId ? getActiveTemplateId(aiSnapshot) : null;
  const workspaceListQuery = isWorkspaceProjectListQuery(input.userMessage);
  const potentialNamedProjectReference = hasPotentialNamedProjectReference(input.userMessage);

  let savedProjectContext: AiSavedProjectContext | undefined;
  let savedProjectContexts: AiSavedProjectContext[] = [];
  let workspaceProjectIndex: AiWorkspaceProjectIndexItem[] = [];
  let explicitProjectResolution:
    | ReturnType<typeof resolveWorkspaceProjectsFromMessage>
    | undefined;

  {
    try {
      workspaceProjectIndex = (await listAiWorkspaceProjects()).filter(project => project.projectId === projectId);
      warnings.push("此上下文只提供绑定项目。跨项目汇总须通过 query_projects；其结果不得代替当前项目结论。其他项目明细需另建会话。");
      explicitProjectResolution = resolveWorkspaceProjectsFromMessage(input.userMessage, workspaceProjectIndex);
      warnings.push(...explicitProjectResolution.warnings);
      if (
        !explicitProjectResolution.hasExplicitProjectReference &&
        (potentialNamedProjectReference || (!projectId && input.userMessage.includes("项目")))
      ) {
        warnings.push("No current Workspace project name matched the user message. Ask the user to provide an exact project name when no active project is selected.");
      }
    } catch (error) {
      warnings.push(`Workspace project index unavailable: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  if (explicitProjectResolution?.projects.length) {
    for (const target of explicitProjectResolution.projects) {
      try {
        let data = await buildAiProjectContext({ projectId: target.projectId });
        const templateResolution = resolveTemplateFromMessage(input.userMessage, data);
        warnings.push(...templateResolution.warnings);
        if (templateResolution.templateId) {
          data = await buildAiProjectContext({
            projectId: target.projectId,
            requestedSources: ["templates", "template_detail"],
            activeTemplateId: templateResolution.templateId,
          });
        }
        savedProjectContexts.push({
          source: "workspace_sqlite",
          projectId: target.projectId,
          projectName: data.projectName,
          resolution: {
            reason: "specified_project",
            matchedName: target.projectName,
            templateId: templateResolution.templateId,
            templateName: templateResolution.templateName,
          },
          data,
        });
      } catch (error) {
        warnings.push(`Specified project context unavailable (${target.projectName}, ${target.projectId}): ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  } else if (projectId && !workspaceListQuery && !potentialNamedProjectReference && !explicitProjectResolution?.hasExplicitProjectReference) {
    try {
      const data = await buildAiProjectContext({
        projectId,
        requestedSources: buildRequestedSources(input.currentView, input.userMessage, activeTemplateId),
        activeTemplateId,
      });
      if (data.projectId !== projectId) {
        warnings.push(`SQLite project context mismatch: expected ${projectId}, got ${data.projectId}`);
      } else {
        savedProjectContext = {
          source: "workspace_sqlite",
          projectId,
          projectName: data.projectName,
          resolution: {
            reason: "active_project",
            templateId: activeTemplateId || undefined,
            templateName: activeTemplateId || undefined,
          },
          data,
        };
        savedProjectContexts = [savedProjectContext];
      }
    } catch (error) {
      warnings.push(`Saved project context unavailable: ${error instanceof Error ? error.message : String(error)}`);
    }
  } else if (workspaceListQuery) {
    warnings.push("Workspace-level project question detected; using the lightweight Workspace project index instead of defaulting to the active project.");
  } else if (potentialNamedProjectReference) {
    warnings.push("The user appears to reference a named project, but it was not found uniquely in the current Workspace. Do not answer using the active project as a substitute.");
  } else if (explicitProjectResolution?.hasExplicitProjectReference) {
    warnings.push("The user mentioned a project, but it was not uniquely resolved in the current Workspace. Do not answer using another project as a substitute.");
  }

  if (/需求分析|需求导入|缺.*附件|补.*附件/.test(input.userMessage)) {
    for (const saved of savedProjectContexts) {
      if (saved.data.templateDetail?.templateId.includes("需求导入表")) continue;
      try {
        const available = await invoke<string[]>('get_available_templates', { moduleId: 'ict_lifecycle' });
        const candidates = available.filter(name => name.includes('需求导入表'));
        if (candidates.length !== 1) { warnings.push("需求导入表模板无法唯一确定，请先在模板页选择目标模板。"); continue; }
        saved.data = await buildAiProjectContext({ projectId: saved.projectId, requestedSources: ['templates', 'template_detail'], activeTemplateId: candidates[0] });
      } catch (error) { warnings.push(`需求表缺项读取失败：${String(error)}`); }
    }
  }

  if (!savedProjectContext && savedProjectContexts.length > 0) {
    savedProjectContext = savedProjectContexts[0];
  }

  const draftProjectId = savedProjectContexts.length > 0
    ? (projectId && savedProjectContexts.some(context => context.projectId === projectId) ? projectId : null)
    : projectId;

  const draftOverlay = pageProjectId !== projectId ? undefined : buildDraftOverlay({
    projectId: draftProjectId,
    currentView: input.currentView,
    dirtyScopes,
    aiSnapshot,
  });

  const shouldAttachWorkspaceProjectIndex =
    savedProjectContexts.length === 0 &&
    workspaceProjectIndex.length > 0 &&
    (!projectId || workspaceListQuery || /项目|哪些|全部|所有|模板|签批表/.test(input.userMessage));

  return {
    savedProjectContext,
    savedProjectContexts,
    draftOverlay,
    warnings,
    contextNodes: {
      savedOfficial: savedProjectContexts.map(buildSavedNode),
      pageContext: [
        ...(shouldAttachWorkspaceProjectIndex ? buildWorkspaceProjectIndexNode(workspaceProjectIndex) : []),
      ],
      draftOverlay: draftOverlay ? [buildDraftNode(draftOverlay)] : [],
      warnings: buildWarningNode(warnings),
    },
  };
}
