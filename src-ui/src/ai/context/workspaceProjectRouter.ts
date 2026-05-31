import type {
  AiProjectContextBundle,
  AiTemplateContextSummary,
  AiWorkspaceProjectIndexItem,
} from "../../services/aiProjectContextService";

export const MAX_SPECIFIED_PROJECT_CONTEXTS = 2;

export interface ResolvedWorkspaceProject {
  projectId: string;
  projectName: string;
  matchType: "exact" | "normalized";
}

export interface WorkspaceProjectResolution {
  projects: ResolvedWorkspaceProject[];
  warnings: string[];
  hasExplicitProjectReference: boolean;
}

export interface TemplateResolution {
  templateId?: string;
  templateName?: string | null;
  warnings: string[];
}

const TEMPLATE_ALIAS_TERMS = [
  "立项签批表",
  "项目概况",
  "需求导入表",
  "会审纪要",
  "效益分析表",
  "立项决策",
];

function normalizeText(value: string) {
  return value
    .trim()
    .replace(/\s+/g, "")
    .toLocaleLowerCase();
}

function uniqueByProjectId(projects: ResolvedWorkspaceProject[]) {
  const seen = new Set<string>();
  return projects.filter(project => {
    if (seen.has(project.projectId)) return false;
    seen.add(project.projectId);
    return true;
  });
}

function hasDuplicateNormalizedName(projects: AiWorkspaceProjectIndexItem[], normalizedName: string) {
  return projects.filter(project => normalizeText(project.projectName) === normalizedName).length > 1;
}

export function resolveWorkspaceProjectsFromMessage(
  userMessage: string,
  projectIndex: AiWorkspaceProjectIndexItem[],
): WorkspaceProjectResolution {
  const message = normalizeText(userMessage);
  const warnings: string[] = [];

  if (!message || projectIndex.length === 0) {
    return {
      projects: [],
      warnings,
      hasExplicitProjectReference: false,
    };
  }

  const mentioned = projectIndex
    .map(project => ({
      project,
      normalizedName: normalizeText(project.projectName),
      exactNameMentioned: userMessage.includes(project.projectName),
    }))
    .filter(item => item.normalizedName.length >= 2 && message.includes(item.normalizedName));

  if (mentioned.length === 0) {
    return {
      projects: [],
      warnings,
      hasExplicitProjectReference: false,
    };
  }

  const ambiguousNames = mentioned
    .filter(item => hasDuplicateNormalizedName(projectIndex, item.normalizedName))
    .map(item => item.project.projectName);

  if (ambiguousNames.length > 0) {
    warnings.push(`Project name is not unique in the current Workspace: ${Array.from(new Set(ambiguousNames)).join(", ")}. Ask the user to choose the exact project.`);
    return {
      projects: [],
      warnings,
      hasExplicitProjectReference: true,
    };
  }

  const resolved = uniqueByProjectId(mentioned.map(item => ({
    projectId: item.project.projectId,
    projectName: item.project.projectName,
    matchType: item.exactNameMentioned ? "exact" : "normalized",
  })));

  if (resolved.length > MAX_SPECIFIED_PROJECT_CONTEXTS) {
    warnings.push(`The user mentioned ${resolved.length} projects. Deep official context loading is limited to ${MAX_SPECIFIED_PROJECT_CONTEXTS} explicitly named projects per turn; ask the user to narrow the comparison.`);
    return {
      projects: [],
      warnings,
      hasExplicitProjectReference: true,
    };
  }

  return {
    projects: resolved,
    warnings,
    hasExplicitProjectReference: true,
  };
}

function templateKeys(template: AiTemplateContextSummary) {
  return [template.templateId, template.templateName]
    .filter((value): value is string => Boolean(value && value.trim()))
    .map(value => normalizeText(value));
}

function uniqueTemplates(templates: AiTemplateContextSummary[]) {
  const seen = new Set<string>();
  return templates.filter(template => {
    const key = template.templateId || template.templateName || "";
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function resolveTemplateFromMessage(
  userMessage: string,
  projectContext: AiProjectContextBundle,
): TemplateResolution {
  const message = normalizeText(userMessage);
  const templates = projectContext.templates || [];
  if (!message || templates.length === 0) {
    return { warnings: [] };
  }

  let matches = templates.filter(template => {
    return templateKeys(template).some(key => key && message.includes(key));
  });

  if (matches.length === 0) {
    const matchedAlias = TEMPLATE_ALIAS_TERMS
      .map(normalizeText)
      .find(alias => alias && message.includes(alias));
    if (matchedAlias) {
      matches = templates.filter(template => {
        return templateKeys(template).some(key => key.includes(matchedAlias));
      });
    }
  }

  matches = uniqueTemplates(matches);

  if (matches.length === 0) {
    return { warnings: [] };
  }
  if (matches.length > 1) {
    return {
      warnings: [`Template name is ambiguous for project ${projectContext.projectName}: ${matches.map(item => item.templateName || item.templateId).join(", ")}. Do not load template detail until the user specifies one.`],
    };
  }

  const [template] = matches;
  return {
    templateId: template.templateId,
    templateName: template.templateName,
    warnings: [],
  };
}
