export interface TemplateCompletionItem {
  label: string;
  filled: boolean;
}

export interface TemplateCompletion {
  completedCount: number;
  totalCount: number;
  percent: number;
}

export function getTemplateCompletion(items: TemplateCompletionItem[]): TemplateCompletion {
  const totalCount = items.length;
  const completedCount = items.filter(item => item.filled).length;
  return {
    completedCount,
    totalCount,
    percent: totalCount > 0 ? Math.round((completedCount / totalCount) * 100) : 0,
  };
}
