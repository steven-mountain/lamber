import { SESSION_SCOPE_HELP } from '../../ai/sessionScopePolicy';
import { useEffect, useState } from 'react';
import { listAiWorkspaceProjects, type AiWorkspaceProjectIndexItem } from '../../services/aiProjectContextService';

export default function AiSessionProjectPicker({ legacy, onChoose }: {
  legacy: boolean; onChoose: (projectId: string | null) => Promise<void>;
}) {
  const [projects, setProjects] = useState<AiWorkspaceProjectIndexItem[]>([]);
  const [choice, setChoice] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  useEffect(() => {
    let active = true;
    listAiWorkspaceProjects().then(items => { if (active) setProjects(items); })
      .catch(error => { if (active) setError(String(error)); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, []);
  const choose = async (projectId: string | null) => {
    setSaving(true); setError('');
    try { await onChoose(projectId); } catch (error) { setError(String(error)); }
    finally { setSaving(false); }
  };
  return <div className="space-y-3 rounded-lg bg-muted/50 p-4 text-caption">
    <div className="font-semibold text-foreground">{legacy ? '保留历史，新建会话继续' : '选择会话项目'}</div>
    <p className="text-muted-foreground">{SESSION_SCOPE_HELP}</p>
    <label className="block space-y-1">
      <span>已有项目</span>
      <select aria-label="会话绑定项目" value={choice} disabled={loading || saving}
        onChange={event => setChoice(event.target.value)} className="w-full rounded-md bg-card p-2 text-foreground">
        <option value="">{loading ? '正在读取项目…' : projects.length ? '请选择项目' : '当前工作区暂无项目'}</option>
        {projects.map(project => <option key={project.projectId} value={project.projectId}>{project.projectName}</option>)}
      </select>
    </label>
    <div className="flex flex-wrap gap-2">
      <button disabled={!choice || loading || saving} onClick={() => void choose(choice)}
        className="rounded-md bg-primary-soft px-3 py-2 font-semibold text-primary disabled:opacity-50">{saving ? '正在创建…' : '绑定项目并开始'}</button>
      <button disabled={loading || saving} onClick={() => void choose(null)}
        className="rounded-md bg-card px-3 py-2 text-secondary-foreground disabled:opacity-50">通用聊天</button>
    </div>
    {error && <p role="alert" className="text-destructive">{error}</p>}
  </div>;
}
