import { useModulePath } from "../hooks/useModulePath";
import AppIcon from "./icons/AppIcon";

interface WorkspaceHeaderProps {
  moduleId: string;
  title: string;
  onBack: () => void;
  onPathChange?: (newPath: string) => void;
}

export default function WorkspaceHeader({ moduleId, title, onBack, onPathChange }: WorkspaceHeaderProps) {
  const { path, isLoading, updatePath } = useModulePath(moduleId);

  const handleUpdate = async () => {
    const newPath = await updatePath();
    if (newPath && onPathChange) {
      onPathChange(newPath);
    }
  };

  return (
    <div className="flex flex-col w-full shrink-0">
      <div className="h-16 bg-card border-b border-border flex items-center px-6 gap-4">
        <button 
          onClick={onBack} 
          className="text-secondary-foreground hover:text-primary hover:bg-secondary font-semibold flex items-center gap-1.5 px-3 py-2 rounded-lg transition-colors"
        >
          <span>←</span> 返回集市
        </button>
        <h2 className="m-0 text-lg font-bold text-foreground border-l-2 border-border pl-4">{title}</h2>
      </div>

      <div className="px-8 py-4 bg-secondary/30 border-b border-border flex items-center justify-between animate-in slide-in-from-top duration-300">
        <div className="flex flex-col gap-1">
          <div className="text-[10px] uppercase tracking-wider font-extrabold text-secondary-foreground opacity-60">
            当前模块工作空间
          </div>
          <div className={`text-sm font-mono flex items-center gap-2 ${path ? 'text-primary' : 'text-destructive font-bold'}`}>
            <AppIcon name={path ? "folder" : "warning"} size={16} />
            {isLoading ? '加载中...' : path || '未设置工作目录 (点击右侧按钮进行配置)'}
          </div>
        </div>
        
        <div className="flex gap-3">
          {path && (
             <button 
               onClick={() => window.open(`file://${path}`)}
               className="px-4 py-2 bg-card border border-border text-xs font-bold rounded-lg hover:bg-secondary transition-all shadow-sm"
             >
               打开目录
             </button>
          )}
          <button 
            onClick={handleUpdate}
            className="px-4 py-2 bg-primary text-primary-foreground text-xs font-bold rounded-lg hover:shadow-lg hover:-translate-y-0.5 transition-all active:translate-y-0"
          >
            {path ? '修改目录' : '立即设置'}
          </button>
        </div>
      </div>
    </div>
  );
}
