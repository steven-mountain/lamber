import { useEffect, useRef, useState } from 'react';
import { loadListTargets, runListAction, type ListTarget } from '../../services/chatTemplateLists';
import type { TechItem } from '../../services/templateListTypes';
const inputClass = 'w-full min-w-0 rounded-md bg-card p-2 text-body';
export default function TechItemsCard({sessionId,proposal,disabled,onReceipt}: {
  sessionId:string;proposal:TechItem[];disabled:boolean;onReceipt:(content:string)=>void;
}) {
  const [targets,setTargets] = useState<ListTarget[]>([]);
  const [selected,setSelected] = useState(0);
  const [rows,setRows] = useState<TechItem[]>([]);
  const [error,setError] = useState('');
  const [busy,setBusy] = useState(false);
  const [revision,setRevision] = useState(0);
  const active = useRef(true); const running = useRef(false);
  useEffect(() => {active.current=true;return () => {active.current=false;};},[]);
  useEffect(() => {
    let valid=true;
    void loadListTargets(sessionId,['demand','meeting']).then(next => {if(valid){setTargets(next);setSelected(0);setRows(next[0]?.snapshot.techItems ?? []);setError('');}})
      .catch(reason=>{if(valid)setError(String(reason));});
    return ()=>{valid=false;};
  },[sessionId,revision]);
  const target = targets[selected];
  const save = async () => {
    if (!target || running.current) return;
    running.current=true;setBusy(true);setError('');
    try {
      const next = await runListAction(target,{type:'saveTech',rows,expected:target.snapshot.techItems,sharedTemplates:targets.map(item=>({templateName:item.templateName,expected:item.snapshot.techItems}))});
      if(active.current){setRows(next.snapshot.techItems);setTargets(previous=>previous.map(item=>({...item,snapshot:{...item.snapshot,techItems:next.snapshot.techItems}})));}
      onReceipt(`技术清单已由用户保存：${target.projectName}，共${next.snapshot.techItems.length}行，同时更新需求导入表与会审纪要。`);
    } catch(reason) {const message=String(reason);if(active.current)setError(message);onReceipt(`技术清单未保存：${message}`);}
    finally {running.current=false;if(active.current)setBusy(false);}
  };
  return <section className="rounded-lg bg-muted/50 p-4 space-y-3 text-body" aria-label="技术清单卡片">
    <h3 className="font-semibold">技术方案可行性清单</h3>
    <p className="font-semibold">此清单同时用于需求导入表与会审纪要</p>
    <p className="text-caption text-secondary-foreground">{target?.projectName} · 数量不参与财务计算。核对表格后点击保存。</p>
    {targets.length>1 && <label className="block text-caption">载入清单来源（切换将替换卡片草稿）<select disabled={busy} className={inputClass} value={selected} onChange={event=>{const index=Number(event.target.value);setSelected(index);setRows(targets[index].snapshot.techItems);}}>{targets.map((item,index)=><option key={item.templateName} value={index}>{item.templateName}（{item.snapshot.techItems.length}行）</option>)}</select></label>}
    {target && <>
      <p className="text-caption">保存将同时更新：{targets.map(item=>item.templateName).join('、')}。</p>
      {proposal.length>0 && <button type="button" disabled={disabled||busy} className="rounded-md bg-card px-3 py-2" onClick={()=>setRows(proposal.map(row=>({...row})))}>采用本轮建议（{proposal.length}行，替换卡片草稿）</button>}
      <div className="overflow-x-auto"><table className="w-full min-w-[520px] table-fixed text-left"><thead><tr>{['服务名称','服务描述','数量','单位','操作'].map((label,index)=><th key={label} className={`p-2 text-caption ${index===1?'w-2/5':index>1?'w-20':''}`}>{label}</th>)}</tr></thead><tbody>
        {rows.map((row,index)=><tr key={index}>
          {(['serviceName','serviceDesc','amount','unit'] as const).map(key=><td className="p-1 align-top" key={key}><input aria-label={`第${index+1}行${{serviceName:'服务名称',serviceDesc:'服务描述',amount:'数量',unit:'单位'}[key]}`} disabled={busy} type={key==='amount'?'number':'text'} min={key==='amount'?0:undefined} className={`${inputClass} ${key==='amount'?'tabular-nums':''}`} value={row[key]} onChange={event=>setRows(previous=>previous.map((item,i)=>i===index?{...item,[key]:event.target.value}:item))}/></td>)}
          <td><button type="button" disabled={busy} aria-label={`删除第${index+1}行`} className="rounded-md bg-card px-2 py-2 text-caption" onClick={()=>setRows(previous=>previous.filter((_,i)=>i!==index))}>删除</button></td>
        </tr>)}
      </tbody></table></div>
      <div className="flex flex-wrap gap-2"><button type="button" disabled={busy||rows.length>=100} className="rounded-md bg-card px-3 py-2" onClick={()=>setRows(previous=>[...previous,{serviceName:'',serviceDesc:'',amount:1,unit:'套'}])}>添加一行</button><button type="button" disabled={disabled||busy} className="rounded-md bg-primary px-4 py-2 text-primary-foreground disabled:opacity-50" onClick={()=>void save()}>{busy?'正在保存…':'保存到两张表'}</button></div>
    </>}
    {error && <p role="alert" className="text-destructive">{error}</p>}
    <button type="button" disabled={busy} className="rounded-md bg-card px-3 py-2 text-caption" onClick={()=>setRevision(value=>value+1)}>重新读取（放弃卡片草稿）</button>
  </section>;
}
