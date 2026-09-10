import { useEffect, useRef, useState } from 'react';
import { loadListTargets, runListAction, readQuoteImage, type ListTarget } from '../../services/chatTemplateLists';
import type { InquiryVendor, PendingQuoteImage } from '../../services/templateListTypes';
const inputClass='w-full rounded-md bg-card p-2 text-body';
function SavedQuoteImage({assetId,target}: {assetId:string;target:ListTarget}) {
  const [url,setUrl]=useState('');const [error,setError]=useState('');
  useEffect(()=>{let valid=true;void readQuoteImage(target,assetId).then(url=>{if(valid)setUrl(url);}).catch(()=>{if(valid)setError('截图文件缺失，请补传');});return()=>{valid=false;};},[assetId,target]);
  return url?<img src={url} alt="已保存的报价截图" className="max-h-36 max-w-full rounded-md object-contain"/>:<span className="text-caption">{error||'正在载入截图…'}</span>;
}
export default function InquiryCard({sessionId,disabled,onReceipt}: {sessionId:string;disabled:boolean;onReceipt:(content:string)=>void}) {
  const [targets,setTargets]=useState<ListTarget[]>([]);
  const [selected,setSelected]=useState(0);
  const [rows,setRows]=useState<InquiryVendor[]>([]);
  const [uploads,setUploads]=useState<PendingQuoteImage[]>([]);
  const [error,setError]=useState('');
  const [busy,setBusy]=useState(false);
  const [revision,setRevision]=useState(0);
  const active=useRef(true);const running=useRef(false);
  useEffect(()=>{active.current=true;return()=>{active.current=false;};},[]);
  useEffect(()=>{
    let valid=true;
    void loadListTargets(sessionId,['meeting']).then(next=>{if(valid){setTargets(next);setSelected(0);setRows(next[0]?.snapshot.inqVendors??[]);setUploads([]);setError('');}})
      .catch(reason=>{if(valid)setError(String(reason));});
    return()=>{valid=false;};
  },[sessionId,revision]);
  const target=targets[selected];
  const targetRef=useRef(target);targetRef.current=target;
  const act=async (generate:boolean)=>{
    if(!target||running.current)return;
    running.current=true;setBusy(true);setError('');
    try {
      const next=await runListAction(target,generate?{type:'generateInquiry',expected:target.snapshot.inqVendors}:{type:'saveInquiry',rows,expected:target.snapshot.inqVendors,uploads});
      if(active.current){setTargets(previous=>previous.map((item,i)=>i===selected?next:item));setRows(next.snapshot.inqVendors);setUploads([]);}
      onReceipt(`${target.projectName} · ${generate?'用户确认后已调用原有生成器，实际生成':'用户已保存'}询价${next.snapshot.inqVendors.length}行。具体报价及证据请在卡片核对；AI不参与厂商或金额编辑。`);
    } catch(reason){const message=String(reason);if(active.current)setError(message);onReceipt(`询价操作未完成：${message}\n请先在项目页完善IT成本或收入；不提供绕过校验的方案。`);}
    finally{running.current=false;if(active.current)setBusy(false);}
  };
  const chooseImage=async (file:File|undefined,row:number)=>{
    if(!file)return;
    const originalTarget=targetRef.current;
    try {
      if(!['image/png','image/jpeg','image/webp'].includes(file.type)||file.size>20*1024*1024)throw new Error('截图仅支持PNG/JPEG/WebP，单张不超过20MB');
      const base64Data=await new Promise<string>((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result));reader.onerror=()=>reject(new Error('读取截图失败'));reader.readAsDataURL(file);});
      const dimensions=await new Promise<{width:number;height:number}>((resolve,reject)=>{const image=new Image();image.onload=()=>resolve({width:image.naturalWidth,height:image.naturalHeight});image.onerror=()=>reject(new Error('无法识别截图'));image.src=base64Data;});
      if(active.current && targetRef.current===originalTarget)setUploads(previous=>[...previous,{row,base64Data,name:file.name,...dimensions}]);
    }catch(reason){if(active.current)setError(String(reason));}
  };
  return <section className="rounded-lg bg-muted/50 p-4 space-y-3 text-body" aria-label="询价卡片">
    <h3 className="font-semibold">询价情况与报价证据</h3>
    <p className="text-caption text-secondary-foreground">{target?.projectName} · {target?.templateName}</p>
    <p>确认后按 IT 投入成本生成三家报价，最高价不超过含税总收入。生成后展示实际结果，由你修改厂商名称、报价和上传截图。</p>
    {targets.length>1&&<select aria-label="询价目标模板" disabled={busy} className={inputClass} value={selected} onChange={event=>{const index=Number(event.target.value);setSelected(index);setRows(targets[index].snapshot.inqVendors);setUploads([]);}}>{targets.map((item,index)=><option key={item.templateName} value={index}>{item.templateName}</option>)}</select>}
    {target&&<>
      {rows.length>0&&<p className="text-caption">重新生成会替换现有厂商和报价，按原有规则保留截图；请核对证据仍与报价一致。</p>}
      <button type="button" disabled={disabled||busy} className="rounded-md bg-primary px-4 py-2 text-primary-foreground disabled:opacity-50" onClick={()=>void act(true)}>{busy?'正在处理…':rows.length?'确认重新生成三家报价':'确认生成三家报价'}</button>
      {rows.map((row,index)=><div className="rounded-md bg-muted p-3 space-y-2" key={index}>
        <div className="grid grid-cols-2 gap-2">{(['vendorName','amount','taxRate','remark'] as const).map(key=><label className="text-caption" key={key}>{ {vendorName:'厂商名称',amount:'含税报价',taxRate:'税率（%）',remark:'备注'}[key] }<input aria-label={`第${index+1}家${{vendorName:'厂商名称',amount:'含税报价',taxRate:'税率',remark:'备注'}[key]}`} className={`${inputClass} ${key==='amount'||key==='taxRate'?'tabular-nums':''}`} disabled={busy} type={key==='amount'||key==='taxRate'?'number':'text'} value={row[key]} onChange={event=>setRows(previous=>previous.map((item,i)=>i===index?{...item,[key]:key==='amount'||key==='taxRate'?Number(event.target.value):event.target.value}:item))}/></label>)}</div>
        <p className="text-caption">已保存截图：{row.images.length} 张 · 保存时按原有金额编辑规则封顶。</p>
        <div className="flex flex-wrap gap-2">{row.images.map(image=><SavedQuoteImage key={image.assetId} assetId={image.assetId} target={target}/>)}</div>
        <label className="block rounded-md bg-card p-3 text-caption">上传第{index+1}家报价截图（保存后入库）<input type="file" accept="image/png,image/jpeg,image/webp" disabled={busy} className="block w-full pt-2" onChange={event=>{void chooseImage(event.target.files?.[0],index);event.target.value='';}}/></label>
        {uploads.filter(image=>image.row===index).map((image,i)=><figure key={i} className="rounded-md bg-card p-2"><img src={image.base64Data} alt={`第${index+1}家待保存报价截图`} className="max-h-36 max-w-full object-contain"/><figcaption className="text-caption">{image.name} · 待保存</figcaption><button disabled={busy} type="button" className="text-caption" onClick={()=>setUploads(previous=>previous.filter(item=>item!==image))}>移除待保存截图</button></figure>)}
      </div>)}
      {rows.length>0&&<button type="button" disabled={disabled||busy} className="rounded-md bg-primary px-4 py-2 text-primary-foreground disabled:opacity-50" onClick={()=>void act(false)}>保存询价与截图</button>}
    </>}
    {error&&<p role="alert" className="text-destructive whitespace-pre-wrap">{error}</p>}
    <button type="button" disabled={busy} className="rounded-md bg-card px-3 py-2 text-caption" onClick={()=>setRevision(value=>value+1)}>重新读取（放弃卡片草稿）</button>
  </section>;
}
