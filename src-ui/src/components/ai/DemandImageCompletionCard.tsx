import { saveChatDemandUpload } from '../../services/chatDemandUpload';
import { useEffect, useRef, useState } from 'react';
import { type DemandUploadTarget } from '../../services/demandUploadTargets';
import type { AiImageAttachment } from '../../ai/types';
import ImageAttachmentPreview from './ImageAttachmentPreview';

interface Props { target: DemandUploadTarget; disabled: boolean; onReceipt: (message: string) => void }

export default function DemandImageCompletionCard({ target, disabled, onReceipt }: Props) {
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [preview, setPreview] = useState<AiImageAttachment[]>([]);
  const active = useRef(true);
  const uploading = useRef(false);
  useEffect(() => { active.current = true; return () => { active.current = false; }; }, []);
  const upload = async (file: File | undefined) => {
    if (!file || !target.usage || disabled || uploading.current) return;
    uploading.current = true;
    setBusy(true); setError('');
    let assetId: string | undefined;
    try {
      if (file.size > 20 * 1024 * 1024) throw new Error('图片大小不能超过 20MB');
      if (!['image/png', 'image/jpeg', 'image/webp'].includes(file.type)) throw new Error('仅支持 PNG、JPEG、WEBP 格式图片');
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = () => reject(new Error('图片读取失败'));
        reader.readAsDataURL(file);
      });
      const dimensions = await new Promise<{ width: number; height: number }>((resolve, reject) => {
        const image = new Image();
        image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
        image.onerror = () => reject(new Error('图片无法解码，请选择有效的图片文件'));
        image.src = dataUrl;
      });
      const saved = await saveChatDemandUpload(target, { name: file.name, dataUrl, ...dimensions }, () => active.current);
      assetId = saved.assetId;
      const path = saved.path;
      onReceipt(`【系统回执】已存入：${target.projectName} / ${target.templateName} / ${target.label}。\n实际位置：${path}`);
      if (active.current) {
        setPreview([{ id: assetId, name: file.name, mimeType: file.type, size: file.size, dataUrl }]);
      }
    } catch (reason) {
      if (active.current) setError(`${assetId ? '图片已入库，但刷新失败，请回模板页查看。' : '上传失败：'}${String(reason)}`);
    } finally {
      uploading.current = false;
      if (active.current) setBusy(false);
    }
  };
  return <div className="space-y-3 rounded-lg bg-muted/50 p-4 text-caption">
    <p className="font-semibold">需求表附件补齐 · {target.projectName}</p>
    <p className="break-all text-muted-foreground">{target.templateName} · 按已保存表单检查</p>
    <div className="space-y-2 rounded-md bg-card p-3" tabIndex={0}
      onPaste={event => { const file = Array.from(event.clipboardData.files)[0]; if (file) { event.preventDefault(); void upload(file); } }}>
      <p>{target.label}</p>
      <label className="inline-block cursor-pointer rounded-md bg-muted px-3 py-2">
        选择图片并存入
        <input type="file" className="sr-only" accept="image/png,image/jpeg,image/webp" disabled={disabled || busy}
          onChange={event => { const file = event.target.files?.[0]; event.target.value = ''; void upload(file); }} />
      </label>
      <p className="text-muted-foreground">也可聚焦此卡片后粘贴。PNG / JPEG / WEBP，最大 20MB。</p>
    </div>
    {busy && <p role="status">正在存入图片…</p>}
    {error && <p role="alert" className="text-destructive">{error}</p>}
    <ImageAttachmentPreview images={preview} onRemove={id => setPreview(current => current.filter(image => image.id !== id))} />
    {preview.length > 0 && <p className="text-muted-foreground">关闭预览不删除资产；移除已存图片请在模板页操作。</p>}
  </div>;
}
