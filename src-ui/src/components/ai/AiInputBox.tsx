import React, { useEffect, useRef, useState } from 'react';
import type { AiImageAttachment } from '../../ai/types';
import ImageAttachmentPreview from './ImageAttachmentPreview';
import AppIcon from '../icons/AppIcon';

const MAX_IMAGE_SIZE = 5 * 1024 * 1024;
const MAX_IMAGE_COUNT = 4;
const ACCEPTED_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/webp']);

interface AiInputBoxProps {
  input: string;
  images: AiImageAttachment[];
  isTyping: boolean;
  visionEnabled: boolean;
  onInputChange: (value: string) => void;
  onImagesChange: (images: AiImageAttachment[]) => void;
  onSend: () => void;
  onStop: () => void;
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function createAttachmentId() {
  return typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export default function AiInputBox({
  input,
  images,
  isTyping,
  visionEnabled,
  onInputChange,
  onImagesChange,
  onSend,
  onStop,
}: AiInputBoxProps) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [attachmentError, setAttachmentError] = useState('');

  const autoResize = () => {
    const el = textareaRef.current;
    if (!el) return;

    el.style.height = 'auto';
    const nextHeight = Math.min(Math.max(el.scrollHeight, 44), 180);
    el.style.height = `${nextHeight}px`;
  };

  useEffect(() => {
    autoResize();
  }, [input]);

  const addImageFiles = async (files: File[], source: 'select' | 'paste') => {
    if (files.length === 0) return;

    if (!visionEnabled || isTyping) {
      setAttachmentError('请先在模型设置中启用图片输入。');
      return;
    }

    const remainingSlots = Math.max(MAX_IMAGE_COUNT - images.length, 0);
    if (remainingSlots === 0) {
      setAttachmentError(`最多添加 ${MAX_IMAGE_COUNT} 张图片。`);
      return;
    }

    const validFiles = files
      .filter((file) => ACCEPTED_IMAGE_TYPES.has(file.type))
      .filter((file) => file.size <= MAX_IMAGE_SIZE)
      .slice(0, remainingSlots);

    if (validFiles.length === 0) {
      setAttachmentError(source === 'paste' ? '剪贴板中没有 5MB 以内的 png、jpg 或 webp 图片。' : '请选择 5MB 以内的 png、jpg 或 webp 图片。');
      return;
    }

    const nextImages = await Promise.all(
      validFiles.map(async (file) => ({
        id: createAttachmentId(),
        name: file.name,
        mimeType: file.type,
        size: file.size,
        dataUrl: await fileToDataUrl(file),
        source: 'user_upload' as const,
      }))
    );

    onImagesChange([...images, ...nextImages].slice(0, MAX_IMAGE_COUNT));
    setAttachmentError(files.length > validFiles.length ? `已添加 ${validFiles.length} 张有效图片。` : '');
  };

  const handleImageSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || []);
    event.target.value = '';
    await addImageFiles(files, 'select');
  };

  const handlePaste = async (event: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const pastedImages = Array.from(event.clipboardData.items)
      .filter((item) => item.kind === 'file' && item.type.startsWith('image/'))
      .map((item) => item.getAsFile())
      .filter((file): file is File => Boolean(file));

    if (pastedImages.length === 0) return;

    event.preventDefault();
    await addImageFiles(pastedImages, 'paste');
  };

  const removeImage = (id: string) => {
    onImagesChange(images.filter((image) => image.id !== id));
    setAttachmentError('');
  };

  const handleKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.key === 'Enter' && event.ctrlKey) {
      event.preventDefault();
      onSend();
    }
  };

  const canSend = !isTyping && (input.trim().length > 0 || images.length > 0);

  return (
    <div className="space-y-2">
      <ImageAttachmentPreview images={images} onRemove={removeImage} />

      {attachmentError && (
        <div className="px-1 text-[11px] font-medium text-destructive">
          {attachmentError}
        </div>
      )}

      <div className="flex items-end gap-2">
        {visionEnabled && !isTyping && (
          <>
            <input
              ref={fileInputRef}
              type="file"
              accept="image/png,image/jpeg,image/webp"
              multiple
              className="hidden"
              onChange={handleImageSelect}
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-xl border border-border bg-muted text-muted-foreground transition-colors hover:bg-primary/10 hover:text-primary"
              title="添加图片"
            >
              <AppIcon name="imageUpload" size={18} />
            </button>
          </>
        )}

        <textarea
          ref={textareaRef}
          value={input}
          onChange={(event) => {
            onInputChange(event.target.value);
            requestAnimationFrame(autoResize);
          }}
          onKeyDown={handleKeyDown}
          onPaste={handlePaste}
          placeholder={isTyping ? 'AI 正在思考中...' : '输入问题，按 Ctrl+Enter 发送...'}
          disabled={isTyping}
          className="min-h-[44px] max-h-[180px] flex-1 resize-none overflow-y-auto rounded-xl border border-border bg-muted px-4 py-2.5 text-sm shadow-inner outline-none transition-colors focus:border-ring disabled:opacity-70"
          rows={1}
        />

        {isTyping ? (
          <button
            type="button"
            onClick={onStop}
            className="flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-xl text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
            title="停止生成"
          >
            <AppIcon name="close" size={20} />
          </button>
        ) : (
          <button
            type="button"
            onClick={onSend}
            disabled={!canSend}
            className="flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground shadow-sm transition-all hover:shadow-md active:scale-95 disabled:cursor-not-allowed disabled:opacity-50"
            title="发送"
          >
            <AppIcon name="send" size={18} className={canSend ? 'ml-0.5' : ''} />
          </button>
        )}
      </div>
    </div>
  );
}
