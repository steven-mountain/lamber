import type { AiImageAttachment } from '../../ai/types';
import AppIcon from '../icons/AppIcon';

interface ImageAttachmentPreviewProps {
  images: AiImageAttachment[];
  onRemove: (id: string) => void;
}

export default function ImageAttachmentPreview({ images, onRemove }: ImageAttachmentPreviewProps) {
  if (images.length === 0) return null;

  return (
    <div className="flex gap-2 overflow-x-auto px-1 pb-2">
      {images.map((image) => (
        <div key={image.id} className="relative h-16 w-16 flex-shrink-0 overflow-hidden rounded-lg border border-border bg-muted">
          {image.dataUrl ? (
            <img src={image.dataUrl} alt={image.name} className="h-full w-full object-cover" />
          ) : (
            <div className="flex h-full w-full flex-col items-center justify-center bg-secondary px-1 text-center text-[10px] font-semibold leading-3 text-secondary-foreground">
              <AppIcon name="imageUpload" size={16} />
              <span className="mt-1 truncate">{image.name}</span>
            </div>
          )}
          <button
            type="button"
            onClick={() => onRemove(image.id)}
            className="absolute right-1 top-1 flex h-5 w-5 items-center justify-center rounded-full bg-black/65 text-white transition-colors hover:bg-black/80"
            title="移除图片"
          >
            <AppIcon name="close" size={12} />
          </button>
        </div>
      ))}
    </div>
  );
}
