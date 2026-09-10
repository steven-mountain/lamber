/** Current-turn UI invitation only; no authorization to replace an asset. */
export function wantsTemplateImages(content: string): boolean {
  const text = content.replace(/```[\s\S]*?(?:```|$)/g, '').replace(/^\s*>.*$/gm, '')
    .replace(/[“「"][\s\S]*?[”」"]/g, '');
  if (/(?:不要|不用|无需|别|暂不|先不|不想).{0,12}(?:图片|附件|展示|替换|修改)/.test(text)) return false;
  return text.split(/[。！？\n]/).some(part =>
    !/(?:为什么|如何|怎么|昨天|之前|如果|客户说|是什么意思)/.test(part)
    && /(?:展示|显示|查看|看看|调出|打开|替换|更换|修改|换掉|给我看)/.test(part)
    && /(?:图片|附件[12一二]?|客户确认材料|公示截图|asset_[a-z0-9]+)/i.test(part));
}

export function templateImagePrompt(requested: boolean): string {
  return requested
    ? 'The user explicitly requests saved project images. The application shows a project image card using trusted session binding. It lists saved demand-form attachments including existing images. Guide the user to view the card, select an exact image, and choose a replacement with before/after preview, then click Confirm replacement. No write occurs until that click. This card only views/replaces existing images; if the list is empty it has no Add/upload action. For missing attachments, explicitly request demand-form filling/generation to use the separate missing-image upload card, or upload on the template page. Uploading a new file elsewhere appends; it does not replace an existing image. The card can attach a selected saved image to the composer for visual analysis; only claim to see its contents when image bytes are actually sent with the user message. Do not call bash/glob/read_image or guess asset paths. No model image editing service is connected: describe crop/text-edit suggestions honestly, do not claim to have edited pixels.'
    : 'Saved image metadata is not image content. For an explicit request to view/replace saved images, the user can open the Project images button in chat or ask to show existing images. That card supports selecting a demand-form attachment, previewing a replacement and confirming it. Do not use filesystem tools or guess asset paths to retrieve images. Ordinary chat attachments are not automatically saved to templates; uploading normally appends, not replaces. No image pixel editing service is connected.';
}
