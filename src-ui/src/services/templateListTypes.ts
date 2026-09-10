export interface TechItem { serviceName: string; serviceDesc: string; amount: string | number; unit: string }
export interface QuoteImage { assetId: string; width?: number; height?: number }
export interface InquiryVendor { vendorName: string; amount: number; taxRate: number; remark: string; images: QuoteImage[] }
export interface PendingQuoteImage { row: number; base64Data: string; name: string; width: number; height: number }
export interface ListSnapshot { techItems: TechItem[]; inqVendors: InquiryVendor[] }
export interface SharedTechTarget { templateName: string; expected: TechItem[] }
export type TemplateListAction =
  | { type: 'readLists' }
  | { type: 'saveTech'; rows: TechItem[]; expected: TechItem[]; sharedTemplates: SharedTechTarget[] }
  | { type: 'generateInquiry'; expected: InquiryVendor[] }
  | { type: 'saveInquiry'; rows: InquiryVendor[]; expected: InquiryVendor[]; uploads: PendingQuoteImage[] };
export function listSnapshot(techItems: TechItem[], inqVendors: InquiryVendor[]): ListSnapshot {
  return { techItems: techItems.map(row => ({serviceName:row.serviceName,serviceDesc:row.serviceDesc,amount:row.amount,unit:row.unit})),
    inqVendors: inqVendors.map(row => ({vendorName:row.vendorName,amount:row.amount,taxRate:row.taxRate,remark:row.remark,
      images:(row.images || []).map(img => ({assetId:img.assetId,width:img.width,height:img.height}))})) };
}
export function assertListAction(action: TemplateListAction, current: ListSnapshot) {
  if (action.type === 'readLists') return;
  const actual = action.type === 'saveTech' ? current.techItems : current.inqVendors;
  if (JSON.stringify(actual) !== JSON.stringify(action.expected)) throw new Error('清单已在模板页或其他窗口修改，请重新读取后核对；没有覆盖现有清单。');
  if (action.type === 'saveTech' && (action.rows.length > 100 || action.rows.some(row =>
    typeof row.serviceName !== 'string' || typeof row.serviceDesc !== 'string' || typeof row.unit !== 'string'
      || !Number.isFinite(Number(row.amount)) || Number(row.amount) < 0))) throw new Error('请检查清单：最多100行，数量必须为非负数。');
  if (action.type === 'saveInquiry' && (action.rows.length !== current.inqVendors.length
    || action.rows.some((row,index) => JSON.stringify(row.images) !== JSON.stringify(current.inqVendors[index].images)))) throw new Error('询价卡片不能增删报价行或替换现有证据，请重新读取。');
}
