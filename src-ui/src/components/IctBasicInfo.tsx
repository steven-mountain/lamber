import React from "react";
import { CommonPresetFieldHeader, CommonPresetLabelHeader } from "./common-presets/CommonPresetQuickFill";
import { PRESET_FIELD_KEYS } from "../lib/presetFieldKeys";

interface IctBasicInfoProps {
  state: any;
  calculations: any;
}

export const IctBasicInfo: React.FC<IctBasicInfoProps> = ({ state }) => {
  const {
    projName, setProjName,
    customerName, setCustomerName,
    propertyRights, setPropertyRights,
    discountRate, setDiscountRate,
    projectYears, setProjectYears,
    projectBackground, setProjectBackground,
  } = state;

  return (
    <div className="bg-card border border-border rounded-xl p-8 shadow-sm">
      <h3 className="text-lg font-bold text-foreground mb-6">项目概况</h3>
      <div className="grid grid-cols-2 gap-6">
        <div className="flex flex-col gap-2">
          <CommonPresetLabelHeader labelClassName="text-sm font-bold text-secondary-foreground">项目名称</CommonPresetLabelHeader>
          <input id="ict-proj-name" type="text" value={projName} onChange={e => setProjName(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <CommonPresetFieldHeader
            fieldKey={PRESET_FIELD_KEYS.projectCustomerName}
            kind="short_value"
            value={customerName}
            onApply={setCustomerName}
            labelClassName="text-sm font-bold text-secondary-foreground"
          >
            客户单位名称
          </CommonPresetFieldHeader>
          <input id="ict-customer-name" type="text" value={customerName} onChange={e => setCustomerName(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <CommonPresetFieldHeader
            fieldKey={PRESET_FIELD_KEYS.projectPropertyRights}
            kind="short_value"
            value={propertyRights}
            onApply={setPropertyRights}
            labelClassName="text-sm font-bold text-secondary-foreground"
          >
            产权归属
          </CommonPresetFieldHeader>
          <input id="ict-property-rights" type="text" value={propertyRights} onChange={e => setPropertyRights(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <CommonPresetLabelHeader labelClassName="text-sm font-bold text-secondary-foreground">项目建设/服务周期 (年)</CommonPresetLabelHeader>
          <input id="ict-project-years" type="number" min={1} max={10} value={projectYears} onChange={e => setProjectYears(Number(e.target.value))} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col gap-2">
          <CommonPresetLabelHeader labelClassName="text-sm font-bold text-secondary-foreground">折现率</CommonPresetLabelHeader>
          <input id="ict-discount-rate" type="number" step={0.001} value={discountRate} onChange={e => setDiscountRate(Number(e.target.value))} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
        <div className="flex flex-col justify-end gap-2 rounded-lg bg-muted/50 px-4 py-3">
          <span className="text-xs font-bold text-secondary-foreground">现金流依据</span>
          <span className="text-sm font-extrabold text-foreground">科目收付款计划</span>
        </div>
        <div className="flex flex-col gap-2 col-span-2">
          <CommonPresetFieldHeader
            fieldKey={PRESET_FIELD_KEYS.projectBackground}
            kind="text_snippet"
            value={projectBackground}
            onApply={setProjectBackground}
            labelClassName="text-sm font-bold text-secondary-foreground"
          >
            项目背景
          </CommonPresetFieldHeader>
          <textarea id="ict-project-bg" rows={3} value={projectBackground} onChange={e => setProjectBackground(e.target.value)} className="bg-card border border-input px-3.5 py-2.5 rounded-md outline-none focus:border-ring" />
        </div>
      </div>
    </div>
  );
};
