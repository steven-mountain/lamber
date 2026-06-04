import { useState, useRef, useEffect } from "react"
import { invoke, convertFileSrc } from "@tauri-apps/api/core"
import AppIcon from "../components/icons/AppIcon"
import { MID_THREE_CAPABILITIES } from "../lib/midThreeConstants"
import { useAiContextStore } from "../store/useAiContextStore"
import { buildAiContextKey } from "../utils/aiContextKeys"
import { projectFileService } from "../services/projectFileService"
import { projectService } from "../utils/projectService"
import { domainSaveService } from "../services/domainSaveService"
import { useSaveStore } from "../store/useSaveStore"
import { useWorkspaceStore } from "../store/useWorkspaceStore"
import { createTemplateAssetSelection, publishTemplateAssetSelection } from "../ai/templateAssetSelection"
import {
  buildExcelSubjectVariables,
  collectDocumentBusinessNames,
  getProjectDataSubjectItem,
  getSubjectDocumentBusinessName,
  ICT_SUBJECT_DEFINITIONS,
  resolveBillingSubjectPresentation,
  type IctDocumentPrefix,
  type IctSubjectGroupId,
  type IctSubjectSide,
} from "../lib/ictSubjectCatalog"

interface Props {
  selectedTemplate: string;
  projectData: any; // Basic/Rev/Cost data from IctLifecycle
  projectBackground: string;
  setProjectBackground: (val: string) => void;
  techItems: any[];
  setTechItems: React.Dispatch<React.SetStateAction<any[]>>;
  inqVendors: any[];
  setInqVendors: React.Dispatch<React.SetStateAction<any[]>>;
  metrics?: any;
  outputDir?: string;
  projectId?: string;
  onGenerated?: () => void;
}

type ProjectScale = "large" | "small"
type SelfThreeRequirement = "integration" | "maintenance"

const SELF_THREE_OPTIONS: Array<{
  value: string;
  reminder?: string;
  requirements: SelfThreeRequirement[];
}> = [
  {
    value: "自主集成，项目自主等级L1。",
    requirements: [],
  },
  {
    value: "自主集成，自主研发，自主运维，项目自主等级L3。",
    reminder: "需要包含集成和维保费以及自主-研发服务费",
    requirements: ["integration", "maintenance"],
  },
  {
    value: "自主集成，自主研发，自主运维，自主交付，项目自主等级L3。",
    reminder: "需要包含集成和维保费以及自主-研发服务费",
    requirements: ["integration", "maintenance"],
  },
  {
    value: "自主集成，自主研发，自主交付，项目自主等级L3。",
    reminder: "需要包含集成费以及自主-研发服务费",
    requirements: ["integration"],
  },
  {
    value: "自主集成，自主运维，项目自主等级L2。",
    reminder: "需要包含维保费",
    requirements: ["maintenance"],
  },
  {
    value: "自主集成，自主运维，自主交付，项目自主等级L2。",
    reminder: "需要包含集成和维保费",
    requirements: ["integration", "maintenance"],
  },
  {
    value: "自主集成，自主交付，项目自主等级L2。",
    reminder: "需要包含集成费",
    requirements: ["integration"],
  },
]

const normalizeProjectScale = (value?: FormDataEntryValue | string | null): ProjectScale => {
  return value === "small" ? "small" : "large"
}

const getTaxItemAmount = (item: any) => Number(item?.incl || 0) + Number(item?.excl || 0)

const getSelfThreeOption = (value: string) => SELF_THREE_OPTIONS.find(option => option.value === value) || SELF_THREE_OPTIONS[0]

const getSelfThreeMissingFees = (requirements: SelfThreeRequirement[], hasItIntegrationFee: boolean, hasItMaintenanceFee: boolean) => requirements.flatMap(requirement => {
  if (requirement === "integration" && !hasItIntegrationFee) return ["投资效益分析中未检测到 IT 集成服务费"]
  if (requirement === "maintenance" && !hasItMaintenanceFee) return ["投资效益分析中未检测到 IT 维保费（不包含 CT 专线维保）"]
  return []
})

const toDocImagePayload = (img: any, title?: string) => ({
  title: title || img?.title || "",
  data: img?.assetId || img?.data || "",
  assetId: img?.assetId || null,
  width: img?.width,
  height: img?.height,
})

const mergeVendorImages = (nextVendors: any[], previousVendors: any[]) => {
  const imagesByName = new Map<string, any[]>()
  previousVendors.forEach(vendor => {
    const name = String(vendor?.vendorName || "").trim()
    if (name && vendor?.images?.length) {
      imagesByName.set(name, vendor.images)
    }
  })

  return nextVendors.map((vendor, index) => {
    const name = String(vendor?.vendorName || "").trim()
    const previousImages = (name && imagesByName.get(name)) || previousVendors[index]?.images || []
    return { ...vendor, images: previousImages }
  })
}

export default function TemplateForms({
  selectedTemplate,
  projectData,
  projectBackground,
  setProjectBackground,
  techItems,
  setTechItems,
  inqVendors,
  setInqVendors,
  metrics,
  outputDir,
  projectId,
  onGenerated
}: Props) {
  const formRef = useRef<HTMLFormElement>(null)
  const markDirty = useSaveStore(state => state.markDirty)
  const clearDirty = useSaveStore(state => state.clearDirty)
  const registerSaveHandler = useSaveStore(state => state.registerSaveHandler)
  const unregisterSaveHandler = useSaveStore(state => state.unregisterSaveHandler)
  const workspaceId = useWorkspaceStore(state => state.workspaceId)

  // Specific state for dynamic toggles
  const [projectScale, setProjectScale] = useState<ProjectScale>("large")
  const [hasMidThree, setHasMidThree] = useState(true)
  const [hasSingleSource, setHasSingleSource] = useState(false)
  const [procurementMethod, setProcurementMethod] = useState("短名单甄选")
  const [hasPublicUrl, setHasPublicUrl] = useState(false)
  const [hasSecurity, setHasSecurity] = useState(false)

  // Images state
  const [attach1Images, setAttach1Images] = useState<any[]>([])
  const [attach2Images, setAttach2Images] = useState<any[]>([])

  const fileInput1Ref = useRef<HTMLInputElement>(null)
  const fileInput2Ref = useRef<HTMLInputElement>(null)

  const todayStr = (() => {
    const d = new Date();
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  })();

  // Linkage States
  const [itContent, setItContent] = useState("")
  const [ctContent, setCtContent] = useState("视频监控")
  const [midThreeCode, setMidThreeCode] = useState("A302600342")
  const [midThreeName, setMidThreeName] = useState("视频监控能力")
  const [itBusMode, setItBusMode] = useState("服务购销")
  const [itFundSrc, setItFundSrc] = useState("分公司成本开支")
  const [revCollection, setRevCollection] = useState("项目验收完成后30天内客户单位支付100%")
  const [expPayment, setExpPayment] = useState("项目验收完成且收到款项后30天内支付100%")
  const [selfThreeValue, setSelfThreeValue] = useState(SELF_THREE_OPTIONS[0].value)
  const [syncTrigger, setSyncTrigger] = useState(0) // Added to trigger AI sync on ref changes

  const [isMidThreeModalOpen, setIsMidThreeModalOpen] = useState(false)
  const [midThreeSearch, setMidThreeSearch] = useState("")

  const [subjectItCost] = useState("IT集成")
  const [subjectCtCost, setSubjectCtCost] = useState("CT-视频监控")
  const [subjectItRev] = useState("小微ICT业务-IoT-集成")
  const [subjectCtRev, setSubjectCtRev] = useState("CT-视频监控")

  const selectedSelfThree = getSelfThreeOption(selfThreeValue)
  const hasItIntegrationFee = getTaxItemAmount(projectData.cost?.it?.integration) > 0
  const hasItMaintenanceFee = getTaxItemAmount(projectData.cost?.it?.maintenance) > 0
  const selfThreeMissingFees = getSelfThreeMissingFees(selectedSelfThree.requirements, hasItIntegrationFee, hasItMaintenanceFee)
  const totalRevenueIncl = [
    projectData.revenue?.it?.integration,
    projectData.revenue?.it?.maintenance,
    projectData.revenue?.it?.device_sales,
    projectData.revenue?.it?.device_lease,
    projectData.revenue?.it?.other,
    projectData.revenue?.it?.cloud,
    projectData.revenue?.ct?.line,
    projectData.revenue?.ct?.product,
    projectData.revenue?.non_it_ct,
  ].reduce((sum, item) => sum + Number(item?.incl || 0), 0)
  const getBusinessNames = (options: { side?: IctSubjectSide; documentPrefix?: IctDocumentPrefix; groupId?: IctSubjectGroupId }) => collectDocumentBusinessNames(projectData, options)
  const getSubjectBusinessName = (subjectCode: string) => {
    const subject = ICT_SUBJECT_DEFINITIONS.find(item => item.subjectCode === subjectCode)
    if (!subject) return ""
    return getSubjectDocumentBusinessName(subject, getProjectDataSubjectItem(projectData, subject))
  }
  const customItBusinessNames = getBusinessNames({ documentPrefix: "IT" })
  const customCtBusinessNames = getBusinessNames({ documentPrefix: "CT" })
  const customNonItCtBusinessNames = getBusinessNames({ documentPrefix: "非IT/CT" })
  const customMixBusinessNames = getBusinessNames({ documentPrefix: "综合类" })
  const customItCostBusinessNames = getBusinessNames({ side: "cost", documentPrefix: "IT" })
  const customCtCostBusinessNames = getBusinessNames({ side: "cost", documentPrefix: "CT" })
  const customItRevenueBusinessNames = getBusinessNames({ side: "revenue", documentPrefix: "IT" })
  const customCtRevenueBusinessNames = getBusinessNames({ side: "revenue", documentPrefix: "CT" })
  const joinedBusinessNames = (names: string[]) => names.join("、")
  const excelSubjectVariables = buildExcelSubjectVariables(projectData)

  const formDataRef = useRef<Record<string, string>>({});
  const [formData, setFormData] = useState<Record<string, string>>({});

  const handleFieldChange = (name: string, value: string) => {
    formDataRef.current[name] = value;
    setFormData({ ...formDataRef.current });
    setSyncTrigger(prev => prev + 1);
  };

  const getBind = (name: string, defaultVal: string = "") => {
    return {
      value: formData[name] ?? defaultVal,
      onChange: (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        handleFieldChange(name, e.target.value);
      }
    };
  };

  const getBindCheckbox = (name: string, defaultVal: boolean = false) => {
    const val = formData[name];
    const checked = val !== undefined ? (val === "true" || val === "on") : defaultVal;
    return {
      checked,
      onChange: (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.checked) {
          handleFieldChange(name, "on");
          return;
        }
        delete formDataRef.current[name];
        setFormData({ ...formDataRef.current });
        setSyncTrigger(prev => prev + 1);
      }
    };
  };

  const handleFormChange = (e: any) => {
    const target = e.target;
    if (target && target.name && target.name.startsWith('gen_')) {
      if (target.type === 'checkbox') {
        if (target.checked) {
          formDataRef.current[target.name] = "on";
        } else {
          delete formDataRef.current[target.name];
        }
      } else {
        formDataRef.current[target.name] = target.value;
      }
      setFormData({ ...formDataRef.current });
      if (target.name === 'gen_project_scale') {
        setProjectScale(normalizeProjectScale(target.value));
      }
      setSyncTrigger(prev => prev + 1); // Trigger the debounced sync effect
    }
  };

  const isLoadingRef = useRef(false);
  const autoSaveTimerRef = useRef<NodeJS.Timeout | null>(null);

  // --- Load Form Settings from DB ---
  const resolveImages = async (imagesList: any[]) => {
    if (!imagesList) return [];
    return Promise.all(imagesList.map(async (img: any) => {
      if (img.assetId && !img.data) {
        try {
          const path = await projectService.getTemplateAssetPath(img.assetId);
          const dataUrl = convertFileSrc(path);
          return {
            assetId: img.assetId,
            data: dataUrl,
            width: img.width,
            height: img.height,
          };
        } catch (e) {
          console.warn("Failed to get asset path for id:", img.assetId, e);
          return {
            assetId: img.assetId,
            data: "",
            width: img.width,
            height: img.height,
            error: true,
          };
        }
      }
      return img;
    }));
  };

  const resetFormToDefaults = () => {
    setItContent("");
    setCtContent("视频监控");
    setMidThreeCode("A302600342");
    setMidThreeName("视频监控能力");
    setItBusMode("服务购销");
    setItFundSrc("分公司成本开支");
    setRevCollection("项目验收完成后30天内客户单位支付100%");
    setExpPayment("项目验收完成且收到款项后30天内支付100%");
    setSelfThreeValue(SELF_THREE_OPTIONS[0].value);
    setProjectScale("large");
    setHasMidThree(true);
    setHasSingleSource(false);
    setProcurementMethod("短名单甄选");
    setHasPublicUrl(false);
    setHasSecurity(false);
    setTechItems([]);
    setInqVendors([]);
    setAttach1Images([]);
    setAttach2Images([]);
    formDataRef.current = {};
    setFormData({});
    if (formRef.current) {
      formRef.current.reset();
    }
    
    let name = projectData.basic?.proj_name || ""
    name = name.replace(/项目/g, "")
    if (name && !name.includes("服务")) name += "服务"
    setItContent(name)
  };

  const loadFormSettings = async () => {
    if (!projectId || !selectedTemplate) {
      resetFormToDefaults();
      return;
    }
    isLoadingRef.current = true;
    try {
      const savedState = await domainSaveService.loadTemplateState(projectId, selectedTemplate);
      if (savedState?.filledDataJson) {
        const parsed = savedState.filledDataJson as any;
        if (parsed.itContent !== undefined) setItContent(parsed.itContent);
        if (parsed.ctContent !== undefined) setCtContent(parsed.ctContent);
        if (parsed.midThreeCode !== undefined) setMidThreeCode(parsed.midThreeCode);
        if (parsed.midThreeName !== undefined) setMidThreeName(parsed.midThreeName);
        if (parsed.itBusMode !== undefined) setItBusMode(parsed.itBusMode);
        if (parsed.itFundSrc !== undefined) setItFundSrc(parsed.itFundSrc);
        if (parsed.revCollection !== undefined) setRevCollection(parsed.revCollection);
        if (parsed.expPayment !== undefined) setExpPayment(parsed.expPayment);
        if (parsed.selfThreeValue !== undefined) setSelfThreeValue(parsed.selfThreeValue);
        if (parsed.projectScale !== undefined) {
          setProjectScale(parsed.projectScale);
          formDataRef.current.gen_project_scale = parsed.projectScale;
        }
        if (parsed.hasMidThree !== undefined) setHasMidThree(parsed.hasMidThree);
        if (parsed.hasSingleSource !== undefined) setHasSingleSource(parsed.hasSingleSource);
        if (parsed.procurementMethod !== undefined) setProcurementMethod(parsed.procurementMethod);
        if (parsed.hasPublicUrl !== undefined) setHasPublicUrl(parsed.hasPublicUrl);
        if (parsed.hasSecurity !== undefined) setHasSecurity(parsed.hasSecurity);
        if (parsed.techItems !== undefined) setTechItems(parsed.techItems);
        
        if (parsed.formData) {
          formDataRef.current = { ...parsed.formData };
          setFormData({ ...parsed.formData });
          if (formRef.current) {
            setTimeout(() => {
              if (formRef.current) {
                Object.entries(parsed.formData).forEach(([name, val]) => {
                  const el = formRef.current?.querySelector(`[name="${name}"]`) as HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;
                  if (el) {
                    el.value = val as string;
                  }
                });
              }
            }, 0);
          }
        } else {
          formDataRef.current = {};
          setFormData({});
        }

        if (parsed.inqVendors) {
          const resolvedVendors = await Promise.all(parsed.inqVendors.map(async (v: any) => {
            const resolvedImages = await resolveImages(v.images || []);
            return { ...v, images: resolvedImages };
          }));
          setInqVendors(resolvedVendors);
        } else {
          setInqVendors([]);
        }

        const resolvedAttach1 = await resolveImages(parsed.attach1Images || []);
        setAttach1Images(resolvedAttach1);

        const resolvedAttach2 = await resolveImages(parsed.attach2Images || []);
        setAttach2Images(resolvedAttach2);
      } else {
        resetFormToDefaults();
      }
    } catch (err) {
      console.error("Failed to load template settings", err);
      resetFormToDefaults();
    } finally {
      setTimeout(() => {
        isLoadingRef.current = false;
      }, 100);
    }
  };

  useEffect(() => {
    loadFormSettings();
  }, [projectId, selectedTemplate]);

  // --- Auto-save and image migration ---
  const ensureAllImagesMigrated = async (imagesList: any[], typeName: string): Promise<any[]> => {
    if (!imagesList || !projectId) return [];
    return Promise.all(imagesList.map(async (img) => {
      if (!img.assetId && img.data && img.data.startsWith("data:image/")) {
        try {
          const assetId = await domainSaveService.saveTemplateAsset(projectId, selectedTemplate, {
            assetType: "image",
            usage: typeName,
            originalFileName: "migrated_image",
            base64Data: img.data,
            width: img.width || null,
            height: img.height || null,
          });
          return {
            assetId,
            data: img.data,
            width: img.width,
            height: img.height
          };
        } catch (err) {
          console.error("Migration failed for image:", err);
          return img;
        }
      }
      return img;
    }));
  };

  const autoSaveFormSettings = async (options: { throwOnError?: boolean } = {}) => {
    if (!projectId || !selectedTemplate || isLoadingRef.current) {
      if (options.throwOnError) throw new Error("模板表单尚未准备好，无法保存");
      return false;
    }

    const migratedAttach1 = await ensureAllImagesMigrated(attach1Images, "attach1");
    const migratedAttach2 = await ensureAllImagesMigrated(attach2Images, "attach2");
    
    const migratedVendors = await Promise.all(inqVendors.map(async (v, idx) => {
      if (v.images && v.images.length > 0) {
        const migratedImgs = await ensureAllImagesMigrated(v.images, `vendor_${idx}`);
        return { ...v, images: migratedImgs };
      }
      return v;
    }));

    let changed = false;
    migratedAttach1.forEach((img, idx) => {
      if (img.assetId !== attach1Images[idx]?.assetId) changed = true;
    });
    migratedAttach2.forEach((img, idx) => {
      if (img.assetId !== attach2Images[idx]?.assetId) changed = true;
    });
    if (changed) {
      setAttach1Images(migratedAttach1);
      setAttach2Images(migratedAttach2);
    }

    let vendorsChanged = false;
    migratedVendors.forEach((v, idx) => {
      const oldV = inqVendors[idx];
      if (oldV && oldV.images && v.images) {
        if (oldV.images.length !== v.images.length) vendorsChanged = true;
        else {
          v.images.forEach((img: any, imgIdx: number) => {
            if (img.assetId !== oldV.images[imgIdx]?.assetId) vendorsChanged = true;
          });
        }
      }
    });
    if (vendorsChanged) {
      setInqVendors(migratedVendors);
    }

    const stripData = (imgs: any[]) => imgs.map(img => ({
      assetId: img.assetId,
      width: img.width,
      height: img.height
    }));

    const payload = {
      itContent,
      ctContent,
      midThreeCode,
      midThreeName,
      itBusMode,
      itFundSrc,
      revCollection,
      expPayment,
      selfThreeValue,
      projectScale,
      hasMidThree,
      hasSingleSource,
      procurementMethod,
      hasPublicUrl,
      hasSecurity,
      techItems,
      inqVendors: migratedVendors.map(v => ({
        ...v,
        images: stripData(v.images || [])
      })),
      attach1Images: stripData(migratedAttach1),
      attach2Images: stripData(migratedAttach2),
      formData: formDataRef.current,
    };

    try {
      await domainSaveService.saveTemplateState(projectId, selectedTemplate, {
        templateName: selectedTemplate,
        templateType: selectedTemplate.endsWith(".xlsx") ? "excel" : "word",
        templatePath: selectedTemplate,
        templatePathType: "module",
        filledDataJson: payload,
        fieldMappingJson: {},
        outputConfigJson: { outputDir: outputDir || null },
      });
      clearDirty("template-forms");
      return true;
    } catch (err) {
      console.error("Failed to auto-save template settings", err);
      if (options.throwOnError) {
        throw err;
      }
      return false;
    }
  };

  useEffect(() => {
    if (!projectId || !selectedTemplate) return;
    const registeredWorkspaceId = workspaceId;
    const registeredProjectId = projectId;
    const registeredTemplate = selectedTemplate;
    registerSaveHandler("template-forms", async (context) => {
      if (context.workspaceId !== registeredWorkspaceId || context.projectId !== registeredProjectId || selectedTemplate !== registeredTemplate) {
        throw new Error("模板、项目或工作区已切换");
      }
      await autoSaveFormSettings({ throwOnError: true });
      return { success: true, savedScopes: ["template-forms"] };
    });
    return () => unregisterSaveHandler("template-forms");
  }, [
    projectId,
    selectedTemplate,
    workspaceId,
    registerSaveHandler,
    unregisterSaveHandler,
    autoSaveFormSettings,
  ]);

  useEffect(() => {
    if (!projectId || !selectedTemplate || isLoadingRef.current) return;
    markDirty("template-forms");
  }, [
    projectId,
    selectedTemplate,
    markDirty,
    itContent,
    ctContent,
    midThreeCode,
    midThreeName,
    itBusMode,
    itFundSrc,
    revCollection,
    expPayment,
    selfThreeValue,
    projectScale,
    hasMidThree,
    hasSingleSource,
    procurementMethod,
    hasPublicUrl,
    hasSecurity,
    techItems,
    inqVendors,
    attach1Images,
    attach2Images,
    syncTrigger
  ]);

  useEffect(() => {
    if (!projectId || !selectedTemplate || isLoadingRef.current) return;

    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);

    autoSaveTimerRef.current = setTimeout(() => {
      autoSaveFormSettings();
    }, 1000);

    return () => {
      if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    };
  }, [
    projectId,
    selectedTemplate,
    itContent,
    ctContent,
    midThreeCode,
    midThreeName,
    itBusMode,
    itFundSrc,
    revCollection,
    expPayment,
    selfThreeValue,
    projectScale,
    hasMidThree,
    hasSingleSource,
    procurementMethod,
    hasPublicUrl,
    hasSecurity,
    techItems,
    inqVendors,
    attach1Images,
    attach2Images,
    syncTrigger
  ]);

  // --- AI Context Sync for Templates ---
  const updateData = useAiContextStore(state => state.updateBusinessData);
  const syncTimerRef = useRef<NodeJS.Timeout | null>(null);

  const buildTemplateContextPayload = (overrides: Record<string, string> = {}) => {
    const form = formRef.current;
    const formEntries = form ? Object.fromEntries(new FormData(form).entries()) : {};
    const nextSelfThreeValue = overrides.gen_self_three || selfThreeValue;
    const nextSelfThree = getSelfThreeOption(nextSelfThreeValue);
    const nextMissingFees = getSelfThreeMissingFees(nextSelfThree.requirements, hasItIntegrationFee, hasItMaintenanceFee);

    return {
      projectId: projectId || null,
      selectedTemplate,
      ...formEntries,
      ...formDataRef.current,
      ...overrides,
      itContent,
      ctContent,
      midThreeName,
      midThreeCode,
      techItems,
      inqVendors: inqVendors.map(v => ({ vendorName: v.vendorName, amount: v.amount })),
      gen_project_scale: overrides.gen_project_scale || projectScale,
      gen_self_three: nextSelfThreeValue,
      self_three_selected: nextSelfThreeValue,
      self_three_reminder: nextSelfThree.reminder || "",
      self_three_missing_fees: nextMissingFees,
    };
  }

  const syncTemplateContextNow = (overrides: Record<string, string> = {}) => {
    const templateId = buildAiContextKey('ict', 'template', selectedTemplate);
    const payload = buildTemplateContextPayload(overrides);
    updateData(templateId, payload);
  }

  const handleProjectScaleChange = (value: ProjectScale) => {
    setProjectScale(value);
    formDataRef.current.gen_project_scale = value;
    setSyncTrigger(prev => prev + 1);
    syncTemplateContextNow({ gen_project_scale: value });
  }

  const handleSelfThreeChange = (value: string) => {
    setSelfThreeValue(value);
    formDataRef.current.gen_self_three = value;
    setSyncTrigger(prev => prev + 1);
    syncTemplateContextNow({ gen_self_three: value });
  }

  useEffect(() => {
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);

    syncTimerRef.current = setTimeout(() => {
      syncTemplateContextNow();
    }, 500);

    return () => {
      if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    };
  }, [selectedTemplate, itContent, ctContent, midThreeName, midThreeCode, techItems, inqVendors, syncTrigger, itBusMode, itFundSrc, revCollection, expPayment, projectScale, selfThreeValue, hasItIntegrationFee, hasItMaintenanceFee]);

  // -- Linkage Logic --
  useEffect(() => {
    // Only set initial name link if we are not loading saved settings
    if (isLoadingRef.current) return;
    let name = projectData.basic?.proj_name || ""
    name = name.replace(/项目/g, "")
    if (name && !name.includes("服务")) name += "服务"
    setItContent(name)
  }, [projectData.basic?.proj_name])

  useEffect(() => {
    if (isLoadingRef.current) return;
    if (hasMidThree) {
      const baseName = midThreeName.replace(/能力/g, "")
      setCtContent(baseName)
      setSubjectCtCost(`CT-${baseName}`)
      setSubjectCtRev(`CT-${baseName}`)
    } else {
      if (ctContent === midThreeName.replace(/能力/g, "")) setCtContent("")
      setSubjectCtCost("CT-专线")
      setSubjectCtRev("CT-专线")
    }
  }, [hasMidThree, midThreeName])

  // -- Dynamic Tables Functions --
  const addTechItem = () => setTechItems([...techItems, { serviceName: '', serviceDesc: '', amount: 1, unit: '项' }])
  const updateTechItem = (i: number, key: string, val: string|number) => {
    const newItems = [...techItems]
    newItems[i] = { ...newItems[i], [key]: val }
    setTechItems(newItems)
  }
  const removeTechItem = (i: number) => setTechItems(techItems.filter((_, idx) => idx !== i))

  const addInqVendor = () => setInqVendors([...inqVendors, { vendorName: '', amount: 0, taxRate: 6, remark: '', images: [] }])
  const updateInqVendor = (i: number, key: string, val: string|number) => {
    const newItems = [...inqVendors]
    newItems[i] = { ...newItems[i], [key]: val }
    setInqVendors(newItems)
  }
  const handleInquiryAmountChange = (i: number, value: string) => {
    const numeric = Number(value)
    if (!Number.isFinite(numeric)) {
      updateInqVendor(i, 'amount', 0)
      return
    }
    const capped = totalRevenueIncl > 0 ? Math.min(numeric, totalRevenueIncl) : numeric
    updateInqVendor(i, 'amount', Number(capped.toFixed(2)))
  }
  const removeInqVendor = (i: number) => setInqVendors(inqVendors.filter((_, idx) => idx !== i))

  const autoGenerateInquiry = () => {
    const it = projectData.cost?.it || {}
    const limit = (it.device?.incl||0) + (it.construction?.incl||0) + (it.survey?.incl||0) +
                  (it.integration?.incl||0) + (it.other?.incl||0) + (it.maintenance?.incl||0) +
                  (it.running?.incl||0) + (it.bidding?.incl||0) + (it.design_eval?.incl||0) + (it.audit?.incl||0)

    if (limit === 0) {
      alert("请先完善 IT 投入明细（当前 IT 总成本为 0），三家报价的底价需硬性绑定成本。")
      return
    }
    if (totalRevenueIncl <= 0) {
      alert("请先完善收入侧含税总收入，三家询价最高价不能超过含税总收入。")
      return
    }
    if (limit > totalRevenueIncl) {
      alert(`当前 IT 投入含税总成本为 ${limit.toFixed(2)}，已超过含税总收入 ${totalRevenueIncl.toFixed(2)}，无法生成合规三家报价。`)
      return
    }

    const quotes = [
      limit,
      Math.min(totalRevenueIncl, Math.round(limit * (1.05 + Math.random() * 0.02))),
      Math.min(totalRevenueIncl, Math.round(limit * (1.10 + Math.random() * 0.05)))
    ].sort((a, b) => a - b)

    const shuffled = [0, 1, 2].sort(() => Math.random() - 0.5)
    const generatedVendors = shuffled.map((idx, i) => ({
      vendorName: `厂商${String.fromCharCode(65 + i)}`,
      amount: quotes[idx], taxRate: 6, remark: idx === 0 ? '最低' : '',
      images: []
    }))
    setInqVendors(previous => mergeVendorImages(generatedVendors, previous))
  }

  const handleImageUpload = (e: any, setImages: any, typeName: string) => {
    let filesList: any[] = []

    if (e.clipboardData && e.clipboardData.items) {
      const items = Array.from(e.clipboardData.items) as any[];
      const imgItems = items.filter(it => it.type && it.type.startsWith('image/'));
      if (imgItems.length > 0) {
        e.preventDefault();
        imgItems.forEach(item => {
          const blob = item.getAsFile();
          if (blob) {
            filesList.push(blob);
          }
        });
      }
    } else if (e.dataTransfer) {
      filesList = Array.from(e.dataTransfer.files);
    } else if (e.target && e.target.files) {
      filesList = Array.from(e.target.files);
    }

    if (filesList.length === 0) return;

    const imageFiles = filesList.filter((file: any) => file.type && file.type.indexOf('image/') === 0);
    if (imageFiles.length === 0) return;

    imageFiles.forEach((file: any) => {
      const reader = new FileReader()
      reader.onload = (event) => {
        const img = new Image()
        img.onload = async () => {
          const base64Data = event.target?.result as string;
          const w = img.width;
          const h = img.height;
          
          if (projectId) {
            try {
              const assetId = await domainSaveService.saveTemplateAsset(projectId, selectedTemplate, {
                assetType: "image",
                usage: typeName,
                originalFileName: file.name || "pasted_image",
                base64Data,
                width: w,
                height: h,
              });
              setImages((prev: any) => [...prev, {
                assetId,
                data: base64Data,
                width: w,
                height: h
              }]);
            } catch (err) {
              alert("上传图片失败: " + err);
            }
          } else {
            alert("请选择项目后再上传图片");
          }
        }
        img.src = event.target?.result as string
      }
      reader.readAsDataURL(file)
    });
  };

  const handleRemoveImage = async (img: any, index: number, setImages: any) => {
    setImages((prev: any) => prev.filter((item: any, idx: number) => {
      if (img.assetId && item.assetId) {
        return item.assetId !== img.assetId;
      }
      return idx !== index;
    }));
    if (img.assetId) {
      try {
        await projectService.deleteTemplateAsset(img.assetId);
      } catch (err) {
        console.warn("Soft delete of template asset failed:", err);
      }
    }
  };

  const handleSendImageToAi = async (img: any, fieldKey: string) => {
    if (!projectId || !selectedTemplate || !img?.assetId) return;
    await publishTemplateAssetSelection(createTemplateAssetSelection({
      projectId,
      templateId: selectedTemplate,
      assetId: img.assetId,
      fieldKey,
      fileName: img.fileName || img.originalFileName || fieldKey,
      mimeType: img.mimeType || null,
      size: img.fileSize || null,
      width: img.width || null,
      height: img.height || null,
    }));
  };

  const handleGenerate = async () => {

    if (!formRef.current) return
    const fd = new FormData(formRef.current)
    const get = (name: string) => fd.get(name)?.toString() || ""

    if (selectedTemplate.includes('会审')) {
      const activeQuotes = inqVendors
        .filter(v => v.vendorName || Number(v.amount || 0) > 0)
        .map(v => Number(v.amount || 0))
        .filter(value => value > 0)
      const maxQuote = activeQuotes.length > 0 ? Math.max(...activeQuotes) : 0
      if (maxQuote > 0 && totalRevenueIncl <= 0) {
        alert("请先完善收入侧含税总收入，三家询价最高价不能超过含税总收入。")
        return
      }
      if (totalRevenueIncl > 0 && maxQuote > totalRevenueIncl + 0.01) {
        alert(`三家询价最高价 ${maxQuote.toFixed(2)} 不能超过含税总收入 ${totalRevenueIncl.toFixed(2)}。`)
        return
      }
    }

    const formatDateStr = (dateStr: string) => {
      if (!dateStr) return ""
      const d = new Date(dateStr)
      if (isNaN(d.getTime())) return dateStr
      return `${d.getFullYear()}年${String(d.getMonth()+1).padStart(2, '0')}月${String(d.getDate()).padStart(2, '0')}日`
    }

    // Attendees logic
    let attendees = ""
    const selectedProjectScale = normalizeProjectScale(get('gen_project_scale') || projectScale)
    if (selectedProjectScale === 'large') {
      attendees += `市公司政企部（解决方案、交付支撑、计划部）：\n        ${get('gen_city_attendees')}\n`
    }
    const branchName = get('gen_branch_name') || "XXXX"
    attendees += `${branchName}分公司（建设、维护、网络/信息安全员）：\n        ${get('gen_branch_attendees')}`
    const customItContent = joinedBusinessNames(customItBusinessNames)
    const customCtContent = joinedBusinessNames(customCtBusinessNames)
    const projectBusinessComposition = joinedBusinessNames([
      ...customItBusinessNames,
      ...customCtBusinessNames,
      ...customNonItCtBusinessNames,
      ...customMixBusinessNames,
    ])
    const originalCtContent = customCtContent || (hasMidThree ? (ctContent ? ctContent.replace(/能力/g, '') : "详见清单") : "无")

    const itCostInclForContent = (projectData.cost?.it?.integration?.incl || 0) + (projectData.cost?.it?.device?.incl || 0) + (projectData.cost?.it?.maintenance?.incl || 0)
    const originalItContent = customItContent || (itCostInclForContent > 0 ? (itContent || "集成服务") : "无")

    const signItContent = get('gen_sign_it_content')
    const signCtContent = get('gen_sign_ct_content')

    const itContentStr = selectedTemplate.includes('立项签批表')
      ? (signItContent.trim() !== "" ? signItContent : originalItContent)
      : originalItContent

    const ctContentStr = selectedTemplate.includes('立项签批表')
      ? (signCtContent.trim() !== "" ? signCtContent : originalCtContent)
      : originalCtContent

    const resolvedSubjectItCost = get('gen_subject_it_cost').trim() || joinedBusinessNames(customItCostBusinessNames) || subjectItCost
    const resolvedSubjectCtCost = get('gen_subject_ct_cost').trim() || joinedBusinessNames(customCtCostBusinessNames) || subjectCtCost
    const resolvedSubjectItRev = get('gen_subject_it_rev').trim() || joinedBusinessNames(customItRevenueBusinessNames) || subjectItRev
    const resolvedSubjectCtRev = get('gen_subject_ct_rev').trim() || joinedBusinessNames(customCtRevenueBusinessNames) || subjectCtRev

    const otherCost = projectData.cost?.ct?.other?.incl || 0
    const otherProductContent = getSubjectBusinessName("cost_ct_other") || (otherCost > 0 ? "详见清单" : "无")

    let hasAnyScreenshot = false;
    const screenshotListArray: any[] = [];
    const textScreenshotList: string[] = [];

    inqVendors.filter(v => v.vendorName).forEach(v => {
      const label = `${v.vendorName}`;
      textScreenshotList.push(label);

      const vendorImgs = v.images || [];
      if (vendorImgs.length > 0) {
        hasAnyScreenshot = true;
        vendorImgs.forEach((img: any) => {
          screenshotListArray.push(toDocImagePayload(img, v.vendorName));
        });
      }
    });

    const vendorScreenshotList = hasAnyScreenshot
      ? JSON.stringify(screenshotListArray)
      : textScreenshotList.join('\n\n');

    // 1) 修复 TABLE_TECH_ITEMS
    const techItemsSafe = techItems && techItems.length > 0 ? techItems : [
      { serviceName: '视频监控服务', serviceDesc: '包含摄像头、网关等设备采购及安装', amount: 1, unit: '套' },
      { serviceName: '信息化改造及集成服务', serviceDesc: '包含专线及网络集成', amount: 1, unit: '套' }
    ]
    const techRowsForDocx = techItemsSafe.map((it, i) => ({
      TECH_ITEM_INDEX: String(i + 1),
      TECH_ITEM_NAME: String(it.serviceName || ""),
      TECH_ITEM_DESC: String(it.serviceDesc || ""),
      TECH_ITEM_QTY: String(it.amount ?? ""),
      TECH_ITEM_UNIT: String(it.unit || ""),
    }))

    // 2) 修复 TABLE_INQ_VENDORS
    const inqVendorsSafe = inqVendors && inqVendors.length > 0 ? inqVendors : [
      { vendorName: '厂商A', amount: 0, taxRate: 6, remark: '最低' }
    ]
    const vendorRowsForDocx = inqVendorsSafe.map((v, i) => ({
      INQ_VENDOR_INDEX: String(i + 1),
      INQ_VENDOR_NAME: String(v.vendorName || ""),
      INQ_QUOTE: String(v.amount ?? ""),
      INQ_TAX_RATE: String(v.taxRate ?? ""),
      INQ_REMARK: String(v.remark || ""),
    }))

    // Generate Inquiry Process text dynamically
    const itInquiryProcess = "";

    // Calculate Project Total Investment (Legacy Detailed format for Meeting Minutes)
    const getExclIt = (key: string) => Number(projectData.cost?.it?.[key]?.excl || 0);
    const getExclCt = (key: string) => Number(projectData.cost?.ct?.[key]?.excl || 0);
    const getExclMix = (key: string) => Number(projectData.cost?.mix?.[key]?.excl || 0);

    const itCost = getExclIt('device') + getExclIt('construction') + getExclIt('survey') + getExclIt('integration') + getExclIt('other') + getExclIt('maintenance') + getExclIt('running') + getExclIt('bidding') + getExclIt('design_eval') + getExclIt('audit');
    const ctCost = getExclCt('construction') + getExclCt('maintenance') + getExclCt('other') + getExclCt('bandwidth') + getExclCt('renewal');
    const nonItCost = getExclMix('non_it_ct');
    const mixCost = getExclMix('marketing') + getExclMix('channel') + getExclMix('other');
    const totalCost = itCost + ctCost + nonItCost + mixCost;

    const isZero = (n: number) => Math.abs(n) < 0.005;
    const fmtYuan = (n: number) => n.toFixed(2);
    const fmtPct = (x: any) => isFinite(x) && x !== null && x !== "" && !isNaN(Number(x)) ? (Number(x) * 100).toFixed(2) + '%' : '--';
    const subjectDetailName = (subject: (typeof ICT_SUBJECT_DEFINITIONS)[number], item: any) => {
      const resolved = resolveBillingSubjectPresentation(subject, item);
      const baseName = resolved.billingSubjectName || resolved.standardName;
      const prefix = `${subject.documentPrefix}-`;
      return baseName.startsWith(prefix) ? baseName : `${prefix}${baseName}`;
    };
    const buildSubjectAmountDetails = (side: IctSubjectSide, documentPrefix: IctDocumentPrefix, actionLabel: "投入" | "收入") => {
      return ICT_SUBJECT_DEFINITIONS
        .filter(subject => subject.side === side && subject.documentPrefix === documentPrefix)
        .map(subject => {
          const item = getProjectDataSubjectItem(projectData, subject);
          const amount = Number(item?.excl || 0);
          if (isZero(amount)) return "";
          return `${subjectDetailName(subject, item)}${actionLabel}${fmtYuan(amount)}元`;
        })
        .filter(Boolean);
    };
    const joinSubjectGroups = (groups: string[][]) => groups.filter(group => group.length > 0).map(group => group.join("，")).join("；");
    const afterApprovalSelectionPhrase = get('gen_after_approval_selection') === "on" ? "申请立项后甄选，" : "";
    const investmentDetailGroups = joinSubjectGroups([
      buildSubjectAmountDetails("cost", "IT", "投入"),
      buildSubjectAmountDetails("cost", "CT", "投入"),
      buildSubjectAmountDetails("cost", "非IT/CT", "投入"),
      buildSubjectAmountDetails("cost", "综合类", "投入"),
    ]);
    const projectInvestmentSituation = `总投入${fmtYuan(totalCost)}元${investmentDetailGroups ? `；其中${investmentDetailGroups}` : ""}。此IT部分费用参考三家询价最低价，${afterApprovalSelectionPhrase}最终费用不超过上述总投入。`;
    const projTotalInvestStr = projectInvestmentSituation;

    // Calculate Demand Table specific fields
    const totalRevIt = Object.values(projectData.revenue?.it || {}).reduce((acc: number, curr: any) => acc + (curr?.incl || 0), 0);
    const totalRevCt = Object.values(projectData.revenue?.ct || {}).reduce((acc: number, curr: any) => acc + (curr?.incl || 0), 0);
    const totalRevNonItCt = projectData.revenue?.non_it_ct?.incl || 0;
    const totalRevIncl = Number(totalRevIt) + Number(totalRevCt) + Number(totalRevNonItCt);

    const totalRevItExcl = Object.values(projectData.revenue?.it || {}).reduce((acc: number, curr: any) => acc + (curr?.excl || 0), 0);
    const totalRevCtExcl = Object.values(projectData.revenue?.ct || {}).reduce((acc: number, curr: any) => acc + (curr?.excl || 0), 0);
    const totalRevNonItCtExcl = projectData.revenue?.non_it_ct?.excl || 0;
    const totalRevExcl = Number(totalRevItExcl) + Number(totalRevCtExcl) + Number(totalRevNonItCtExcl);
    const revenueDetailGroups = joinSubjectGroups([
      buildSubjectAmountDetails("revenue", "IT", "收入"),
      buildSubjectAmountDetails("revenue", "CT", "收入"),
      buildSubjectAmountDetails("revenue", "非IT/CT", "收入"),
      buildSubjectAmountDetails("revenue", "综合类", "收入"),
    ]);
    const projectRevenueSituation = `总收入${fmtYuan(totalRevExcl)}元${revenueDetailGroups ? `；其中${revenueDetailGroups}` : ""}。`;

    const branchNameFinal = get('gen_demand_branch_name') || get('gen_branch_name') || "XXX分公司";

    let demandServiceContent = get('gen_demand_service_content');
    if (!demandServiceContent) {
      const parts = [itContentStr, ctContentStr].filter(x => x && x !== "无");
      demandServiceContent = parts.length > 0 ? parts.join("；") : "无";
    }

    const demandCustomerConfirm = get('gen_demand_customer_confirm') || "微信截图";
    const demandDeviceList = get('gen_demand_device_list') || "不涉及";
    const demandEnvRequire = get('gen_demand_env_require') || "客户提供部署环境，不包含在本次项目范围内";

    const demandUrlStr = get('gen_demand_public_url') || "";
    const demandPublicUrlLine = hasPublicUrl ? `\n7、项目有效的公示网址及招标文件：${demandUrlStr}` : "";

    const securityIdx = hasPublicUrl ? 8 : 7;
    const securityDetailStr = get('gen_demand_security_detail');
    const demandSecurityLine = `\n${securityIdx}、信息安全、密评：${hasSecurity ? (securityDetailStr || "有") : "无"}`;

    const attach2TitleLine = hasPublicUrl ? `\n附件2、有效的挂网链接截图/招标文件（有效地址：${demandUrlStr}）` : "";
    const attach1ImageStr = attach1Images.length > 0 ? JSON.stringify(attach1Images.map(img => toDocImagePayload(img))) : "";
    const attach2ImageStr = (hasPublicUrl && attach2Images.length > 0) ? JSON.stringify(attach2Images.map(img => toDocImagePayload(img))) : "";

    const now = new Date();
    const currDate = `${now.getFullYear()}年${String(now.getMonth()+1).padStart(2, '0')}月${String(now.getDate()).padStart(2, '0')}日`;

    const leaderLine = totalRevIncl >= 3000000 ? "分管领导（签字）：________________" : "";

    const variables: any = {
      'PROJECT_NAME': projectData.basic?.proj_name || "",
      'CUSTOMER_NAME': projectData.basic?.customer_name || "",
      'PROJECT_YEARS': String(projectData.basic?.project_years || 1),

      'MEETING_START_DATE': formatDateStr(get('gen_meet_start')),
      'MEETING_END_DATE': formatDateStr(get('gen_meet_end')),
      'MEETING_MODE': get('gen_meet_mode'),
      'ATTENDEES': attendees,
      'ONSITE_SUPPORT': get('gen_onsite_support'),
      'PROJECT_BACKGROUND': projectBackground,
      'PROJECT_BUSINESS_COMPOSITION': projectBusinessComposition,
      'IT_CONTENT': itContentStr,
      'CT_CONTENT': ctContentStr,
      'OTHER_PRODUCT_CONTENT': otherProductContent,
      'TECH_SOLUTION': get('gen_tech_solution'),
      'SELF_THREE_Q': get('gen_self_three'),
      'MID_THREE_Q': hasMidThree ? `本项目涉及中台能力编号：${midThreeCode}，能力名称：${midThreeName}。` : "不涉及",
      'THREEIZATION_PLAN': get('gen_threeization'),
      'TECH_CONCLUSION': get('gen_tech_conclusion'),
      'STRATEGIC_VALUE': get('gen_strategic_value'),
      'IT_BUSINESS_MODE': selectedTemplate.includes("需求导入表") ? (get('gen_demand_it_business_mode') || "服务模式") : (get('gen_it_bus_mode') || itBusMode || "服务购销"),
      'IT_FUNDING_SOURCE': get('gen_it_fund_src') || itFundSrc || "分公司成本开支",
      'IS_JOINT_BIDDING': get('gen_is_joint'),
      'REVENUE_COLLECTION_METHOD': get('gen_rev_collection') || revCollection,
      'EXPENDITURE_PAYMENT_METHOD': get('gen_exp_payment') || expPayment,
      'REV_COLLECTION': get('gen_rev_collection') || revCollection,
      'EXP_PAYMENT': get('gen_exp_payment') || expPayment,
      'PROJECT_REVIEW_ACCURACY': get('gen_review_acc'),
      'SINGLE_SOURCE_EXPLANATION': hasSingleSource ? get('gen_single_source') : "",
      'IS_SME': "是",
      'IS_ADVANCE_PAYMENT': get('gen_is_advance') === "on" ? "是" : "否",

      'SUBJECT_IT_COST': resolvedSubjectItCost,
      'SUBJECT_CT_COST': resolvedSubjectCtCost,
      'SUBJECT_IT_REV': resolvedSubjectItRev,
      'SUBJECT_CT_REV': resolvedSubjectCtRev,
      'CONSTRUCTION_TIME_REQ': get('gen_construction_time_req'),
      'PROCUREMENT_METHOD': procurementMethod === '其他' ? get('gen_procurement_method_other') : procurementMethod,
      'CONSTRUCTION_INTERFACE': get('gen_construction_interface'),
      'RISK_OWNER': get('gen_risk_owner'),

      'IT_INQUIRY_PROCESS': itInquiryProcess,
      'PROJECT_INVESTMENT_SITUATION': projectInvestmentSituation,
      'PROJECT_REVENUE_SITUATION': projectRevenueSituation,
      'PROJECT_TOTAL_INVESTMENT_DETAIL': projTotalInvestStr,
      'PROJECT_TOTAL_INVESTMENT': selectedTemplate.includes("会审") ? projTotalInvestStr : totalCost.toFixed(2),
      'IT_INVESTMENT': itCost.toFixed(2),
      'CT_INVESTMENT': ctCost.toFixed(2),
      'PROJECT_TOTAL_REVENUE': totalRevExcl.toFixed(2),
      'IT_REVENUE': totalRevItExcl.toFixed(2),
      'CT_REVENUE': totalRevCtExcl.toFixed(2),
      'DYNAMIC_PAYBACK_PERIOD': String(metrics?.dynamic_payback || "--"),
      'IT_NET_PRESENT_VALUE_RATE': fmtPct(metrics?.it_npv_rate),
      'NET_PRESENT_VALUE_RATE': fmtPct(metrics?.npv_rate),
      'PROJECT_GROSS_PROFIT_MARGIN': fmtPct(metrics?.margin_rate),

      'TABLE_TECH_ITEMS': JSON.stringify(techRowsForDocx),
      'TABLE_INQ_VENDORS': JSON.stringify(vendorRowsForDocx),
      'VENDOR_SCREENSHOT_LIST': vendorScreenshotList,

      // Demand specific variables
      'BRANCH_NAME': branchNameFinal,
      'CURR_DATE': currDate,
      'DEMAND_BUDGET': String(totalRevIncl),
      'DEMAND_CUSTOMER_CONFIRM': demandCustomerConfirm,
      'DEMAND_DEVICE_LIST': demandDeviceList,
      'DEMAND_ENV_REQUIRE': demandEnvRequire,
      'DEMAND_PUBLIC_URL_LINE': demandPublicUrlLine,
      'DEMAND_SECURITY_LINE': demandSecurityLine,
      'DEMAND_SERVICE_CONTENT': demandServiceContent,
      'LEADER_LINE': leaderLine,
      'ATTACH1_IMAGE': attach1ImageStr,
      'ATTACH2_IMAGE': attach2ImageStr,
      'ATTACH2_TITLE_LINE': attach2TitleLine,

      // --- Excel Specific Variable Back-filling ---
      ...excelSubjectVariables,
      ...(() => {
        const cfVars: Record<string, string> = {}
        const cashflows = Array.isArray(metrics?.cashflow)
          ? metrics.cashflow
          : Array.isArray(metrics?.cashflows)
            ? metrics.cashflows
            : []
        for (let i = 0; i < 10; i++) {
          const row = cashflows[i]
          if (!row) continue
          cfVars[`CASH_IN_Y${i + 1}`] = String(row.cash_in ?? "0")
          cfVars[`CASH_OUT_Y${i + 1}`] = String(row.cash_out ?? "0")
          cfVars[`IT_CASH_IN_Y${i + 1}`] = String(row.it_cash_in ?? "0")
          cfVars[`IT_CASH_OUT_Y${i + 1}`] = String(row.it_cash_out ?? "0")
        }
        return cfVars
      })(),

      'RENEWAL_PROJECT_FLAG': (projectData.cost?.ct?.renewal?.excl ?? 0) > 0 ? "是" : "否",
      'CONTRACT_DURATION': String(projectData.basic?.project_years || 1),
    }

    const runGenerate = (overwriteExisting = false) => invoke<string>('generate_lifecycle_docs', {
          moduleId: "ict_lifecycle",
          variables: variables,
          selectedTemplates: [selectedTemplate],
          outputDir,
          projectId,
          overwriteExisting
      })

    try {
      let generatedOutputDir: string
      try {
        generatedOutputDir = await runGenerate(false)
      } catch (err) {
        const message = String(err)
        if (!message.startsWith("FILE_EXISTS::")) {
          throw err
        }
        const conflictPath = message.replace("FILE_EXISTS::", "")
        const shouldOverwrite = confirm(`目标文件已存在，是否覆盖？\n${conflictPath}`)
        if (!shouldOverwrite) return
        generatedOutputDir = await runGenerate(true)
      }
      if (projectId && outputDir) {
        try {
          await projectFileService.scanProjectFolder(projectId, false)
          onGenerated?.()
        } catch (scanErr) {
          console.warn("生成后同步项目文件失败", scanErr)
        }
      }
      if (confirm(`生成成功！文件已保存至：\n${generatedOutputDir}\n是否立即打开输出目录？`)) {
        invoke('open_file', { path: generatedOutputDir })
      }
    } catch(e) {
      alert("生成失败：" + e)
    }
  }
  const itCostInclForContent = (projectData.cost?.it?.integration?.incl || 0) + (projectData.cost?.it?.device?.incl || 0) + (projectData.cost?.it?.maintenance?.incl || 0)
  const defaultSignItContent = joinedBusinessNames(customItBusinessNames) || (itCostInclForContent > 0 ? (itContent || "集成服务") : "无")
  const defaultSignCtContent = joinedBusinessNames(customCtBusinessNames) || (hasMidThree ? (ctContent ? ctContent.replace(/能力/g, '') : "详见清单") : "无")

  return (
    <div className="flex flex-col gap-6">
      <form ref={formRef} className="flex flex-col gap-6" onSubmit={(e) => e.preventDefault()} onChange={handleFormChange}>

        {/* Excel 预算表/评估表专属配置 */}
        {selectedTemplate.endsWith('.xlsx') && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm flex flex-col gap-4">
            <h4 className="font-bold text-primary flex items-center gap-2">
              <AppIcon name="spreadsheet" size={18} /> 《项目经济效益评估表》Excel 测算模版配置
            </h4>
            <p className="text-sm text-secondary-foreground leading-relaxed">
              您选择的是 Excel 预算/评估模板。系统将根据您在 **收入侧测算** 和 **支出侧测算** 中填写的精细化数据，自动将 **29 项输入指标、项目背景、产权归属、项目周期等基础信息** 全自动回填到 Excel 的对应 sheet 单元格中。
            </p>
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分商务模式</label>
                <select name="gen_it_bus_mode" value={itBusMode} onChange={e => setItBusMode(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm">
                  <option value="服务模式">服务模式</option>
                  <option value="集成购销">集成购销</option>
                  <option value="投资">投资</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分资金来源</label>
                <select name="gen_it_fund_src" value={itFundSrc} onChange={e => setItFundSrc(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm">
                  <option value="分公司成本开支">分公司成本开支</option>
                  <option value="市公司专项资源">市公司专项资源</option>
                </select>
              </div>
            </div>
            <div className="bg-muted p-4 rounded-lg border border-border mt-2">
              <h5 className="font-semibold text-xs text-foreground uppercase tracking-wider mb-2">自动回填映射规则一览</h5>
              <ul className="text-xs text-secondary-foreground space-y-1.5 list-disc list-inside">
                <li><strong className="text-foreground">Sheet 3-直接经济效益评估表：</strong> 自动回填 9 项不含税/含税收入指标，以及 20 项不含税/含税支出指标（包括设备、施工、运行、营销、渠道等）。</li>
                <li><strong className="text-foreground">Sheet 2-ICT项目评估结果：</strong> 自动同步回填项目名称、客户名称、项目周期、续签标志、业务模式、IT 资金来源。</li>
              </ul>
            </div>
          </div>
        )}

        {/* 会审纪要 专属字段 (排除立项签批表、立项决策、Excel 效益分析表和需求导入表) */}
        {selectedTemplate && selectedTemplate.includes('会审') && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
            <h4 className="font-bold text-primary mb-4">补充文档信息 (会审纪要等所需)</h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">会审开始日期</label>
                <input type="date" name="gen_meet_start" {...getBind("gen_meet_start", todayStr)} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">会审结束日期</label>
                <input type="date" name="gen_meet_end" {...getBind("gen_meet_end", todayStr)} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">会审方式</label>
                <select name="gen_meet_mode" {...getBind("gen_meet_mode", "线上")} className="bg-card border border-input px-3 py-2 rounded-md">
                  <option value="线上">线上</option>
                  <option value="线下">线下</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">项目规模</label>
                <select
                  name="gen_project_scale"
                  value={projectScale}
                  onChange={e => handleProjectScaleChange(normalizeProjectScale(e.target.value))}
                  className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm"
                >
                  <option value="large">大项目 (市/省)</option>
                  <option value="small">小项目 (分公司)</option>
                </select>
              </div>
              {projectScale === 'large' && (
                <div className="flex flex-col gap-1 col-span-2">
                  <label className="text-sm font-semibold">市公司政企部参会人员</label>
                  <input type="text" name="gen_city_attendees" {...getBind("gen_city_attendees")} placeholder="人员A、人员B" className="bg-card border border-input px-3 py-2 rounded-md" />
                </div>
              )}
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">分公司参会人员</label>
                <div className="flex gap-2">
                  <input type="text" name="gen_branch_name" {...getBind("gen_branch_name", "XXXX")} placeholder="分公司名称" className="w-32 bg-card border border-input px-3 py-2 rounded-md" />
                  <input type="text" name="gen_branch_attendees" {...getBind("gen_branch_attendees")} placeholder="人员D、人员E" className="flex-1 bg-card border border-input px-3 py-2 rounded-md" />
                </div>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">驻点支撑人员</label>
                <input type="text" name="gen_onsite_support" {...getBind("gen_onsite_support")} placeholder="如有请填写" className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">项目背景</label>
                <textarea name="gen_proj_bg" rows={3} value={projectBackground} onChange={e => setProjectBackground(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold flex items-center justify-between">IT建设内容 <span className="text-xs text-secondary-foreground font-normal">根据项目名称自动生成</span></label>
                <textarea name="gen_it_content" value={itContent} onChange={e => setItContent(e.target.value)} rows={2} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold flex items-center justify-between">CT建设内容 <span className="text-xs text-secondary-foreground font-normal">中台能力联动修改</span></label>
                <textarea name="gen_ct_content" value={ctContent} onChange={e => setCtContent(e.target.value)} rows={2} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">技术方案</label>
                <textarea name="gen_tech_solution" rows={2} {...getBind("gen_tech_solution", "采用端-管-云架构...")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>

              <div className="col-span-2 border border-border rounded-lg p-4 bg-background">
                <div className="flex justify-between items-center mb-3">
                  <label className="text-sm font-bold">技术方案可行性清单</label>
                  <button type="button" onClick={addTechItem} className="text-xs bg-primary-soft text-primary px-3 py-1.5 rounded font-semibold hover:bg-primary-soft/80">+ 新增一行</button>
                </div>
                <div className="flex flex-col gap-2">
                  {techItems.map((item, i) => (
                    <div key={i} className="flex gap-2 items-center">
                      <input type="text" placeholder="服务名称" value={item.serviceName} onChange={e => updateTechItem(i, 'serviceName', e.target.value)} className="w-1/4 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                      <input type="text" placeholder="服务说明" value={item.serviceDesc} onChange={e => updateTechItem(i, 'serviceDesc', e.target.value)} className="flex-1 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                      <input type="number" placeholder="数量" value={item.amount} onChange={e => updateTechItem(i, 'amount', Number(e.target.value))} className="w-16 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                      <input type="text" placeholder="单位" value={item.unit} onChange={e => updateTechItem(i, 'unit', e.target.value)} className="w-16 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                      <button type="button" onClick={() => removeTechItem(i)} className="text-destructive hover:bg-destructive/10 p-1.5 rounded" title="删除">
                        <AppIcon name="delete" size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">自主三问</label>
                <select
                  name="gen_self_three"
                  value={selfThreeValue}
                  onChange={e => handleSelfThreeChange(e.target.value)}
                  className="bg-card border border-input px-3 py-2 rounded-md outline-none text-sm"
                >
                  {SELF_THREE_OPTIONS.map(option => (
                    <option key={option.value} value={option.value}>{option.value}</option>
                  ))}
                </select>
                {(selectedSelfThree.reminder || selfThreeMissingFees.length > 0) && (
                  <div className="mt-1 space-y-1 text-xs leading-5">
                    {selectedSelfThree.reminder && (
                      <div className="rounded-md border border-primary/20 bg-primary/5 px-3 py-2 text-primary">
                        {selectedSelfThree.reminder}
                      </div>
                    )}
                    {selfThreeMissingFees.length > 0 && (
                      <div className="rounded-md border border-warning/20 bg-warning-soft px-3 py-2 text-warning-foreground">
                        {selfThreeMissingFees.join("；")}。
                      </div>
                    )}
                  </div>
                )}
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">三化方案</label>
                <input type="text" name="gen_threeization" {...getBind("gen_threeization", "本项目不涉及三化方案。")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">战略价值</label>
                <input type="text" name="gen_strategic_value" {...getBind("gen_strategic_value")} placeholder="战略价值说明" className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">结论</label>
                <input type="text" name="gen_tech_conclusion" {...getBind("gen_tech_conclusion", "方案可行同时能满足客户需求。")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold flex items-center gap-2">
                  <input type="checkbox" checked={hasMidThree} onChange={e => setHasMidThree(e.target.checked)} className="w-4 h-4" />
                  涉及中台能力调用
                </label>
                {hasMidThree && (
                  <div className="flex flex-col gap-2 mt-1">
                    <div className="flex gap-2">
                      <input
                        type="text"
                        name="gen_mid_three_code"
                        value={midThreeCode}
                        onChange={e => setMidThreeCode(e.target.value)}
                        placeholder="能力编号"
                        className="w-1/3 bg-card border border-input px-3 py-2 rounded-md text-sm"
                      />
                      <input
                        type="text"
                        list="mid-three-capabilities-list"
                        name="gen_mid_three_name"
                        value={midThreeName}
                        onChange={e => {
                          const val = e.target.value;
                          setMidThreeName(val);
                          const matched = MID_THREE_CAPABILITIES.find(c => c.label === val || c.value === val);
                          if (matched) {
                            setMidThreeCode(matched.code);
                          }
                        }}
                        placeholder="请选择或输入所需的中台能力"
                        className="flex-1 bg-card border border-input px-3 py-2 rounded-md text-sm"
                      />
                      <button
                        type="button"
                        onClick={() => setIsMidThreeModalOpen(true)}
                        className="px-3 bg-primary text-primary-foreground rounded-md hover:bg-primary/90 flex items-center justify-center shrink-0 transition-colors"
                        title="全局能力库"
                      >
                        <AppIcon name="tableProperties" size={16} />
                      </button>
                    </div>
                    <datalist id="mid-three-capabilities-list">
                      {MID_THREE_CAPABILITIES.map((cap, idx) => (
                        <option key={`${cap.code}-${idx}`} value={cap.value}>
                          {cap.label} ({cap.code})
                        </option>
                      ))}
                    </datalist>
                  </div>
                )}
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">收入侧收款方式</label>
                <input type="text" name="gen_rev_collection" value={revCollection} onChange={e => setRevCollection(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">支出侧付款方式</label>
                <input type="text" name="gen_exp_payment" value={expPayment} onChange={e => setExpPayment(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分商务模式</label>
                <select name="gen_it_bus_mode" value={itBusMode} onChange={e => setItBusMode(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md">
                  <option value="服务模式">服务模式</option>
                  <option value="集成购销">集成购销</option>
                  <option value="投资">投资</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分资金来源</label>
                <select name="gen_it_fund_src" value={itFundSrc} onChange={e => setItFundSrc(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md">
                  <option value="分公司成本开支">分公司成本开支</option>
                  <option value="市公司专项资源">市公司专项资源</option>
                </select>
              </div>

              <div className="col-span-2 border border-border rounded-lg p-4 bg-background">
                <div className="flex justify-between items-center mb-3">
                  <div>
                    <label className="text-sm font-bold">IT部分询价过程</label>
                    {totalRevenueIncl > 0 && (
                      <div className="mt-1 text-[11px] font-semibold text-secondary-foreground">
                        三家询价最高价不超过含税总收入 {totalRevenueIncl.toFixed(2)}
                      </div>
                    )}
                  </div>
                  <div className="flex gap-2">
                    <button type="button" onClick={autoGenerateInquiry} className="inline-flex items-center gap-1.5 text-xs bg-warning-soft text-warning-foreground px-3 py-1.5 rounded font-bold hover:bg-warning/20">
                      <AppIcon name="quickAction" size={14} /> 一键生成三家报价
                    </button>
                    <button type="button" onClick={addInqVendor} className="text-xs bg-primary-soft text-primary px-3 py-1.5 rounded font-semibold hover:bg-primary-soft/80">+ 新增厂商</button>
                  </div>
                </div>
                <div className="flex flex-col gap-2">
                  {inqVendors.map((item, i) => (
                    <div key={i} className="flex gap-2 items-center">
                      <input type="text" placeholder="厂商名称" value={item.vendorName} onChange={e => updateInqVendor(i, 'vendorName', e.target.value)} className="flex-1 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                      <input type="number" placeholder="含税报价" value={item.amount === 0 ? '' : item.amount} onChange={e => handleInquiryAmountChange(i, e.target.value)} className="w-28 bg-card border border-input px-2 py-1.5 rounded-md text-sm" title={totalRevenueIncl > 0 ? `最高不超过含税总收入 ${totalRevenueIncl.toFixed(2)}` : undefined} />
                      <div className="flex items-center gap-1">
                        <input type="number" placeholder="税率" value={item.taxRate} onChange={e => updateInqVendor(i, 'taxRate', Number(e.target.value))} className="w-16 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                        <span className="text-xs text-secondary-foreground">%</span>
                      </div>
                      <input type="text" placeholder="备注" value={item.remark} onChange={e => updateInqVendor(i, 'remark', e.target.value)} className="w-20 bg-card border border-input px-2 py-1.5 rounded-md text-sm" />
                      <button type="button" onClick={() => removeInqVendor(i)} className="text-destructive hover:bg-destructive/10 p-1.5 rounded" title="删除">
                        <AppIcon name="delete" size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {inqVendors.some(v => v.vendorName) && (
                <div className="col-span-2 border border-border rounded-lg p-5 bg-background">
                  <label className="text-sm font-bold text-foreground block mb-3">询价厂商报价截图上传</label>
                  <div className="flex flex-col gap-4">
                    {inqVendors.map((v, i) => {
                      if (!v.vendorName) return null;
                      const vendorImages = v.images || [];

                      const setVendorImages = (updater: any) => {
                        setInqVendors(previous => previous.map((vendor, index) => {
                          if (index !== i) return vendor;
                          const prevImages = vendor.images || [];
                          const newImages = typeof updater === 'function' ? updater(prevImages) : updater;
                          return { ...vendor, images: newImages };
                        }));
                      };

                      return (
                        <div key={i} className="border border-border/60 bg-muted/5 p-4 rounded-xl flex flex-col gap-3 transition-colors hover:border-primary/40">
                          <div className="flex justify-between items-center">
                            <span className="text-sm font-bold text-foreground flex items-center gap-1.5">
                              <span className="bg-primary-soft text-primary text-xs w-5 h-5 rounded-full flex items-center justify-center font-bold">{i + 1}</span>
                              <span className="text-primary font-extrabold">{v.vendorName}</span> 报价截图
                            </span>
                            {vendorImages.length > 0 && (
                              <span className="text-xs bg-success-soft text-success px-2 py-0.5 rounded-full font-medium">
                                已上传 {vendorImages.length} 张图片
                              </span>
                            )}
                          </div>

                          <input
                            type="file"
                            multiple
                            accept="image/*"
                            className="hidden"
                            id={`vendor-file-input-${i}`}
                            onChange={(e) => handleImageUpload(e, setVendorImages, "vendor_" + i)}
                          />

                          <div
                            className="border border-dashed border-border rounded-lg p-5 text-center cursor-pointer hover:bg-muted/50 focus:border-ring focus:ring-2 focus:ring-ring/20 outline-none transition-all flex flex-col items-center justify-center gap-2 bg-muted/20"
                            onClick={(e) => e.currentTarget.focus()}
                            onDragOver={e => e.preventDefault()}
                            onDrop={e => { e.preventDefault(); handleImageUpload(e, setVendorImages, "vendor_" + i); }}
                            onPaste={e => handleImageUpload(e, setVendorImages, "vendor_" + i)}
                            tabIndex={0}
                          >
                            <p className="text-xs text-secondary-foreground">
                              点击聚焦后直接按下 <kbd className="bg-background px-1 border rounded text-[10px] font-mono font-bold">Ctrl+V</kbd> / <kbd className="bg-background px-1 border rounded text-[10px] font-mono font-bold">Cmd+V</kbd> 粘贴截图
                            </p>
                            <div className="flex items-center gap-2">
                              <span className="text-xs text-secondary-foreground">或拖拽图片到此，或者</span>
                              <button
                                type="button"
                                onClick={(e) => { e.stopPropagation(); document.getElementById(`vendor-file-input-${i}`)?.click(); }}
                                className="inline-flex items-center gap-1.5 text-[11px] bg-primary text-primary-foreground px-2.5 py-1 rounded font-semibold hover:bg-primary/90 transition-colors shadow-sm"
                              >
                                <AppIcon name="imageUpload" size={14} /> 选择本地图片
                              </button>
                            </div>
                          </div>

                          {vendorImages.length > 0 && (
                            <div className="flex gap-2 flex-wrap mt-1">
                              {vendorImages.map((img: any, imgIdx: number) => (
                                <div key={imgIdx} className="relative w-20 h-20 border rounded-lg overflow-hidden group">
                                  <img src={img.data} className="w-full h-full object-cover" />
                                  <button
                                    type="button"
                                    onClick={() => {
                                      handleRemoveImage(img, imgIdx, setVendorImages);
                                    }}
                                    className="absolute top-1 right-1 bg-background/90 text-destructive rounded-full w-5 h-5 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity text-xs shadow-sm"
                                    title="移除图片"
                                  >
                                    <AppIcon name="close" size={12} strokeWidth={2} />
                                  </button>
                                  {img.assetId && (
                                    <button
                                      type="button"
                                      onClick={() => handleSendImageToAi(img, `vendor_${i}`)}
                                      className="absolute bottom-1 left-1 right-1 rounded bg-background/95 px-1 py-0.5 text-[10px] font-semibold text-primary opacity-0 shadow-sm transition-opacity hover:bg-primary hover:text-primary-foreground group-hover:opacity-100"
                                      title="发送给 AI 分析"
                                    >
                                      AI 分析
                                    </button>
                                  )}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">风险点及其他责任人</label>
                <input type="text" name="gen_risk_owner" {...getBind("gen_risk_owner", "人员A")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">是否联合体投标</label>
                <input type="text" name="gen_is_joint" {...getBind("gen_is_joint", "否")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">项目评审表准确完整</label>
                <input type="text" name="gen_review_acc" {...getBind("gen_review_acc", "是，项目投入收入核算完整，各表填写准确")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2 mt-2">
                <label className="text-sm font-bold text-foreground flex items-center gap-2">
                  <input type="checkbox" checked={hasSingleSource} onChange={e => setHasSingleSource(e.target.checked)} className="w-4 h-4" />
                  是否涉及单一来源
                </label>
                {hasSingleSource && (
                  <textarea name="gen_single_source" rows={3} {...getBind("gen_single_source", "单一来源决策依据：符合单一来源场景...")} className="bg-card border border-input px-3 py-2 rounded-md" />
                )}
              </div>

              <div className="flex flex-col gap-1">
                <label className="text-sm font-bold text-foreground">采购方式</label>
                <select value={procurementMethod} onChange={e => setProcurementMethod(e.target.value)} className="bg-card border border-input px-3 py-2 rounded-md">
                  <option value="短名单甄选">短名单甄选</option>
                  <option value="采购">采购</option>
                  <option value="其他">其他</option>
                </select>
                {procurementMethod === '其他' && (
                  <input type="text" name="gen_procurement_method_other" {...getBind("gen_procurement_method_other")} placeholder="请输入其他采购方式" className="bg-card border border-input px-3 py-2 rounded-md mt-1" />
                )}
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">时间要求</label>
                <textarea name="gen_construction_time_req" rows={2} {...getBind("gen_construction_time_req", "合同签定后30天内。")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">售中建设及施工界面</label>
                <textarea name="gen_construction_interface" rows={2} {...getBind("gen_construction_interface", "本项目采购统一集成单位实施。分公司负责客户侧的协调工作，并协调管理合作伙伴完成交付。")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
            </div>
          </div>
        )}

        {/* 立项签批表专属配置 */}
        {selectedTemplate.includes('立项签批表') && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
            <h4 className="font-bold text-primary mb-4">《ICT项目立项签批表》专属配置</h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">项目背景</label>
                <textarea
                  name="gen_proj_bg"
                  rows={3}
                  value={projectBackground}
                  onChange={e => setProjectBackground(e.target.value)}
                  className="bg-card border border-input px-3 py-2 rounded-md"
                  placeholder="请输入项目背景..."
                />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold flex items-center justify-between">
                  <span>IT服务内容</span>
                  <span className="text-xs text-secondary-foreground font-normal">为空则用系统默认</span>
                </label>
                <textarea
                  name="gen_sign_it_content"
                  rows={2}
                  {...getBind("gen_sign_it_content", defaultSignItContent)}
                  className="bg-card border border-input px-3 py-2 rounded-md"
                  placeholder={defaultSignItContent}
                />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold flex items-center justify-between">
                  <span>CT服务内容</span>
                  <span className="text-xs text-secondary-foreground font-normal">为空则用系统默认</span>
                </label>
                <textarea
                  name="gen_sign_ct_content"
                  rows={2}
                  {...getBind("gen_sign_ct_content", defaultSignCtContent)}
                  className="bg-card border border-input px-3 py-2 rounded-md"
                  placeholder={defaultSignCtContent}
                />
              </div>
              <div className="flex items-center gap-6 col-span-2">
                <label className="text-sm font-semibold flex items-center gap-2">
                  <input type="checkbox" name="gen_is_advance" {...getBindCheckbox("gen_is_advance")} className="w-4 h-4" />
                  是否涉及垫资
                </label>
                <label className="text-sm font-semibold flex items-center gap-2">
                  <input type="checkbox" name="gen_after_approval_selection" {...getBindCheckbox("gen_after_approval_selection")} className="w-4 h-4" />
                  是否立项后甄选
                </label>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">收入侧收款方式</label>
                <input
                  type="text"
                  name="gen_rev_collection"
                  value={revCollection}
                  onChange={e => setRevCollection(e.target.value)}
                  className="bg-card border border-input px-3 py-2 rounded-md"
                  placeholder="请输入收入侧收款方式..."
                />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">支出侧付款方式</label>
                <input
                  type="text"
                  name="gen_exp_payment"
                  value={expPayment}
                  onChange={e => setExpPayment(e.target.value)}
                  className="bg-card border border-input px-3 py-2 rounded-md"
                  placeholder="请输入支出侧付款方式..."
                />
              </div>
            </div>
          </div>
        )}

        {/* 需求导入表专属配置 */}
        {selectedTemplate.includes('需求导入表') && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
            <h4 className="font-bold text-primary mb-4">《ICT项目需求导入表》专属配置</h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">项目需求单位</label>
                <input type="text" name="gen_demand_branch_name" {...getBind("gen_demand_branch_name", "XXX分公司")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">业务模式</label>
                <select name="gen_demand_it_business_mode" {...getBind("gen_demand_it_business_mode", "服务模式")} className="bg-card border border-input px-3 py-2 rounded-md">
                  <option value="服务模式">服务模式</option>
                  <option value="投资">投资</option>
                </select>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">服务内容</label>
                <textarea name="gen_demand_service_content" {...getBind("gen_demand_service_content", "IT；CT")} rows={2} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">设备清单</label>
                <input type="text" name="gen_demand_device_list" {...getBind("gen_demand_device_list", "不涉及")} className="bg-card border border-input px-3 py-2 rounded-md" />
              </div>
              <div className="col-span-2 border border-border rounded-lg p-4 bg-background">
                <div className="flex justify-between items-center mb-3">
                  <label className="text-sm font-bold text-foreground">技术方案可行性清单 (设备需求清单)</label>
                  <button type="button" onClick={addTechItem} className="text-xs bg-primary-soft text-primary px-3 py-1.5 rounded font-semibold hover:bg-primary-soft/80">+ 新增一行</button>
                </div>
                <div className="flex flex-col gap-2">
                  {techItems.map((item, i) => (
                    <div key={i} className="flex gap-2 items-center">
                      <input type="text" placeholder="服务名称" value={item.serviceName} onChange={e => updateTechItem(i, 'serviceName', e.target.value)} className="w-1/4 bg-card border border-input px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <input type="text" placeholder="服务说明" value={item.serviceDesc} onChange={e => updateTechItem(i, 'serviceDesc', e.target.value)} className="flex-1 bg-card border border-input px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <input type="number" placeholder="数量" value={item.amount} onChange={e => updateTechItem(i, 'amount', Number(e.target.value))} className="w-16 bg-card border border-input px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <input type="text" placeholder="单位" value={item.unit} onChange={e => updateTechItem(i, 'unit', e.target.value)} className="w-16 bg-card border border-input px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <button type="button" onClick={() => removeTechItem(i)} className="text-destructive hover:bg-destructive/10 p-1.5 rounded" title="删除">
                        <AppIcon name="delete" size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">客户确认</label>
                <input type="text" name="gen_demand_customer_confirm" {...getBind("gen_demand_customer_confirm", "微信截图")} className="bg-card border border-input px-3 py-2 rounded-md text-foreground" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">部署环境要求</label>
                <input type="text" name="gen_demand_env_require" {...getBind("gen_demand_env_require", "客户提供部署环境，不包含在本次项目范围内")} className="bg-card border border-input px-3 py-2 rounded-md text-foreground" />
              </div>

              <div className="flex flex-col gap-1 col-span-2 mt-2">
                <label className="text-sm font-bold text-foreground flex items-center gap-2">
                  <input type="checkbox" checked={hasPublicUrl} onChange={e => {
                    setHasPublicUrl(e.target.checked);
                    handleFieldChange("gen_has_public_url", e.target.checked ? "on" : "off");
                    if (!e.target.checked) {
                      setAttach2Images([]);
                      handleFieldChange("gen_demand_public_url", "");
                    }
                  }} className="w-4 h-4" />
                  项目有效的公示网址及招标文件
                </label>
                {hasPublicUrl && (
                  <input type="text" name="gen_demand_public_url" {...getBind("gen_demand_public_url")} placeholder="https://..." className="bg-card border border-input px-3 py-2 rounded-md text-foreground" />
                )}
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-bold text-foreground flex items-center gap-2">
                  <input type="checkbox" checked={hasSecurity} onChange={e => {
                    setHasSecurity(e.target.checked);
                    handleFieldChange("gen_has_security", e.target.checked ? "on" : "off");
                    if (!e.target.checked) {
                      handleFieldChange("gen_demand_security_detail", "");
                    }
                  }} className="w-4 h-4" />
                  信息安全、密评
                </label>
                {hasSecurity && (
                  <input type="text" name="gen_demand_security_detail" {...getBind("gen_demand_security_detail")} placeholder="例如：已做密评/待补充" className="bg-card border border-input px-3 py-2 rounded-md text-foreground" />
                )}
              </div>


              {/* Image Attachments */}
              <div className="flex flex-col gap-1 col-span-2 mt-2">
                <label className="text-sm font-bold text-foreground">附件1截图（客户确认材料）</label>
                <input type="file" multiple accept="image/*" className="hidden" ref={fileInput1Ref} onChange={(e) => handleImageUpload(e, setAttach1Images, "attach1")} />
                <div
                  className="border-2 border-dashed border-border rounded-lg p-6 text-center cursor-pointer hover:bg-muted/50 focus:border-ring focus:ring-2 focus:ring-ring/20 outline-none transition-all flex flex-col items-center justify-center gap-3 bg-muted/20"
                  onClick={(e) => e.currentTarget.focus()}
                  onDragOver={e => e.preventDefault()}
                  onDrop={e => { e.preventDefault(); handleImageUpload(e, setAttach1Images, "attach1"); }}
                  onPaste={e => handleImageUpload(e, setAttach1Images, "attach1")}
                  tabIndex={0}
                >
                  <p className="text-sm text-secondary-foreground">
                    点击聚焦本区域，然后直接按下 <kbd className="bg-background px-1.5 py-0.5 border rounded text-xs font-mono font-bold">Ctrl+V</kbd> / <kbd className="bg-background px-1.5 py-0.5 border rounded text-xs font-mono font-bold">Cmd+V</kbd> 粘贴截图
                  </p>
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-secondary-foreground">或拖拽图片到此，或者</span>
                    <button
                      type="button"
                      onClick={(e) => { e.stopPropagation(); fileInput1Ref.current?.click(); }}
                      className="inline-flex items-center gap-1.5 text-xs bg-primary text-primary-foreground px-3 py-1.5 rounded-md font-semibold hover:bg-primary/90 transition-colors shadow-sm"
                    >
                      <AppIcon name="imageUpload" size={14} /> 选择本地图片
                    </button>
                  </div>
                </div>
                <div className="flex gap-2 flex-wrap mt-2">
                  {attach1Images.map((img, i) => (
                    <div key={i} className="relative w-24 h-24 border rounded overflow-hidden group">
                      <img src={img.data} className="w-full h-full object-cover" />
                      <button type="button" onClick={() => handleRemoveImage(img, i, setAttach1Images)} className="absolute top-1 right-1 bg-background/90 text-destructive rounded-full w-5 h-5 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity text-xs shadow-sm" title="移除图片">
                        <AppIcon name="close" size={12} strokeWidth={2} />
                      </button>
                      {img.assetId && (
                        <button type="button" onClick={() => handleSendImageToAi(img, "attach1")} className="absolute bottom-1 left-1 right-1 rounded bg-background/95 px-1 py-0.5 text-[10px] font-semibold text-primary opacity-0 shadow-sm transition-opacity hover:bg-primary hover:text-primary-foreground group-hover:opacity-100" title="发送给 AI 分析">
                          AI 分析
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {hasPublicUrl && (
                <div className="flex flex-col gap-1 col-span-2 mt-2">
                  <label className="text-sm font-bold text-foreground">附件2截图（招标文件/挂网截图）</label>
                  <input type="file" multiple accept="image/*" className="hidden" ref={fileInput2Ref} onChange={(e) => handleImageUpload(e, setAttach2Images, "attach2")} />
                  <div
                    className="border-2 border-dashed border-border rounded-lg p-6 text-center cursor-pointer hover:bg-muted/50 focus:border-ring focus:ring-2 focus:ring-ring/20 outline-none transition-all flex flex-col items-center justify-center gap-3 bg-muted/20"
                    onClick={(e) => e.currentTarget.focus()}
                    onDragOver={e => e.preventDefault()}
                    onDrop={e => { e.preventDefault(); handleImageUpload(e, setAttach2Images, "attach2"); }}
                    onPaste={e => handleImageUpload(e, setAttach2Images, "attach2")}
                    tabIndex={0}
                  >
                    <p className="text-sm text-secondary-foreground">
                      点击聚焦本区域，然后直接按下 <kbd className="bg-background px-1.5 py-0.5 border rounded text-xs font-mono font-bold">Ctrl+V</kbd> / <kbd className="bg-background px-1.5 py-0.5 border rounded text-xs font-mono font-bold">Cmd+V</kbd> 粘贴截图
                    </p>
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-secondary-foreground">或拖拽图片到此，或者</span>
                      <button
                        type="button"
                        onClick={(e) => { e.stopPropagation(); fileInput2Ref.current?.click(); }}
                        className="inline-flex items-center gap-1.5 text-xs bg-primary text-primary-foreground px-3 py-1.5 rounded-md font-semibold hover:bg-primary/90 transition-colors shadow-sm"
                      >
                        <AppIcon name="imageUpload" size={14} /> 选择本地图片
                      </button>
                    </div>
                  </div>
                  <div className="flex gap-2 flex-wrap mt-2">
                    {attach2Images.map((img, i) => (
                      <div key={i} className="relative w-24 h-24 border rounded overflow-hidden group">
                        <img src={img.data} className="w-full h-full object-cover" />
                        {img.assetId && (
                          <button type="button" onClick={() => handleSendImageToAi(img, "attach2")} className="absolute bottom-1 left-1 right-1 rounded bg-background/95 px-1 py-0.5 text-[10px] font-semibold text-primary opacity-0 shadow-sm transition-opacity hover:bg-primary hover:text-primary-foreground group-hover:opacity-100" title="发送给 AI 分析">
                            AI 分析
                          </button>
                        )}
                        <button type="button" onClick={() => handleRemoveImage(img, i, setAttach2Images)} className="absolute top-1 right-1 bg-background/90 text-destructive rounded-full w-5 h-5 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity text-xs shadow-sm" title="移除图片">
                          <AppIcon name="close" size={12} strokeWidth={2} />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

      </form>

      <button className="inline-flex items-center gap-2 bg-primary text-primary-foreground font-bold py-3 px-6 rounded-lg self-start shadow-sm hover:opacity-90 transition-opacity" onClick={handleGenerate}>
        <AppIcon name="generate" size={18} /> 立即生成此文件
      </button>

      {isMidThreeModalOpen && (
        <div className="fixed inset-0 z-50 bg-background/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-card rounded-xl shadow-md w-full max-w-3xl max-h-[80vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between p-4 border-b border-border">
              <h3 className="font-bold text-lg flex items-center gap-2">
                <AppIcon name="tableProperties" size={20} className="text-primary" />
                全局能力库
              </h3>
              <button onClick={() => setIsMidThreeModalOpen(false)} className="p-1.5 rounded-full hover:bg-muted text-muted-foreground transition-colors" title="关闭">
                <AppIcon name="close" size={20} />
              </button>
            </div>
            <div className="p-4 border-b border-border bg-muted/30">
              <div className="relative">
                <AppIcon name="search" size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
                <input
                  type="text"
                  value={midThreeSearch}
                  onChange={e => setMidThreeSearch(e.target.value)}
                  placeholder="搜索能力名称、编号或所属类别..."
                  className="w-full pl-9 pr-4 py-2.5 rounded-md border border-input bg-card text-sm focus:outline-none focus:ring-2 focus:ring-ring/20 shadow-sm"
                  autoFocus
                />
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-0">
              <table className="w-full text-sm text-left">
                <thead className="text-xs text-muted-foreground uppercase bg-muted/50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="px-6 py-3.5 font-medium">能力名称</th>
                    <th className="px-6 py-3.5 font-medium w-32">能力编号</th>
                    <th className="px-6 py-3.5 font-medium w-36">所属类别</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border">
                  {MID_THREE_CAPABILITIES.filter(c => c.label.includes(midThreeSearch) || c.value.includes(midThreeSearch) || c.code.includes(midThreeSearch) || c.type.includes(midThreeSearch)).map((cap, idx) => (
                    <tr
                      key={`${cap.code}-${idx}`}
                      className="hover:bg-primary/5 cursor-pointer transition-colors group"
                      onClick={() => {
                        setMidThreeName(cap.value);
                        setMidThreeCode(cap.code);
                        setIsMidThreeModalOpen(false);
                      }}
                    >
                      <td className="px-6 py-4 font-medium text-foreground group-hover:text-primary transition-colors">{cap.label}</td>
                      <td className="px-6 py-4 text-muted-foreground font-mono text-xs">{cap.code}</td>
                      <td className="px-6 py-4 text-muted-foreground">
                        <span className="bg-secondary text-secondary-foreground px-2.5 py-1 rounded-full text-xs font-medium">
                          {cap.type}
                        </span>
                      </td>
                    </tr>
                  ))}
                  {MID_THREE_CAPABILITIES.filter(c => c.label.includes(midThreeSearch) || c.value.includes(midThreeSearch) || c.code.includes(midThreeSearch) || c.type.includes(midThreeSearch)).length === 0 && (
                    <tr>
                      <td colSpan={3} className="px-6 py-12 text-center text-muted-foreground">没有找到匹配的中台能力数据</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
