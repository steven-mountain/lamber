import { useState, useRef, useEffect } from "react"
import { TableProperties, X, Search } from "lucide-react"
import { invoke } from "@tauri-apps/api/core"
import { MID_THREE_CAPABILITIES } from "../lib/midThreeConstants"

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
  metrics
}: Props) {
  const formRef = useRef<HTMLFormElement>(null)

  // Specific state for dynamic toggles
  const [projectScale, setProjectScale] = useState("large")
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
  const [midThreeCode, setMidThreeCode] = useState("A301000041")
  const [midThreeName, setMidThreeName] = useState("视频监控能力")
  const [itBusMode, setItBusMode] = useState("服务购销")
  const [itFundSrc, setItFundSrc] = useState("分公司成本开支")
  const [revCollection, setRevCollection] = useState("项目验收完成后30天内客户单位支付100%")
  const [expPayment, setExpPayment] = useState("项目验收完成且收到款项后30天内支付100%")
  
  const [isMidThreeModalOpen, setIsMidThreeModalOpen] = useState(false)
  const [midThreeSearch, setMidThreeSearch] = useState("")
  
  const [subjectItCost, setSubjectItCost] = useState("IT集成")
  const [subjectCtCost, setSubjectCtCost] = useState("CT-视频监控")
  const [subjectItRev, setSubjectItRev] = useState("小微ICT业务-IoT-集成")
  const [subjectCtRev, setSubjectCtRev] = useState("CT-视频监控")



  const formDataRef = useRef<Record<string, string>>({});
  const handleFormInput = (e: any) => {
    const target = e.target;
    if (target && target.name && target.name.startsWith('gen_')) {
      formDataRef.current[target.name] = target.value;
    }
  };

  useEffect(() => {
    if (formRef.current) {
      Object.entries(formDataRef.current).forEach(([name, value]) => {
        const el = formRef.current?.querySelector(`[name="${name}"]`) as HTMLInputElement | HTMLTextAreaElement;
        if (el && el.value !== value) {
          el.value = value;
        }
      });
    }
  }, [selectedTemplate]);

  // -- Linkage Logic --
  useEffect(() => {
    let name = projectData.basic?.proj_name || ""
    name = name.replace(/项目/g, "")
    if (name && !name.includes("服务")) name += "服务"
    setItContent(name)
  }, [projectData.basic?.proj_name])

  useEffect(() => {
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
    
    const quotes = [
      limit, // 最低价直接等于 IT 投入含税总成本
      Math.round(limit * (1.05 + Math.random() * 0.02)), // 基准价上浮约 5%-7%
      Math.round(limit * (1.10 + Math.random() * 0.05))  // 最高价上浮约 10%-15%
    ].sort((a, b) => a - b)

    const shuffled = [0, 1, 2].sort(() => Math.random() - 0.5)
    setInqVendors(shuffled.map((idx, i) => ({
      vendorName: `厂商${String.fromCharCode(65 + i)}`,
      amount: quotes[idx], taxRate: 6, remark: idx === 0 ? '最低' : '',
      images: []
    })))
  }

  const handleImageUpload = (e: any, setImages: any) => {
    let filesList: any[] = []

    if (e.clipboardData && e.clipboardData.items) {
      // Paste event
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
      // Drop event
      filesList = Array.from(e.dataTransfer.files);
    } else if (e.target && e.target.files) {
      // File input event
      filesList = Array.from(e.target.files);
    }

    if (filesList.length === 0) return;

    const newImages: any[] = []
    let processed = 0
    const imageFiles = filesList.filter((file: any) => file.type && file.type.indexOf('image/') === 0);
    
    if (imageFiles.length === 0) return;

    imageFiles.forEach((file: any) => {
      const reader = new FileReader()
      reader.onload = (event) => {
        const img = new Image()
        img.onload = () => {
          newImages.push({
            data: event.target?.result as string,
            width: img.width,
            height: img.height
          })
          processed++
          if (processed === imageFiles.length) {
            setImages((prev: any) => [...prev, ...newImages])
          }
        }
        img.src = event.target?.result as string
      }
      reader.readAsDataURL(file)
    })
  }

  const removeImage = (index: number, setImages: any) => {
    setImages((prev: any) => prev.filter((_: any, i: number) => i !== index))
  }

  const handleGenerate = async () => {
    if (!formRef.current) return
    const fd = new FormData(formRef.current)
    const get = (name: string) => fd.get(name)?.toString() || ""

    const formatDateStr = (dateStr: string) => {
      if (!dateStr) return ""
      const d = new Date(dateStr)
      if (isNaN(d.getTime())) return dateStr
      return `${d.getFullYear()}年${String(d.getMonth()+1).padStart(2, '0')}月${String(d.getDate()).padStart(2, '0')}日`
    }

    // Attendees logic
    let attendees = ""
    if (projectScale === 'large') {
      attendees += `市公司政企部（解决方案、交付支撑、计划部）：\n        ${get('gen_city_attendees')}\n`
    }
    const branchName = get('gen_branch_name') || "XXXX"
    attendees += `${branchName}分公司（建设、维护、网络/信息安全员）：\n        ${get('gen_branch_attendees')}`

    const ctContentStr = hasMidThree ? (ctContent ? ctContent.replace(/能力/g, '') : "详见清单") : "无"
    
    const itCostInclForContent = (projectData.cost?.it?.integration?.incl || 0) + (projectData.cost?.it?.device?.incl || 0) + (projectData.cost?.it?.maintenance?.incl || 0)
    const itContentStr = itCostInclForContent > 0 ? (itContent || "集成服务") : "无"

    const otherCost = projectData.cost?.ct?.other?.incl || 0
    const otherProductContent = otherCost > 0 ? "详见清单" : "无"

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
          screenshotListArray.push({
            title: v.vendorName,
            data: img.data,
            width: img.width,
            height: img.height
          });
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

    const itConstruction = getExclIt('device') + getExclIt('construction') + getExclIt('survey') + getExclIt('integration') + getExclIt('other');
    const itMaintenance = getExclIt('maintenance') + getExclIt('running');
    const ctConstruction = getExclCt('construction');
    const ctMaintenance = getExclCt('maintenance');
    const ctProduct = getExclCt('other') + getExclCt('bandwidth') + getExclCt('renewal');
    const itOther = getExclIt('bidding') + getExclIt('design_eval') + getExclIt('audit');

    const isZero = (n: number) => Math.abs(n) < 0.005;
    const fmtYuan = (n: number) => n.toFixed(2);
    const fmtPct = (x: any) => isFinite(x) && x !== null && x !== "" && !isNaN(Number(x)) ? (Number(x) * 100).toFixed(2) + '%' : '--';
    const parts: string[] = [];
    const itParts: string[] = [];
    if (!isZero(itConstruction)) itParts.push(`建设投入${fmtYuan(itConstruction)}元（不含税）`);
    if (!isZero(itMaintenance)) itParts.push(`维护投入${fmtYuan(itMaintenance)}元（不含税）`);
    if (itParts.length) parts.push(`IT部分${itParts.join('，')}`);

    const ctParts: string[] = [];
    if (!isZero(ctConstruction)) ctParts.push(`建设投入${fmtYuan(ctConstruction)}元（不含税）`);
    if (!isZero(ctMaintenance)) ctParts.push(`维护投入${fmtYuan(ctMaintenance)}元（不含税）`);
    if (ctParts.length) parts.push(`CT部分${ctParts.join('，')}`);

    const ctLabel = hasMidThree ? midThreeName.replace(/能力/g, '').trim() : '专线';
    const finalCtLabel = ctLabel || '专线';
    if (!isZero(ctProduct)) parts.push(`CT-${finalCtLabel}投入${fmtYuan(ctProduct)}元（不含税）`);
    if (!isZero(itOther)) parts.push(`中标服务费/专项审计/第三方项目核算等费用${fmtYuan(itOther)}元（不含税）`);

    let projTotalInvestStr = `整体投入${fmtYuan(totalCost)}元`;
    if (parts.length) {
      projTotalInvestStr += `，其中${parts.join('；')}。`;
    } else {
      projTotalInvestStr += `。`;
    }

    // Calculate Demand Table specific fields
    const totalRevIt = Object.values(projectData.revenue?.it || {}).reduce((acc: number, curr: any) => acc + (curr?.incl || 0), 0);
    const totalRevCt = Object.values(projectData.revenue?.ct || {}).reduce((acc: number, curr: any) => acc + (curr?.incl || 0), 0);
    const totalRevNonItCt = projectData.revenue?.non_it_ct?.incl || 0;
    const totalRevIncl = Number(totalRevIt) + Number(totalRevCt) + Number(totalRevNonItCt);

    const totalRevItExcl = Object.values(projectData.revenue?.it || {}).reduce((acc: number, curr: any) => acc + (curr?.excl || 0), 0);
    const totalRevCtExcl = Object.values(projectData.revenue?.ct || {}).reduce((acc: number, curr: any) => acc + (curr?.excl || 0), 0);
    const totalRevNonItCtExcl = projectData.revenue?.non_it_ct?.excl || 0;
    const totalRevExcl = Number(totalRevItExcl) + Number(totalRevCtExcl) + Number(totalRevNonItCtExcl);

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
    const attach1ImageStr = attach1Images.length > 0 ? JSON.stringify(attach1Images) : "";
    const attach2ImageStr = (hasPublicUrl && attach2Images.length > 0) ? JSON.stringify(attach2Images) : "";

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
      
      'SUBJECT_IT_COST': get('gen_subject_it_cost') || subjectItCost,
      'SUBJECT_CT_COST': get('gen_subject_ct_cost') || subjectCtCost,
      'SUBJECT_IT_REV': get('gen_subject_it_rev') || subjectItRev,
      'SUBJECT_CT_REV': get('gen_subject_ct_rev') || subjectCtRev,
      'CONSTRUCTION_TIME_REQ': get('gen_construction_time_req'),
      'PROCUREMENT_METHOD': procurementMethod === '其他' ? get('gen_procurement_method_other') : procurementMethod,
      'CONSTRUCTION_INTERFACE': get('gen_construction_interface'),
      'RISK_OWNER': get('gen_risk_owner'),

      'IT_INQUIRY_PROCESS': itInquiryProcess,
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
      'EXCEL_REV_IT_INTEGRATION_EXCL': String(projectData.revenue?.it?.integration?.excl ?? 0),
      'EXCEL_REV_IT_MAINTENANCE_EXCL': String(projectData.revenue?.it?.maintenance?.excl ?? 0),
      'EXCEL_REV_IT_DEVICE_SALES_EXCL': String(projectData.revenue?.it?.device_sales?.excl ?? 0),
      'EXCEL_REV_IT_DEVICE_LEASE_EXCL': String(projectData.revenue?.it?.device_lease?.excl ?? 0),
      'EXCEL_REV_IT_OTHER_EXCL': String(projectData.revenue?.it?.other?.excl ?? 0),
      'EXCEL_REV_IT_CLOUD_EXCL': String(projectData.revenue?.it?.cloud?.excl ?? 0),
      'EXCEL_REV_CT_LINE_EXCL': String(projectData.revenue?.ct?.line?.excl ?? 0),
      'EXCEL_REV_CT_PRODUCT_INCL': String(projectData.revenue?.ct?.product?.incl ?? 0),
      'EXCEL_REV_NON_IT_CT_EXCL': String(projectData.revenue?.non_it_ct?.excl ?? 0),

      'EXCEL_COST_IT_DEVICE_EXCL': String(projectData.cost?.it?.device?.excl ?? 0),
      'EXCEL_COST_IT_CONSTRUCTION_EXCL': String(projectData.cost?.it?.construction?.excl ?? 0),
      'EXCEL_COST_IT_SURVEY_EXCL': String(projectData.cost?.it?.survey?.excl ?? 0),
      'EXCEL_COST_IT_INTEGRATION_EXCL': String(projectData.cost?.it?.integration?.excl ?? 0),
      'EXCEL_COST_IT_OTHER_EXCL': String(projectData.cost?.it?.other?.excl ?? 0),
      'EXCEL_COST_IT_MAINTENANCE_EXCL': String(projectData.cost?.it?.maintenance?.excl ?? 0),
      'EXCEL_COST_IT_RUNNING_EXCL': String(projectData.cost?.it?.running?.excl ?? 0),
      'EXCEL_COST_IT_BIDDING_EXCL': String(projectData.cost?.it?.bidding?.excl ?? 0),
      'EXCEL_COST_IT_DESIGN_EVAL_EXCL': String(projectData.cost?.it?.design_eval?.excl ?? 0),
      'EXCEL_COST_IT_AUDIT_EXCL': String(projectData.cost?.it?.audit?.excl ?? 0),

      'EXCEL_COST_CT_CONSTRUCTION_INCL': String(projectData.cost?.ct?.construction?.incl ?? 0),
      'EXCEL_COST_CT_MAINTENANCE_INCL': String(projectData.cost?.ct?.maintenance?.incl ?? 0),
      'EXCEL_COST_CT_OTHER_INCL': String(projectData.cost?.ct?.other?.incl ?? 0),
      'EXCEL_COST_CT_BANDWIDTH_EXCL': String(projectData.cost?.ct?.bandwidth?.excl ?? 0),
      'EXCEL_COST_CT_RENEWAL_EXCL': String(projectData.cost?.ct?.renewal?.excl ?? 0),

      'EXCEL_COST_NON_IT_CT_EXCL': String(projectData.cost?.mix?.non_it_ct?.excl ?? 0),
      'EXCEL_COST_MIX_MARKETING_EXCL': String(projectData.cost?.mix?.marketing?.excl ?? 0),
      'EXCEL_COST_MIX_CHANNEL_EXCL': String(projectData.cost?.mix?.channel?.excl ?? 0),
      'EXCEL_COST_MIX_OTHER_EXCL': String(projectData.cost?.mix?.other?.excl ?? 0),

      'RENEWAL_PROJECT_FLAG': (projectData.cost?.ct?.renewal?.excl ?? 0) > 0 ? "是" : "否",
      'CONTRACT_DURATION': String(projectData.basic?.project_years || 1),
    }

    try {
      const resultPath: string = await invoke('generate_lifecycle_docs', { 
          moduleId: "ict_lifecycle",
          variables: variables,
          selectedTemplates: [selectedTemplate]
      })
      if (confirm("✅ 生成成功！文件已保存至工作空间 output 目录。\n是否立即打开输出目录？")) {
        // We know the output dir is parent of the file, but we can just open the file's dir or use open_file
        // Actually, resolve_module_path 'output' is best.
        const modulePath: string = await invoke('get_module_path', { moduleId: 'ict_lifecycle' })
        invoke('open_file', { path: `${modulePath}/output` })
      }
    } catch(e) {
      alert("生成失败：" + e)
    }
  }

  return (
    <div className="flex flex-col gap-6">
      <form ref={formRef} className="flex flex-col gap-6" onSubmit={(e) => e.preventDefault()} onInput={handleFormInput}>
        
        {/* Excel 预算表/评估表专属配置 */}
        {selectedTemplate.endsWith('.xlsx') && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm flex flex-col gap-4">
            <h4 className="font-bold text-primary flex items-center gap-2">
              <span>📊</span> 《项目经济效益评估表》Excel 测算模版配置
            </h4>
            <p className="text-sm text-secondary-foreground leading-relaxed">
              您选择的是 Excel 预算/评估模板。系统将根据您在 **收入侧测算** 和 **支出侧测算** 中填写的精细化数据，自动将 **29 项输入指标、项目背景、产权归属、项目周期等基础信息** 全自动回填到 Excel 的对应 sheet 单元格中。
            </p>
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分商务模式</label>
                <select name="gen_it_bus_mode" value={itBusMode} onChange={e => setItBusMode(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md outline-none text-sm">
                  <option value="服务模式">服务模式</option>
                  <option value="集成购销">集成购销</option>
                  <option value="投资">投资</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分资金来源</label>
                <select name="gen_it_fund_src" value={itFundSrc} onChange={e => setItFundSrc(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md outline-none text-sm">
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
                <input type="date" name="gen_meet_start" defaultValue={todayStr} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">会审结束日期</label>
                <input type="date" name="gen_meet_end" defaultValue={todayStr} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">会审方式</label>
                <select name="gen_meet_mode" defaultValue="线上" className="bg-muted border border-border px-3 py-2 rounded-md">
                  <option value="线上">线上</option>
                  <option value="线下">线下</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">项目规模</label>
                <select name="gen_project_scale" value={projectScale} onChange={e => setProjectScale(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md">
                  <option value="large">大项目 (市/省)</option>
                  <option value="small">小项目 (分公司)</option>
                </select>
              </div>
              {projectScale === 'large' && (
                <div className="flex flex-col gap-1 col-span-2">
                  <label className="text-sm font-semibold">市公司政企部参会人员</label>
                  <input type="text" name="gen_city_attendees" placeholder="人员A、人员B" className="bg-muted border border-border px-3 py-2 rounded-md" />
                </div>
              )}
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">分公司参会人员</label>
                <div className="flex gap-2">
                  <input type="text" name="gen_branch_name" defaultValue="XXXX" placeholder="分公司名称" className="w-32 bg-muted border border-border px-3 py-2 rounded-md" />
                  <input type="text" name="gen_branch_attendees" placeholder="人员D、人员E" className="flex-1 bg-muted border border-border px-3 py-2 rounded-md" />
                </div>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">驻点支撑人员</label>
                <input type="text" name="gen_onsite_support" placeholder="如有请填写" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">项目背景</label>
                <textarea name="gen_proj_bg" rows={3} value={projectBackground} onChange={e => setProjectBackground(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold flex items-center justify-between">IT建设内容 <span className="text-xs text-secondary-foreground font-normal">根据项目名称自动生成</span></label>
                <textarea name="gen_it_content" value={itContent} onChange={e => setItContent(e.target.value)} rows={2} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold flex items-center justify-between">CT建设内容 <span className="text-xs text-secondary-foreground font-normal">中台能力联动修改</span></label>
                <textarea name="gen_ct_content" value={ctContent} onChange={e => setCtContent(e.target.value)} rows={2} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">技术方案</label>
                <textarea name="gen_tech_solution" rows={2} defaultValue="采用端-管-云架构..." className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>

              <div className="col-span-2 border border-border rounded-lg p-4 bg-background">
                <div className="flex justify-between items-center mb-3">
                  <label className="text-sm font-bold">技术方案可行性清单</label>
                  <button type="button" onClick={addTechItem} className="text-xs bg-primary/10 text-primary px-3 py-1.5 rounded font-semibold hover:bg-primary/20">+ 新增一行</button>
                </div>
                <div className="flex flex-col gap-2">
                  {techItems.map((item, i) => (
                    <div key={i} className="flex gap-2 items-center">
                      <input type="text" placeholder="服务名称" value={item.serviceName} onChange={e => updateTechItem(i, 'serviceName', e.target.value)} className="w-1/4 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <input type="text" placeholder="服务说明" value={item.serviceDesc} onChange={e => updateTechItem(i, 'serviceDesc', e.target.value)} className="flex-1 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <input type="number" placeholder="数量" value={item.amount} onChange={e => updateTechItem(i, 'amount', Number(e.target.value))} className="w-16 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <input type="text" placeholder="单位" value={item.unit} onChange={e => updateTechItem(i, 'unit', e.target.value)} className="w-16 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <button type="button" onClick={() => removeTechItem(i)} className="text-red-500 hover:bg-red-50 p-1.5 rounded">×</button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">自主三问</label>
                <input type="text" name="gen_self_three" defaultValue="自主集成，项目自主等级L1。" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">三化方案</label>
                <input type="text" name="gen_threeization" defaultValue="本项目不涉及三化方案。" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">战略价值</label>
                <input type="text" name="gen_strategic_value" placeholder="战略价值说明" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">结论</label>
                <input type="text" name="gen_tech_conclusion" defaultValue="方案可行同时能满足客户需求。" className="bg-muted border border-border px-3 py-2 rounded-md" />
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
                        className="w-1/3 bg-muted border border-border px-3 py-2 rounded-md text-sm" 
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
                        className="flex-1 bg-muted border border-border px-3 py-2 rounded-md text-sm" 
                      />
                      <button 
                        type="button" 
                        onClick={() => setIsMidThreeModalOpen(true)} 
                        className="px-3 bg-primary text-primary-foreground rounded-md hover:bg-primary/90 flex items-center justify-center shrink-0 transition-colors"
                        title="全局能力库"
                      >
                        <TableProperties className="w-4 h-4" />
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
                <input type="text" name="gen_rev_collection" value={revCollection} onChange={e => setRevCollection(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">支出侧付款方式</label>
                <input type="text" name="gen_exp_payment" value={expPayment} onChange={e => setExpPayment(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分商务模式</label>
                <select name="gen_it_bus_mode" value={itBusMode} onChange={e => setItBusMode(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md">
                  <option value="服务模式">服务模式</option>
                  <option value="集成购销">集成购销</option>
                  <option value="投资">投资</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">IT部分资金来源</label>
                <select name="gen_it_fund_src" value={itFundSrc} onChange={e => setItFundSrc(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md">
                  <option value="分公司成本开支">分公司成本开支</option>
                  <option value="市公司专项资源">市公司专项资源</option>
                </select>
              </div>

              <div className="col-span-2 border border-border rounded-lg p-4 bg-background">
                <div className="flex justify-between items-center mb-3">
                  <label className="text-sm font-bold">IT部分询价过程</label>
                  <div className="flex gap-2">
                    <button type="button" onClick={autoGenerateInquiry} className="text-xs bg-amber-100 text-amber-700 px-3 py-1.5 rounded font-bold hover:bg-amber-200">⚡ 一键生成三家报价</button>
                    <button type="button" onClick={addInqVendor} className="text-xs bg-primary/10 text-primary px-3 py-1.5 rounded font-semibold hover:bg-primary/20">+ 新增厂商</button>
                  </div>
                </div>
                <div className="flex flex-col gap-2">
                  {inqVendors.map((item, i) => (
                    <div key={i} className="flex gap-2 items-center">
                      <input type="text" placeholder="厂商名称" value={item.vendorName} onChange={e => updateInqVendor(i, 'vendorName', e.target.value)} className="flex-1 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <input type="number" placeholder="含税报价" value={item.amount === 0 ? '' : item.amount} onChange={e => updateInqVendor(i, 'amount', Number(e.target.value))} className="w-28 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <div className="flex items-center gap-1">
                        <input type="number" placeholder="税率" value={item.taxRate} onChange={e => updateInqVendor(i, 'taxRate', Number(e.target.value))} className="w-16 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                        <span className="text-xs text-secondary-foreground">%</span>
                      </div>
                      <input type="text" placeholder="备注" value={item.remark} onChange={e => updateInqVendor(i, 'remark', e.target.value)} className="w-20 bg-muted border border-border px-2 py-1.5 rounded-md text-sm" />
                      <button type="button" onClick={() => removeInqVendor(i)} className="text-red-500 hover:bg-red-50 p-1.5 rounded">×</button>
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
                        const updated = [...inqVendors];
                        const prevImages = updated[i].images || [];
                        const newImages = typeof updater === 'function' ? updater(prevImages) : updater;
                        updated[i].images = newImages;
                        setInqVendors(updated);
                      };

                      return (
                        <div key={i} className="border border-border/60 bg-muted/5 p-4 rounded-xl flex flex-col gap-3 transition-colors hover:border-primary/40">
                          <div className="flex justify-between items-center">
                            <span className="text-sm font-bold text-foreground flex items-center gap-1.5">
                              <span className="bg-primary/10 text-primary text-xs w-5 h-5 rounded-full flex items-center justify-center font-bold">{i + 1}</span>
                              <span className="text-primary font-extrabold">{v.vendorName}</span> 报价截图
                            </span>
                            {vendorImages.length > 0 && (
                              <span className="text-xs bg-emerald-500/10 text-emerald-600 px-2 py-0.5 rounded-full font-medium">
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
                            onChange={(e) => handleImageUpload(e, setVendorImages)} 
                          />

                          <div 
                            className="border border-dashed border-border rounded-lg p-5 text-center cursor-pointer hover:bg-muted/50 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all flex flex-col items-center justify-center gap-2 bg-muted/20"
                            onClick={(e) => e.currentTarget.focus()}
                            onDragOver={e => e.preventDefault()}
                            onDrop={e => { e.preventDefault(); handleImageUpload(e, setVendorImages); }}
                            onPaste={e => handleImageUpload(e, setVendorImages)}
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
                                className="text-[11px] bg-primary text-primary-foreground px-2.5 py-1 rounded font-semibold hover:bg-primary/90 transition-colors shadow-sm"
                              >
                                📂 选择本地图片
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
                                      setVendorImages(vendorImages.filter((_: any, idx: number) => idx !== imgIdx));
                                    }} 
                                    className="absolute top-1 right-1 bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity text-xs shadow-sm font-bold"
                                  >
                                    ×
                                  </button>
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
                <input type="text" name="gen_risk_owner" defaultValue="人员A" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">是否联合体投标</label>
                <input type="text" name="gen_is_joint" defaultValue="否" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">项目评审表准确完整</label>
                <input type="text" name="gen_review_acc" defaultValue="是，项目投入收入核算完整，各表填写准确" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>

              <div className="flex flex-col gap-1 col-span-2 mt-2">
                <label className="text-sm font-bold text-foreground flex items-center gap-2">
                  <input type="checkbox" checked={hasSingleSource} onChange={e => setHasSingleSource(e.target.checked)} className="w-4 h-4" />
                  是否涉及单一来源
                </label>
                {hasSingleSource && (
                  <textarea name="gen_single_source" rows={3} defaultValue="单一来源决策依据：符合单一来源场景..." className="bg-muted border border-border px-3 py-2 rounded-md" />
                )}
              </div>
              
              <div className="flex flex-col gap-1">
                <label className="text-sm font-bold text-foreground">采购方式</label>
                <select value={procurementMethod} onChange={e => setProcurementMethod(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md">
                  <option value="短名单甄选">短名单甄选</option>
                  <option value="采购">采购</option>
                  <option value="其他">其他</option>
                </select>
                {procurementMethod === '其他' && (
                  <input type="text" name="gen_procurement_method_other" placeholder="请输入其他采购方式" className="bg-muted border border-border px-3 py-2 rounded-md mt-1" />
                )}
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">时间要求</label>
                <textarea name="gen_construction_time_req" rows={2} defaultValue="合同签定后30天内。" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">售中建设及施工界面</label>
                <textarea name="gen_construction_interface" rows={2} defaultValue="本项目采购统一集成单位实施。分公司负责客户侧的协调工作，并协调管理合作伙伴完成交付。" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
            </div>
          </div>
        )}

        {/* 立项签批表专属配置 */}
        {selectedTemplate.includes('立项签批表') && (
          <div className="bg-card border border-border rounded-xl p-6 shadow-sm">
            <h4 className="font-bold text-primary mb-4">《ICT项目立项签批表》专属配置</h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">计费科目：IT投入</label>
                <input type="text" name="gen_subject_it_cost" value={subjectItCost} onChange={e => setSubjectItCost(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">计费科目：CT投入</label>
                <input type="text" name="gen_subject_ct_cost" value={subjectCtCost} onChange={e => setSubjectCtCost(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">计费科目：IT收入</label>
                <input type="text" name="gen_subject_it_rev" value={subjectItRev} onChange={e => setSubjectItRev(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">计费科目：CT收入</label>
                <input type="text" name="gen_subject_ct_rev" value={subjectCtRev} onChange={e => setSubjectCtRev(e.target.value)} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold flex items-center gap-2">
                  <input type="checkbox" name="gen_is_advance" className="w-4 h-4" />
                  是否涉及垫资
                </label>
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
                <input type="text" name="gen_demand_branch_name" defaultValue="XXX分公司" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-sm font-semibold">业务模式</label>
                <select name="gen_demand_it_business_mode" className="bg-muted border border-border px-3 py-2 rounded-md">
                  <option value="服务模式">服务模式</option>
                  <option value="投资">投资</option>
                </select>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">服务内容</label>
                <textarea name="gen_demand_service_content" defaultValue="IT；CT" rows={2} className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">设备清单</label>
                <input type="text" name="gen_demand_device_list" defaultValue="不涉及" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="col-span-2 border border-border rounded-lg p-4 bg-background">
                <div className="flex justify-between items-center mb-3">
                  <label className="text-sm font-bold text-foreground">技术方案可行性清单 (设备需求清单)</label>
                  <button type="button" onClick={addTechItem} className="text-xs bg-primary/10 text-primary px-3 py-1.5 rounded font-semibold hover:bg-primary/20">+ 新增一行</button>
                </div>
                <div className="flex flex-col gap-2">
                  {techItems.map((item, i) => (
                    <div key={i} className="flex gap-2 items-center">
                      <input type="text" placeholder="服务名称" value={item.serviceName} onChange={e => updateTechItem(i, 'serviceName', e.target.value)} className="w-1/4 bg-muted border border-border px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <input type="text" placeholder="服务说明" value={item.serviceDesc} onChange={e => updateTechItem(i, 'serviceDesc', e.target.value)} className="flex-1 bg-muted border border-border px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <input type="number" placeholder="数量" value={item.amount} onChange={e => updateTechItem(i, 'amount', Number(e.target.value))} className="w-16 bg-muted border border-border px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <input type="text" placeholder="单位" value={item.unit} onChange={e => updateTechItem(i, 'unit', e.target.value)} className="w-16 bg-muted border border-border px-2 py-1.5 rounded-md text-sm text-foreground" />
                      <button type="button" onClick={() => removeTechItem(i)} className="text-red-500 hover:bg-red-50 p-1.5 rounded">×</button>
                    </div>
                  ))}
                </div>
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">客户确认</label>
                <input type="text" name="gen_demand_customer_confirm" defaultValue="微信截图" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-semibold">部署环境要求</label>
                <input type="text" name="gen_demand_env_require" defaultValue="客户提供部署环境，不包含在本次项目范围内" className="bg-muted border border-border px-3 py-2 rounded-md" />
              </div>
              
              <div className="flex flex-col gap-1 col-span-2 mt-2">
                <label className="text-sm font-bold text-foreground flex items-center gap-2">
                  <input type="checkbox" checked={hasPublicUrl} onChange={e => {
                    setHasPublicUrl(e.target.checked);
                    if (!e.target.checked) setAttach2Images([]);
                  }} className="w-4 h-4" />
                  项目有效的公示网址及招标文件
                </label>
                {hasPublicUrl && (
                  <input type="text" name="gen_demand_public_url" placeholder="https://..." className="bg-muted border border-border px-3 py-2 rounded-md" />
                )}
              </div>

              <div className="flex flex-col gap-1 col-span-2">
                <label className="text-sm font-bold text-foreground flex items-center gap-2">
                  <input type="checkbox" checked={hasSecurity} onChange={e => setHasSecurity(e.target.checked)} className="w-4 h-4" />
                  信息安全、密评
                </label>
                {hasSecurity && (
                  <input type="text" name="gen_demand_security_detail" placeholder="例如：已做密评/待补充" className="bg-muted border border-border px-3 py-2 rounded-md" />
                )}
              </div>


              {/* Image Attachments */}
              <div className="flex flex-col gap-1 col-span-2 mt-2">
                <label className="text-sm font-bold text-foreground">附件1截图（客户确认材料）</label>
                <input type="file" multiple accept="image/*" className="hidden" ref={fileInput1Ref} onChange={(e) => handleImageUpload(e, setAttach1Images)} />
                <div 
                  className="border-2 border-dashed border-border rounded-lg p-6 text-center cursor-pointer hover:bg-muted/50 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all flex flex-col items-center justify-center gap-3 bg-muted/20"
                  onClick={(e) => e.currentTarget.focus()}
                  onDragOver={e => e.preventDefault()}
                  onDrop={e => { e.preventDefault(); handleImageUpload(e, setAttach1Images); }}
                  onPaste={e => handleImageUpload(e, setAttach1Images)}
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
                      className="text-xs bg-primary text-primary-foreground px-3 py-1.5 rounded-md font-semibold hover:bg-primary/90 transition-colors shadow-sm"
                    >
                      📂 选择本地图片
                    </button>
                  </div>
                </div>
                <div className="flex gap-2 flex-wrap mt-2">
                  {attach1Images.map((img, i) => (
                    <div key={i} className="relative w-24 h-24 border rounded overflow-hidden group">
                      <img src={img.data} className="w-full h-full object-cover" />
                      <button type="button" onClick={() => removeImage(i, setAttach1Images)} className="absolute top-1 right-1 bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity text-xs">×</button>
                    </div>
                  ))}
                </div>
              </div>

              {hasPublicUrl && (
                <div className="flex flex-col gap-1 col-span-2 mt-2">
                  <label className="text-sm font-bold text-foreground">附件2截图（招标文件/挂网截图）</label>
                  <input type="file" multiple accept="image/*" className="hidden" ref={fileInput2Ref} onChange={(e) => handleImageUpload(e, setAttach2Images)} />
                  <div 
                    className="border-2 border-dashed border-border rounded-lg p-6 text-center cursor-pointer hover:bg-muted/50 focus:border-primary focus:ring-2 focus:ring-primary/20 outline-none transition-all flex flex-col items-center justify-center gap-3 bg-muted/20"
                    onClick={(e) => e.currentTarget.focus()}
                    onDragOver={e => e.preventDefault()}
                    onDrop={e => { e.preventDefault(); handleImageUpload(e, setAttach2Images); }}
                    onPaste={e => handleImageUpload(e, setAttach2Images)}
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
                        className="text-xs bg-primary text-primary-foreground px-3 py-1.5 rounded-md font-semibold hover:bg-primary/90 transition-colors shadow-sm"
                      >
                        📂 选择本地图片
                      </button>
                    </div>
                  </div>
                  <div className="flex gap-2 flex-wrap mt-2">
                    {attach2Images.map((img, i) => (
                      <div key={i} className="relative w-24 h-24 border rounded overflow-hidden group">
                        <img src={img.data} className="w-full h-full object-cover" />
                        <button type="button" onClick={() => removeImage(i, setAttach2Images)} className="absolute top-1 right-1 bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity text-xs">×</button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

      </form>
      
      <button className="bg-primary text-primary-foreground font-bold py-3 px-6 rounded-lg self-start shadow-sm hover:opacity-90 transition-opacity" onClick={handleGenerate}>
        🚀 立即生成此文件
      </button>

      {isMidThreeModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
          <div className="bg-background rounded-xl shadow-lg w-full max-w-3xl max-h-[80vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between p-4 border-b border-border">
              <h3 className="font-bold text-lg flex items-center gap-2">
                <TableProperties className="w-5 h-5 text-primary" />
                全局能力库
              </h3>
              <button onClick={() => setIsMidThreeModalOpen(false)} className="p-1.5 rounded-full hover:bg-muted text-muted-foreground transition-colors"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-4 border-b border-border bg-muted/30">
              <div className="relative">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
                <input 
                  type="text" 
                  value={midThreeSearch} 
                  onChange={e => setMidThreeSearch(e.target.value)} 
                  placeholder="搜索能力名称、编号或所属类别..." 
                  className="w-full pl-9 pr-4 py-2.5 rounded-md border border-border bg-background text-sm focus:outline-none focus:ring-1 focus:ring-primary shadow-sm"
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
