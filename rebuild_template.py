import zipfile
import xml.etree.ElementTree as ET
import os
import shutil

template_path = "./项目全生命周期文件模版/_originals/【2024版】ICT项目售前方案会审纪要模板.docx"
output_path = "./项目全生命周期文件模版/【2024版】ICT项目售前方案会审纪要模板_变量版.docx"

shutil.copy(template_path, output_path)

file_path = output_path
backup_path = file_path + ".bak"
os.rename(file_path, backup_path)

# Register all namespaces
namespaces = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
}
for prefix, uri in namespaces.items():
    ET.register_namespace(prefix, uri)

with zipfile.ZipFile(backup_path, 'r') as zin, zipfile.ZipFile(file_path, 'w') as zout:
    for item in zin.infolist():
        content = zin.read(item.filename)
        if item.filename == "word/document.xml":
            root = ET.fromstring(content)
            for p in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
                t_nodes = list(p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
                if not t_nodes:
                    continue
                
                full_text = "".join([t.text for t in t_nodes if t.text])
                new_text = None
                
                if "关于" in full_text and "的售前方案会审" in full_text:
                    new_text = full_text.replace("xxxxxx", "{PROJECT_NAME}")
                elif "202X年XX月XX日 – 202X年XX月XX日" in full_text:
                    new_text = full_text.replace("202X年XX月XX日 – 202X年XX月XX日", "{MEETING_START_DATE} – {MEETING_END_DATE}")
                elif "线上|线下会审" in full_text:
                    new_text = full_text.replace("线上|线下", "{MEETING_MODE}")
                elif "市公司政企部（解决方案、交付支撑、计划部）" in full_text:
                    new_text = full_text.replace("市公司政企部（解决方案、交付支撑、计划部）", "{ATTENDEES}")
                elif "分公司（建设、维护、网络/信息安全员）" in full_text:
                    new_text = "" # Just clear this node's text
                elif "驻点支撑人员：" in full_text:
                    new_text = full_text.replace("驻点支撑人员： ", "驻点支撑人员：{ONSITE_SUPPORT}")
                elif "简要叙述项目背景。。。" in full_text:
                    new_text = full_text.replace("简要叙述项目背景。。。", "{PROJECT_BACKGROUND}")
                elif "建设内容包括：IT部分" in full_text:
                    new_text = full_text.replace("。。。。。如硬件、集成、维保等；", "{IT_CONTENT}；").replace("。。。。如X条互联网/数据专线、X张物联卡、X台云视讯等。", "{CT_CONTENT}")
                elif "技术方案：。。。。。" in full_text:
                    new_text = full_text.replace("。。。。。", "{TECH_SOLUTION}")
                elif "自主三问：与立项汇报材料模板填写内容一致" in full_text:
                    new_text = full_text.replace("与立项汇报材料模板填写内容一致", "{SELF_THREE_Q}")
                elif "中台三问：与立项汇报材料模板填写内容一致" in full_text:
                    new_text = full_text.replace("与立项汇报材料模板填写内容一致", "{MID_THREE_Q}")
                elif "三化方案：融三化，则写名称、揭榜省、实现的功能；未融则简要描述未融原因。" in full_text:
                    new_text = full_text.replace("融三化，则写名称、揭榜省、实现的功能；未融则简要描述未融原因。", "{THREEIZATION_PLAN}")
                elif "结论：方案可行同时能满足客户需求。" in full_text:
                    new_text = full_text.replace("方案可行同时能满足客户需求。", "{TECH_CONCLUSION}")
                elif "1、IT部分商务模式：投资/设备销售/服务购销" in full_text:
                    new_text = full_text.replace("投资/设备销售/服务购销", "{IT_BUSINESS_MODE}")
                elif "2、IT部分资金来源：分公司成本开支/资本开支" in full_text:
                    new_text = full_text.replace("分公司成本开支/资本开支", "{IT_FUNDING_SOURCE}")
                elif "3、项目整体投入金额：" in full_text:
                    new_text = "{PROJECT_TOTAL_INVESTMENT}"
                elif "4、IT部分询价过程：服务模式需填写该部分内容，原则上询价厂商不少于3家" in full_text:
                    new_text = full_text.replace("服务模式需填写该部分内容，原则上询价厂商不少于3家", "{IT_INQUIRY_PROCESS}")
                elif "5、是否涉及联合体投标：如填是，则应描述联合体投标情况，如主体附体，合作伙伴名称等。" in full_text:
                    new_text = full_text.replace("如填是，则应描述联合体投标情况，如主体附体，合作伙伴名称等。", "{IS_JOINT_BIDDING}")
                elif "6、收入侧收款方式：项目XXXX交付后客户单位/按服务进度付款等。" in full_text:
                    new_text = full_text.replace("项目XXXX交付后客户单位/按服务进度付款等。", "{REVENUE_COLLECTION_METHOD}")
                elif "7、支出侧付款方式： XXX后支付XXX%，XXXX之后支付剩余XXX%。" in full_text:
                    new_text = full_text.replace("XXX后支付XXX%，XXXX之后支付剩余XXX%。", "{EXPENDITURE_PAYMENT_METHOD}")
                elif "8、项目评审表适用及填写是否准确完整" in full_text:
                    new_text = "8、项目评审表适用及填写是否准确完整：{PROJECT_REVIEW_ACCURACY}"
                elif "9、单一来源情况说明" in full_text:
                    new_text = "{SINGLE_SOURCE_EXPLANATION}"
                elif "（1）单一来源决策依据：" in full_text or "（2）单一来源供应商：" in full_text or "（3）预计单一来源采购金额：" in full_text:
                    new_text = ""
                elif "此处CT指：" in full_text:
                    new_text = ""
                elif "审核要点：项目投入收入是否核算完整" in full_text:
                    new_text = ""
                elif "如为服务模式，需判定是否涉及垫资" in full_text:
                    new_text = ""
                
                if new_text is not None:
                    for i in range(1, len(t_nodes)):
                        t_nodes[i].text = ""
                    t_nodes[0].text = new_text
                    
            content = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
        zout.writestr(item, content)

os.remove(backup_path)
print("Template Rebuilt Successfully!")
