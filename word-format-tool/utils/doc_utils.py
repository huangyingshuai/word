"""
模块名称：utils.doc_utils
模块职责：docx文档相关通用工具函数
依赖模块：config.constants
"""
import re
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from docx.text.paragraph import Paragraph
from config.constants import FONT_SIZE_NUM

# ====================== 文档内容保护机制 ======================
def is_protected_para(para: Paragraph) -> bool:
    """
    升级版内容保护：包含图片/分页/分节/公式/OLE/文本框/书签/表格/脚注引用的段落完全不修改
    """
    if not para:
        return True
    try:
        # 新增：检查是否在表格内
        if para._element.getparent().tag.endswith('}tc'):
            return True
        
        # 分页/分节符保护
        if para.paragraph_format.page_break_before:
            return True
        if para._element.find('.//w:sectPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
            return True
        
        # 遍历run检查保护内容
        for run in para.runs:
            if run.contains_page_break:
                return True
            # 图片/形状保护
            if run._element.find('.//w:drawing', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            if run._element.find('.//w:pict', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            # 公式保护
            if run._element.find('.//m:oMath', namespaces={'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}) is not None:
                return True
            # OLE对象保护
            if run._element.find('.//o:OLEObject', namespaces={'o': 'urn:schemas-microsoft-com:office:office'}) is not None:
                return True
            # 文本框保护
            if run._element.find('.//w:txbxContent', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            # 书签保护
            if run._element.find('.//w:bookmarkStart', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            # 新增：超链接保护
            if run._element.find('.//w:hyperlink', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            # 新增：脚注/尾注引用保护
            if run._element.find('.//w:footnoteReference', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
            if run._element.find('.//w:endnoteReference', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True
        
        return False
    except Exception:
        return True

# ====================== 字体设置工具 ======================
def set_run_font(run, font_name: str, font_size: float, bold: bool = None):
    try:
        rPr = run._element.rPr
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            run._element.append(rPr)
        
        run.font.name = font_name
        rPr.rFonts.set(qn('w:eastAsia'), font_name)
        rPr.rFonts.set(qn('w:ascii'), font_name)
        rPr.rFonts.set(qn('w:hAnsi'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

def set_en_number_font(run, font_name: str, font_size: float, bold: bool = None):
    try:
        if font_name == "和正文一致":
            return
        rPr = run._element.rPr
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            run._element.append(rPr)
        
        run.font.name = font_name
        rPr.rFonts.set(qn('w:ascii'), font_name)
        rPr.rFonts.set(qn('w:hAnsi'), font_name)
        rPr.rFonts.set(qn('w:cs'), font_name)
        run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
    except Exception:
        pass

# ====================== 正文数字/英文处理工具 ======================
def process_number_in_para(para, body_font: str, body_size: float, number_config: dict):
    try:
        if not number_config["enable"]:
            for run in para.runs:
                set_run_font(run, body_font, body_size)
            return
        
        number_size = FONT_SIZE_NUM[number_config["size"]] if not number_config["size_same_as_body"] else body_size
        number_font = number_config["font"]
        number_bold = number_config["bold"]
        number_en_pattern = re.compile(r"[a-zA-Z0-9\.%\+\-]+")

        for run in para.runs:
            run_text = run.text
            if not run_text:
                continue
            
            if number_en_pattern.fullmatch(run_text):
                set_en_number_font(run, number_font, number_size, number_bold)
            elif number_en_pattern.search(run_text):
                original_bold = run.font.bold
                run.text = ""
                parts = []
                last_end = 0
                for match in number_en_pattern.finditer(run_text):
                    start, end = match.span()
                    if start > last_end:
                        parts.append(("text", run_text[last_end:start]))
                    parts.append(("number", run_text[start:end]))
                    last_end = end
                if last_end < len(run_text):
                    parts.append(("text", run_text[last_end:]))
                
                for part_type, part_text in parts:
                    new_run = para.add_run(part_text)
                    if part_type == "text":
                        set_run_font(new_run, body_font, body_size, original_bold)
                    else:
                        set_en_number_font(new_run, number_font, number_size, number_bold)
            else:
                set_run_font(run, body_font, body_size)
    except Exception:
        pass