"""
模块名称：core.title_recognizer
模块职责：四层标题识别算法+上下文层级校验
依赖模块：utils.doc_utils, config.constants
"""
import re
from typing import List, Tuple
from docx.text.paragraph import Paragraph
from utils.doc_utils import is_protected_para
from config.constants import TITLE_MAX_LENGTH, TITLE_MIN_LENGTH, LEVEL_TO_NAME, NAME_TO_LEVEL

# ====================== 标题识别规则 ======================
TITLE_BLACKLIST = [
    re.compile(r"^图\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^表\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^figure\s*[0-9]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^table\s*[0-9]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^注\s*[0-9]*[：:.]\s*"),
    re.compile(r"^参考文献\s*[:：]?$"),
    re.compile(r"^附录\s*[0-9A-Z]*[:：]?$"),
    re.compile(r"^[（(]\s*[0-9]+[)）]\s*.*[。？！；;]$"),
    re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*.*[。？！；;]$"),
    re.compile(r"^[a-zA-Z][.)]\s*.*[。？！；;]$"),
]

TITLE_RULE = {
    "三级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+\s*[^\s。？！；;]{2,60}$"),
        re.compile(r"^[（(][0-9]+[)）]\s*[^\s。？！；;]{2,60}$"),
    ],
    "二级标题": [
        re.compile(r"^[0-9]+\.[0-9]+\s*[^\s。？！；;]{2,50}$"),
        re.compile(r"^[（(][一二三四五六七八九十]+[)）]\s*[^\s。？！；;]{2,50}$"),
    ],
    "一级标题": [
        re.compile(r"^第[一二三四五六七八九十0-9]+章\s*[^\s。？！；;]{2,40}$"),
        re.compile(r"^[一二三四五六七八九十]+、\s*[^\s。？！；;]{2,40}$"),
    ]
}

# ====================== 核心算法 ======================
def get_title_level_with_context(
    para_list: List[Paragraph],
    enable_regex: bool = True,
    enable_font_feature: bool = True,
    enable_context_check: bool = True
) -> List[Tuple[int, str]]:
    total_para = len(para_list)
    result = [ (i, "正文") for i in range(total_para) ]
    max_level_appeared = 0

    candidate_titles = []
    for para_idx, para in enumerate(para_list):
        if is_protected_para(para):
            continue
        
        text = para.text.strip()
        text_length = len(text)

        if not text:
            continue
        # 修复：黑名单生效
        if any(pattern.match(text) for pattern in TITLE_BLACKLIST):
            continue
        if text.endswith(("。", "？", "！", "；", ".", "?", "!", ";")):
            continue
        if text_length > TITLE_MAX_LENGTH or text_length < TITLE_MIN_LENGTH:
            continue

        # 大纲级别校验
        try:
            outline_level = para.paragraph_format.outline_level
            if outline_level in [1, 2, 3]:
                level_name = LEVEL_TO_NAME[outline_level]
                candidate_titles.append( (para_idx, level_name, "大纲级别") )
                continue
        except Exception:
            pass

        # 内置样式校验
        style_name = para.style.name.lower()
        if "heading 1" in style_name or "标题 1" in style_name or "标题1" in style_name:
            candidate_titles.append( (para_idx, "一级标题", "内置样式") )
            continue
        if "heading 2" in style_name or "标题 2" in style_name or "标题2" in style_name:
            candidate_titles.append( (para_idx, "二级标题", "内置样式") )
            continue
        if "heading 3" in style_name or "标题 3" in style_name or "标题3" in style_name:
            candidate_titles.append( (para_idx, "三级标题", "内置样式") )
            continue

        if not enable_regex:
            continue

        # 编号模式校验
        matched_level = None
        for level_name in ["三级标题", "二级标题", "一级标题"]:
            for pattern in TITLE_RULE[level_name]:
                if pattern.match(text):
                    matched_level = level_name
                    break
            if matched_level:
                break
        if matched_level:
            candidate_titles.append( (para_idx, matched_level, "编号模式") )
            continue

        if not enable_font_feature:
            continue

        # 字体特征校验
        try:
            is_bold = any(run.font.bold for run in para.runs if run.font.bold is not None)
            is_larger_font = any(run.font.size and run.font.size.pt > 12 for run in para.runs)
            has_spacing = (para.paragraph_format.space_before and para.paragraph_format.space_before.pt > 0) or \
                          (para.paragraph_format.space_after and para.paragraph_format.space_after.pt > 0)
            
            if sum([is_bold, is_larger_font, has_spacing]) >= 2:
                candidate_titles.append( (para_idx, "一级标题", "字体特征") )
        except Exception:
            continue

    # 上下文校验
    for para_idx, level_name, match_type in candidate_titles:
        level = NAME_TO_LEVEL[level_name]

        if enable_context_check:
            if level == 2 and max_level_appeared < 1:
                result[para_idx] = (para_idx, "正文")
                continue
            if level == 3 and max_level_appeared < 2:
                result[para_idx] = (para_idx, "正文")
                continue
            if level > max_level_appeared + 1:
                result[para_idx] = (para_idx, "正文")
                continue

        result[para_idx] = (para_idx, level_name)
        max_level_appeared = max(max_level_appeared, level)

    return result

def get_title_level(para: Paragraph, enable_regex: bool = True) -> str:
    result = get_title_level_with_context([para], enable_regex, enable_font_feature=False, enable_context_check=False)
    return result[0][1]