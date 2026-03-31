"""
模块名称：core.number_recognizer
模块职责：全类型序号识别核心算法
"""
import re
from typing import Optional, Dict
from docx.text.paragraph import Paragraph
from utils.char_utils import full2half
from utils.doc_utils import is_protected_para

# ====================== 全类型序号定义 ======================
NUMBER_TYPE_DEF: Dict[str, Dict] = {
    "阿拉伯数字多级": {
        "pattern": re.compile(r"^(\s*)([0-9]+(\.[0-9]+)*[、.])\s*"),
        "level_calc": lambda match: len(match.group(2).split(".")),
        "gb_allowed_levels": [1,2,3,4,5],
        "gb_name": "阿拉伯数字多级序号"
    },
    "括号阿拉伯数字": {
        "pattern": re.compile(r"^(\s*)([（(][0-9]+[)）])\s*"),
        "level_calc": lambda match: 3,
        "gb_allowed_levels": [3,4],
        "gb_name": "括号阿拉伯数字序号"
    },
    "带圈数字": {
        "pattern": re.compile(r"^(\s*)([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚])\s*"),
        "level_calc": lambda match: 4,
        "gb_allowed_levels": [4,5],
        "gb_name": "带圈数字序号"
    },
    "字母序号": {
        "pattern": re.compile(r"^(\s*)([a-zA-Z][.)]|[(（][a-zA-Z][)）])\s*"),
        "level_calc": lambda match: 5,
        "gb_allowed_levels": [5],
        "gb_name": "英文字母序号"
    },
    "中文数字序号": {
        "pattern": re.compile(r"^(\s*)([一二三四五六七八九十百千]+[、.])\s*"),
        "level_calc": lambda match: 1,
        "gb_allowed_levels": [1],
        "gb_name": "中文数字序号"
    },
    "括号中文数字": {
        "pattern": re.compile(r"^(\s*)([（(][一二三四五六七八九十百千]+[)）])\s*"),
        "level_calc": lambda match: 2,
        "gb_allowed_levels": [2],
        "gb_name": "括号中文数字序号"
    },
}

ALL_NUMBER_PATTERNS = [(name, def_info["pattern"], def_info["level_calc"]) for name, def_info in NUMBER_TYPE_DEF.items()]
NUMBER_BLACKLIST = [
    re.compile(r"^图\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^表\s*[0-9一二三四五六七八九十]+[-.、:：]\s*", re.IGNORECASE),
    re.compile(r"^注\s*[0-9]*[：:.]\s*"),
]

# ====================== 核心识别函数 ======================
def identify_number_item(para: Paragraph, para_index: int = -1) -> Optional[Dict]:
    if is_protected_para(para):
        return None
    
    raw_text = para.text.strip()
    if not raw_text:
        return None
    
    for pattern in NUMBER_BLACKLIST:
        if pattern.match(raw_text):
            return None
    
    normalized_text = full2half(para.text)

    for type_name, pattern, level_calc in ALL_NUMBER_PATTERNS:
        match = pattern.match(normalized_text)
        if match:
            indent = len(match.group(1))
            number_text = match.group(2)
            base_level = level_calc(match)
            final_level = max(1, min(base_level + (indent // 2), 9))
            return {
                "type": type_name,
                "level": final_level,
                "number_text": number_text,
                "full_text": raw_text,
                "indent": indent,
                "para_index": para_index
            }
    return None