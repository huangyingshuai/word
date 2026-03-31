"""
模块名称：config.templates
模块职责：模板库管理、模板校验、模板应用逻辑
依赖模块：config.constants
"""
import copy
from typing import Dict, Tuple
from .constants import FONT_LIST, FONT_SIZE_LIST, ALIGN_LIST, LINE_TYPE_LIST

# ====================== 官方标准模板库 ======================
TEMPLATE_LIBRARY: Dict[str, Dict] = {
    "默认通用格式": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
        "序号列表": {"font": "宋体", "size": "小四", "bold": False, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "河北科技大学-本科毕业论文": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
        "序号列表": {"font": "宋体", "size": "小四", "bold": False, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "国标-本科毕业论文通用": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 12},
        "二级标题": {"font": "黑体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 12, "space_after": 6},
        "三级标题": {"font": "黑体", "size": "四号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "宋体", "size": "小四", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "宋体", "size": "五号", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
        "序号列表": {"font": "宋体", "size": "小四", "bold": False, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0}
    },
    "党政机关公文国标GB/T 9704-2012": {
        "一级标题": {"font": "黑体", "size": "二号", "bold": True, "align": "居中", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 6},
        "二级标题": {"font": "楷体", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 6},
        "三级标题": {"font": "仿宋", "size": "三号", "bold": True, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 6, "space_after": 3},
        "正文": {"font": "仿宋", "size": "三号", "bold": False, "align": "两端对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 2, "space_before": 0, "space_after": 0},
        "表格": {"font": "仿宋", "size": "小三", "bold": False, "align": "居中", "line_type": "单倍行距", "line_value": 1.0, "indent": 0, "space_before": 0, "space_after": 0},
        "序号列表": {"font": "仿宋", "size": "三号", "bold": False, "align": "左对齐", "line_type": "多倍行距", "line_value": 1.5, "indent": 0, "space_before": 0, "space_after": 0}
    }
}

# 默认模板基准
DEFAULT_TEMPLATE = copy.deepcopy(TEMPLATE_LIBRARY["默认通用格式"])

# ====================== 核心函数 ======================
def validate_template(template: Dict) -> Tuple[bool, str]:
    required_levels = ["一级标题", "二级标题", "三级标题", "正文", "表格", "序号列表"]
    required_properties = ["font", "size", "bold", "align", "line_type", "line_value"]
    
    for level in required_levels:
        if level not in template:
            return False, f"模板校验失败：缺少【{level}】的格式定义"
        
        level_config = template[level]
        for prop in required_properties:
            if prop not in level_config:
                return False, f"模板校验失败：【{level}】缺少必填字段【{prop}】"
        
        if level_config["font"] not in FONT_LIST:
            return False, f"模板校验失败：【{level}】的字体【{level_config['font']}】不支持"
        if level_config["size"] not in FONT_SIZE_LIST:
            return False, f"模板校验失败：【{level}】的字号【{level_config['size']}】不支持"
        if level_config["align"] not in ALIGN_LIST:
            return False, f"模板校验失败：【{level}】的对齐方式不合法"
        if level_config["line_type"] not in LINE_TYPE_LIST:
            return False, f"模板校验失败：【{level}】的行距类型不合法"
        if not isinstance(level_config["line_value"], (int, float)) or level_config["line_value"] <= 0:
            return False, f"模板校验失败：【{level}】的行距值必须大于0"
    
    return True, "模板格式校验通过"

def apply_template_to_config(
    template_name: str,
    keep_custom: bool = False,
    current_config: Dict = None
) -> Dict:
    if template_name not in TEMPLATE_LIBRARY:
        raise ValueError(f"模板【{template_name}】不存在")
    
    target_template = copy.deepcopy(TEMPLATE_LIBRARY[template_name])
    is_valid, msg = validate_template(target_template)
    if not is_valid:
        raise ValueError(msg)
    
    if not keep_custom or current_config is None:
        return target_template
    
    new_config = copy.deepcopy(current_config)
    for level in target_template.keys():
        if level not in new_config:
            new_config[level] = copy.deepcopy(target_template[level])
            continue
        for prop in target_template[level].keys():
            if prop not in new_config[level]:
                new_config[level][prop] = target_template[level][prop]
                continue
            if level in DEFAULT_TEMPLATE and prop in DEFAULT_TEMPLATE[level]:
                if new_config[level][prop] == DEFAULT_TEMPLATE[level][prop]:
                    new_config[level][prop] = target_template[level][prop]
    
    is_valid, msg = validate_template(new_config)
    if not is_valid:
        raise ValueError(f"生成配置校验失败：{msg}")
    
    return new_config