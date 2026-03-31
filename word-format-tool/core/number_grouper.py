"""
模块名称：core.number_grouper
模块职责：序号智能分组、国标合规校验、序号格式转换
"""
import copy
from typing import List, Dict, Tuple
from datetime import datetime
from config.constants import GB_STANDARD_LEVEL_RULE
from core.number_recognizer import NUMBER_TYPE_DEF, identify_number_item

# 内置标准测试用例集
STANDARD_TEST_CASES = [
    {"text": "1. 研究背景", "expected_type": "阿拉伯数字多级", "expected_level": 1},
    {"text": "1.1 研究现状", "expected_type": "阿拉伯数字多级", "expected_level": 2},
    {"text": "1.1.1 国内研究", "expected_type": "阿拉伯数字多级", "expected_level": 3},
    {"text": "一、引言", "expected_type": "中文数字序号", "expected_level": 1},
    {"text": "（一）研究意义", "expected_type": "括号中文数字", "expected_level": 2},
    {"text": "(1) 实验方法", "expected_type": "括号阿拉伯数字", "expected_level": 3},
    {"text": "① 数据采集", "expected_type": "带圈数字", "expected_level": 4},
    {"text": "a. 样本选择", "expected_type": "字母序号", "expected_level": 5},
    {"text": "图1 系统架构", "expected_type": None, "expected_level": 0},
    {"text": "表2 实验结果", "expected_type": None, "expected_level": 0},
]

# ====================== 核心函数 ======================
def group_number_items(number_items: List[Dict]) -> List[Dict]:
    if not number_items:
        return []
    
    groups = []
    current_group = None
    
    for item in number_items:
        if current_group is None:
            current_group = {
                "group_id": len(groups) + 1,
                "type": item["type"],
                "level": item["level"],
                "start_index": item["para_index"],
                "end_index": item["para_index"],
                "items": [item]
            }
        else:
            if item["type"] == current_group["type"] and item["level"] == current_group["level"]:
                current_group["items"].append(item)
                current_group["end_index"] = item["para_index"]
            else:
                groups.append(current_group)
                current_group = {
                    "group_id": len(groups) + 1,
                    "type": item["type"],
                    "level": item["level"],
                    "start_index": item["para_index"],
                    "end_index": item["para_index"],
                    "items": [item]
                }
    
    if current_group:
        groups.append(current_group)
    
    return groups

def check_gb_compliance(number_type: str, level: int) -> Tuple[bool, List[str]]:
    is_compliant = True
    issues = []
    type_info = NUMBER_TYPE_DEF[number_type]

    allowed_levels = type_info["gb_allowed_levels"]
    if level not in allowed_levels:
        is_compliant = False
        issues.append(f"「{type_info['gb_name']}」国标推荐层级{allowed_levels}，当前{level}级")
    
    if level in GB_STANDARD_LEVEL_RULE:
        standard_types = GB_STANDARD_LEVEL_RULE[level]
        if number_type not in standard_types:
            is_compliant = False
            standard_names = [NUMBER_TYPE_DEF[t]["gb_name"] for t in standard_types]
            issues.append(f"国标{level}级推荐：{standard_names}")
    
    return is_compliant, issues

def calc_recognition_accuracy(test_cases: List[Dict]) -> Tuple[float, List[Dict]]:
    total = len(test_cases)
    correct = 0
    error_cases = []
    
    for case in test_cases:
        class MockPara:
            def __init__(self, text):
                self.text = text
                self.runs = []
                self.style = type('MockStyle', (), {'name': 'Normal'})()
                self.paragraph_format = type('MockFormat', (), {'page_break_before': False})()
                self._element = type('MockElement', (), {'find': lambda *args: None})()
        
        para = MockPara(case["text"])
        result = identify_number_item(para)
        
        if (result and result["type"] == case["expected_type"] and result["level"] == case["expected_level"]) or (not result and case["expected_type"] is None):
            correct += 1
        else:
            error_cases.append({
                "text": case["text"],
                "expected": {"type": case["expected_type"], "level": case["expected_level"]},
                "actual": {"type": result["type"] if result else None, "level": result["level"] if result else None}
            })
    
    accuracy = correct / total * 100 if total > 0 else 0
    return accuracy, error_cases

def save_correction_history(history_list: List, operation_desc: str, before_state, after_state):
    history_list.append({
        "operation": operation_desc,
        "before": copy.deepcopy(before_state),
        "after": copy.deepcopy(after_state),
        "timestamp": datetime.now().strftime("%H:%M:%S")
    })

# ====================== 【新增】序号格式转换函数 ======================
def convert_number_format(group: Dict, target_format: str = "1.") -> List[Dict]:
    """
    将序号组统一转换为指定格式
    :param group: 序号组信息
    :param target_format: 目标格式，支持："1."、"(1)"、"①"、"一、"
    :return: 转换后的序号列表
    """
    converted_items = []
    chinese_nums = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
                   "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"]
    
    for i, item in enumerate(group["items"], start=1):
        # 提取原序号后的内容（保留缩进）
        original_text = item["full_text"][len(item["number_text"]):].lstrip()
        
        # 根据目标格式生成新序号
        if target_format == "1.":
            new_number_text = f"{i}."
        elif target_format == "(1)":
            new_number_text = f"({i})"
        elif target_format == "①":
            # 带圈数字（1-20）
            if 1 <= i <= 20:
                new_number_text = chr(0x2460 + i - 1)
            else:
                new_number_text = f"{i}."  # 超过20回退到阿拉伯数字
        elif target_format == "一、":
            if 1 <= i <= len(chinese_nums):
                new_number_text = chinese_nums[i-1] + "、"
            else:
                new_number_text = f"{i}."  # 超过20回退到阿拉伯数字
        else:
            new_number_text = item["number_text"]
        
        converted_items.append({
            "type": group["type"],
            "level": group["level"],
            "number_text": new_number_text,
            "full_text": f"{new_number_text} {original_text}",
            "indent": item["indent"],
            "para_index": item["para_index"]
        })
    
    return converted_items