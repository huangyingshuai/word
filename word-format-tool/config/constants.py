"""
模块名称：config.constants
模块职责：全局常量定义，无任何业务逻辑与外部依赖，全项目唯一常量入口
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

# ====================== 对齐方式映射 ======================
ALIGN_MAP = {
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "不修改": None
}
ALIGN_LIST = list(ALIGN_MAP.keys())

# ====================== 行距类型映射 ======================
LINE_TYPE_MAP = {
    "单倍行距": WD_LINE_SPACING.SINGLE,
    "1.5倍行距": WD_LINE_SPACING.ONE_POINT_FIVE,
    "2倍行距": WD_LINE_SPACING.DOUBLE,
    "多倍行距": WD_LINE_SPACING.MULTIPLE,
    "固定值": WD_LINE_SPACING.EXACTLY
}
LINE_TYPE_LIST = list(LINE_TYPE_MAP.keys())

# ====================== 行距规则配置 ======================
LINE_RULE = {
    "单倍行距": {"default": 1.0, "min": 1.0, "max": 1.0, "step": 1.0, "label": "行距倍数"},
    "1.5倍行距": {"default": 1.5, "min": 1.5, "max": 1.5, "step": 0.1, "label": "行距倍数"},
    "2倍行距": {"default": 2.0, "min": 2.0, "max": 2.0, "step": 0.1, "label": "行距倍数"},
    "多倍行距": {"default": 1.5, "min": 0.5, "max": 5.0, "step": 0.1, "label": "行距倍数"},
    "固定值": {"default": 20.0, "min": 1.0, "max": 100.0, "step": 0.1, "label": "固定值(磅)"}
}

# ====================== 字体配置 ======================
FONT_LIST = ["宋体", "黑体", "微软雅黑", "楷体", "仿宋", "Times New Roman", "Arial"]
FONT_SIZE_LIST = ["初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六"]
FONT_SIZE_NUM = {k:v for k,v in zip(FONT_SIZE_LIST, [42.0,36.0,26.0,24.0,22.0,18.0,16.0,15.0,14.0,12.0,10.5,9.0,7.5,6.5])}
EN_FONT_LIST = ["和正文一致", "Times New Roman", "Arial", "Calibri"]

# ====================== 标题识别配置 ======================
TITLE_MAX_LENGTH = 60
TITLE_MIN_LENGTH = 2
LEVEL_TO_NAME = {1: "一级标题", 2: "二级标题", 3: "三级标题"}
NAME_TO_LEVEL = {v: k for k, v in LEVEL_TO_NAME.items()}

# ====================== 国标GB/T 1.1-2020 序号层级规范 ======================
GB_STANDARD_LEVEL_RULE = {
    1: ["中文数字序号"],
    2: ["括号中文数字"],
    3: ["阿拉伯数字多级", "括号阿拉伯数字"],
    4: ["括号阿拉伯数字", "带圈数字"],
    5: ["字母序号"]
}