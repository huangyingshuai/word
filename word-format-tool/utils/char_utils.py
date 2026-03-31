"""
模块名称：utils.char_utils
模块职责：字符处理通用工具，无业务依赖
"""
def full2half(s: str) -> str:
    result = []
    for char in s:
        code = ord(char)
        if code == 0x3000:
            result.append(chr(0x0020))
        elif 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        else:
            result.append(char)
    return ''.join(result)