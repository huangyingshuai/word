"""
模块名称：utils.file_utils
模块职责：文件处理、临时文件管理、内存清理通用工具
"""
import os
import gc
import tempfile
from typing import Optional

def safe_save_temp_file(file_content: bytes, suffix: str = ".docx") -> Optional[str]:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(file_content)
            return tmp.name
    except Exception as e:
        print(f"临时文件保存失败：{str(e)}")
        return None

def safe_delete_file(file_path: str):
    if file_path and os.path.exists(file_path):
        try:
            os.unlink(file_path)
        except Exception:
            pass

def clear_memory(*args):
    for obj in args:
        if obj is not None:
            try:
                del obj
            except Exception:
                pass
    gc.collect()