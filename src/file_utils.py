import os
import re
from pypinyin import pinyin, Style as PinyinStyle

def windows_sort_key(s):
    """Windows文件排序的键函数，考虑中文拼音"""
    # 首先获取拼音首字母
    s = os.path.basename(s)
    result = []
    for char in s:
        if '\u4e00' <= char <= '\u9fa5':  # 如果是中文字符
            # 获取拼音首字母并转小写
            py = pinyin(char, style=PinyinStyle.FIRST_LETTER)
            if py:
                result.append(py[0][0].lower())
        else:
            # 非中文字符，用自然排序处理
            if char.isdigit():
                # 补充零，确保数字排序正确
                result.append(char.zfill(10))
            else:
                result.append(char.lower())
    
    return ''.join(result)

def log_debug(message, debug_mode=False):
    """输出调试日志"""
    if debug_mode:
        print(message)