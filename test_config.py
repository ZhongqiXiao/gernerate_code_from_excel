# -*- coding: utf-8 -*-
"""测试配置文件中的字典访问"""

import sys
import os

# 添加src目录到Python路径
project_root = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(project_root, 'src')
sys.path.append(src_path)

# 从core模块导入
from core.config import SUCCESS_TITLES, SUCCESS_MESSAGES, INFO_MESSAGES, ERROR_MESSAGES

# 测试访问各个字典的COMPLETE键
try:
    print("SUCCESS_TITLES['COMPLETE'] =", SUCCESS_TITLES['COMPLETE'])
except Exception as e:
    print("访问SUCCESS_TITLES['COMPLETE']出错:", str(e))

try:
    print("SUCCESS_MESSAGES['COMPLETE'] =", SUCCESS_MESSAGES['COMPLETE'])
except Exception as e:
    print("访问SUCCESS_MESSAGES['COMPLETE']出错:", str(e))

try:
    print("INFO_MESSAGES['COMPLETE'] =", INFO_MESSAGES['COMPLETE'])
except Exception as e:
    print("访问INFO_MESSAGES['COMPLETE']出错:", str(e))

# 打印所有字典内容以进行比较
print("\nSUCCESS_TITLES:", SUCCESS_TITLES)
print("SUCCESS_MESSAGES:", SUCCESS_MESSAGES)
print("INFO_MESSAGES:", INFO_MESSAGES)