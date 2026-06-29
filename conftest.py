"""pytest 根配置：把项目根目录加入 sys.path，让测试能 import core 包。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
