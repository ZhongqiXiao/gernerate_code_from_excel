#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目配置文件，集中管理所有配置参数
"""

import os
import multiprocessing

# 系统设置
MAX_WORKERS = os.cpu_count() or 4  # 根据CPU核心数自动调整线程数
MAX_IMAGE_WORKERS = min(MAX_WORKERS, 4)  # 图像处理对内存要求较高，限制线程数

# 文件处理设置
BATCH_SIZE_EXCEL = 5000  # Excel文件读取的批次大小
BATCH_SIZE_QR = 100  # 二维码生成的批处理大小
QR_PER_IMAGE = 10  # 每个二维码图片包含的字符串数量
QR_PER_A4 = 15  # 每个A4页面包含的二维码数量
DEFAULT_QR_LENGTH = 3  # 二维码默认边长，单位厘米

# 辅助函数：根据二维码边长计算A4页面上可容纳的行列数
def calculate_a4_layout(qr_length_cm=DEFAULT_QR_LENGTH):
    """
    根据二维码边长计算A4页面上可容纳的最佳行列数
    
    Args:
        qr_length_cm (float): 二维码边长，单位厘米
        
    Returns:
        tuple: (rows, cols) - 最佳行数和列数
    """
    # 计算二维码边长（像素）
    # 1英寸 = 2.54厘米，600 DPI表示每英寸600像素
    qr_length_px = int(qr_length_cm / 2.54 * IMAGE_DPI)
    
    # 计算有效区域（减去边距）
    effective_width = A4_WIDTH - 2 * MARGIN_PIXELS
    effective_height = A4_HEIGHT - 2 * MARGIN_PIXELS
    
    # 计算行列数（向下取整，确保二维码不会超出页面）
    cols = effective_width // qr_length_px
    rows = effective_height // qr_length_px
    
    # 确保至少有1行1列
    cols = max(1, cols)
    rows = max(1, rows)
    
    return rows, cols

# 二维码设置
QR_VERSION = 2  # 二维码版本，增加版本以容纳更多数据
QR_ERROR_CORRECTION = 3  # ERROR_CORRECT_H级别（3），高纠错级别更适合打印
QR_BOX_SIZE = 12  # 二维码方块大小
QR_BORDER = 4  # 二维码边框大小

# 图像处理设置
IMAGE_DPI = 600  # 图像DPI值，影响打印质量
IMAGE_QUALITY = 85  # 图像保存质量，略微降低可提升速度
A4_WIDTH = 4960  # A4纸张宽度（像素，600 DPI）
A4_HEIGHT = 7016  # A4纸张高度（像素，600 DPI）
MARGIN_PIXELS = 200  # 图像边距（像素）

# GUI设置
DEFAULT_START_ROW = 1  # 默认开始行数
DEFAULT_OUTPUT_DIR = "."  # 默认输出目录
DEFAULT_BATCH_SIZE = 5000  # 默认批处理大小
WINDOW_WIDTH = 800  # GUI窗口宽度
WINDOW_HEIGHT = 500  # GUI窗口高度
APP_NAME = "二维码生成器"
APP_GEOMETRY = f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
RESIZABLE_WIDTH = True
RESIZABLE_HEIGHT = True

# 字体设置
try:
    # 尝试使用中文字体
    UI_FONT = ('SimHei', 10)
    UI_FONT_BOLD = ('SimHei', 16, 'bold')
    TEXT_FONT_SIZE = 48  # 二维码下方文字大小
    TEXT_FONT = "arial.ttf"  # 文字字体文件
except:
    # 回退到默认字体
    UI_FONT = None
    UI_FONT_BOLD = None
    TEXT_FONT = None

# 日志设置
LOG_LEVEL = "INFO"  # 日志级别: DEBUG, INFO, WARNING, ERROR

# 进度条设置
PROGRESS_INTERVAL = 500  # 进度条更新间隔（毫秒）
PROGRESS_INCREMENT = 0.5  # 每次更新的进度增量

# 路径设置
def get_temp_qr_dir(output_dir):
    """获取临时二维码目录路径"""
    return os.path.join(output_dir, 'temp_qr')

# 颜色设置
QR_FILL_COLOR = "black"  # 二维码填充颜色
QR_BACK_COLOR = "white"  # 二维码背景颜色
TEXT_COLOR = "black"  # 文字颜色
BACKGROUND_COLOR = "white"  # 背景颜色

# 错误标题模板
ERROR_TITLES = {
    "FILE_ERROR": "文件错误",
    "INPUT_ERROR": "输入错误",
    "DIR_ERROR": "目录错误",
    "GENERAL_ERROR": "发生错误"
}

# 警告标题模板
WARNING_TITLES = {
    "NO_DATA": "没有数据",
    "CANCEL": "取消确认"
}

# 警告消息模板
WARNING_MESSAGES = {
    "NO_DATA": "Excel文件中没有找到有效数据",
    "CANCEL_CONFIRM": "确定要取消当前操作吗？"
}

# 成功标题模板
SUCCESS_TITLES = {
    "COMPLETE": "操作完成",
    "SUCCESS_TITLE": "成功"
}

# 错误消息模板
ERROR_MESSAGES = {
    "FILE_NOT_FOUND": "请选择有效的Excel文件",
    "INVALID_START_ROW": "请输入有效的开始行数（必须大于等于1）",
    "INVALID_BATCH_SIZE": "请输入有效的批次大小（必须大于等于1）",
    "OUTPUT_DIR_ERROR": "请选择有效的输出目录",
    "NO_DATA": "没有读取到任何数据",
    "EXCEL_ERROR": "读取Excel文件时出错: {}",
    "QR_GENERATION_ERROR": "生成二维码时出错 (任务 {}): {}",
    "IMAGE_GENERATION_ERROR": "生成A4图片时出错 (页面 {}): {}",
    "GENERAL_ERROR": "程序执行出错: {}",
    "CREATE_DIR_ERROR": "创建目录时出错: {}"
}

# 成功消息模板
SUCCESS_MESSAGES = {
    "COMPLETE": "所有操作完成！",
    "FILE_GENERATED": "已生成图片: {}",
    "SUCCESS_TITLE": "成功",
    "SUCCESS_MESSAGE": "二维码生成完成！"
}

# 信息消息模板
INFO_MESSAGES = {
    "START_EXCEL_READ": "开始从第{}行读取Excel文件...",
    "EXCEL_READ_COMPLETE": "成功读取{}条数据",
    "START_QR_GENERATION": "开始生成二维码...(共{}批，使用多线程加速)",
    "START_IMAGE_GENERATION": "开始生成A4图片...(使用多线程加速)",
    "QR_GENERATION_COMPLETE": "生成{}个二维码耗时: {:.2f}秒",
    "IMAGE_GENERATION_COMPLETE": "生成A4图片耗时: {:.2f}秒",
    "COMPLETE": "所有操作完成！",
    "EXCEL_READ_TIME": "读取Excel文件耗时: {:.2f}秒",
    "TOTAL_TIME": "总用时: {:.2f}秒",
    "CANCELLED": "操作已取消",
    "BATCH_COMPLETED": "批次生成完成: 第{}批 - 共{}个二维码，用时: {:.2f}秒"
}