#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
二维码生成器命令行接口
"""

import argparse
import sys
import os
# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import time
from src.core.qrcode_processor import qr_processor
from src.core.config import *


def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='从Excel文件生成二维码图片')
    parser.add_argument('excel_file', help='Excel文件路径')
    parser.add_argument('n', type=int, nargs='?', default=DEFAULT_START_ROW, help=f'从第几行开始读取数据（默认：{DEFAULT_START_ROW}）')
    parser.add_argument('--output_dir', default=DEFAULT_OUTPUT_DIR, help=f'输出目录（默认：{DEFAULT_OUTPUT_DIR}）')
    parser.add_argument('--batch_size', type=int, default=BATCH_SIZE_EXCEL, help=f'分批读取的批次大小（默认：{BATCH_SIZE_EXCEL}）')
    args = parser.parse_args()
    
    try:
        total_start_time = time.time()
        
        # 1. 分批读取Excel文件
        info_msg = INFO_MESSAGES["START_EXCEL_READ"].format(args.n)
        print(info_msg)
        
        start_time = time.time()
        strings = qr_processor.read_excel_in_batches(args.excel_file, args.n, args.batch_size)
        end_time = time.time()
        
        info_msg = INFO_MESSAGES["EXCEL_READ_TIME"].format(end_time - start_time)
        print(info_msg)
        
        if not strings:
            print(ERROR_MESSAGES["NO_DATA"])
            return
        
        info_msg = INFO_MESSAGES["EXCEL_READ_COMPLETE"].format(len(strings))
        print(info_msg)
        
        # 2. 生成临时二维码文件目录
        temp_qr_dir = get_temp_qr_dir(args.output_dir)
        
        # 3. 生成二维码
        print(INFO_MESSAGES["START_QR_GENERATION"])
        qr_files = qr_processor.generate_qr_codes(strings, temp_qr_dir)
        
        # 4. 生成A4图片
        print(INFO_MESSAGES["START_IMAGE_GENERATION"])
        qr_processor.create_a4_image(qr_files, args.output_dir)
        
        total_end_time = time.time()
        info_msg = INFO_MESSAGES["TOTAL_TIME"].format(total_end_time - total_start_time)
        print(info_msg)
        
    except Exception as e:
        error_msg = ERROR_MESSAGES["GENERAL_ERROR"].format(str(e))
        print(error_msg)
        raise

if __name__ == "__main__":
    main()