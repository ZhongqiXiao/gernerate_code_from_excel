import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
import math
import argparse
import concurrent.futures
from typing import List, Tuple, Dict
import time

# 设置最大线程数，根据系统CPU核心数调整
MAX_WORKERS = os.cpu_count() or 4
BATCH_SIZE_QR = 100  # 二维码生成的批处理大小
def read_excel_in_batches(file_path: str, start_row: int, batch_size: int = 1000) -> List[str]:
    """
    分批读取Excel文件，避免内存溢出
    """
    all_strings = []
    # 计算需要跳过的行数（pandas从0开始计数）
    skip_rows = start_row - 1 if start_row > 1 else 0
    
    try:
        # 使用ExcelFile对象打开文件，并明确指定引擎为openpyxl
        xl = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_name = xl.sheet_names[0]  # 使用第一个工作表
        
        # 获取工作表对象
        sheet = xl.parse(sheet_name)
        
        # 计算数据总行数
        total_rows = len(sheet)
        
        # 逐批读取数据
        for i in range(skip_rows, total_rows, batch_size):
            end_row = min(i + batch_size, total_rows)
            # 读取当前批次的数据
            chunk = sheet.iloc[i:end_row]
            
            # 处理数据
            for _, row in chunk.iterrows():
                # 检查第一个单元格是否有值
                if pd.notna(row.iloc[0]):
                    all_strings.append(str(row.iloc[0]))
        
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        raise
    
    return all_strings

def create_qr_code(data: str, output_path: str) -> None:
    """
    创建高清二维码
    """
    qr = qrcode.QRCode(
        version=2,  # 增加版本以容纳更多数据
        error_correction=qrcode.constants.ERROR_CORRECT_H,  # 高纠错级别更适合打印
        box_size=12,  # 增加box_size使二维码更清晰
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    # 保存高清二维码，提高DPI值
    img.save(output_path, dpi=(600, 600))

def generate_qr_code_worker(data_group: Tuple[str, str, int, int]) -> Tuple[str, int, int]:
    """
    线程工作函数，用于并行生成二维码
    """
    data, output_dir, start_idx, end_idx = data_group
    qr_file = os.path.join(output_dir, f"qr_{start_idx}_{end_idx}.png")
    create_qr_code(data, qr_file)
    return (qr_file, start_idx, end_idx)

def generate_qr_codes(strings: List[str], output_dir: str) -> List[Tuple[str, int, int]]:
    """
    使用多线程并行生成二维码，每10个字符串生成一个二维码
    """
    qr_files = []
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 准备工作任务
    tasks = []
    for i in range(0, len(strings), 10):
        end_i = min(i + 10, len(strings))
        group = strings[i:end_i]
        data = ";".join(group)
        tasks.append((data, output_dir, i+1, end_i))
    
    # 分批提交任务到线程池，避免一次性创建过多任务
    # 注意：使用线程池而非进程池，因为GIL在图像处理时会释放
    start_time = time.time()
    
    # 使用有序字典来保存结果，确保顺序正确
    result_dict = {}
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # 提交所有任务
        future_to_idx = {
            executor.submit(generate_qr_code_worker, task): i 
            for i, task in enumerate(tasks)
        }
        
        # 收集结果
        for future in concurrent.futures.as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                result = future.result()
                result_dict[idx] = result
            except Exception as e:
                print(f"生成二维码时出错 (任务 {idx}): {e}")
    
    # 按原始顺序重建结果列表
    qr_files = [result_dict[i] for i in sorted(result_dict.keys())]
    
    end_time = time.time()
    print(f"生成{len(qr_files)}个二维码耗时: {end_time - start_time:.2f}秒")
    
    return qr_files

def process_a4_page_worker(page_data: Tuple[List[Tuple[str, int, int]], str, int, int]) -> str:
    """
    线程工作函数，用于并行处理A4页面
    """
    qr_files_group, output_dir, start_i, end_i = page_data
    
    # A4纸张尺寸（像素，600 dpi）
    a4_width = 4960
    a4_height = 7016
    
    # 创建A4大小的白色背景图片
    a4_image = Image.new('RGB', (a4_width, a4_height), color='white')
    draw = ImageDraw.Draw(a4_image)
    
    # 计算二维码的位置（3行5列的布局）
    cols = 3
    rows = math.ceil(len(qr_files_group) / cols)
    qr_width = (a4_width - 400) // cols  # 左右各留200像素边距
    qr_height = (a4_height - 400) // rows  # 上下各留200像素边距
    
    # 放置二维码
    for idx, (qr_file, start_num, end_num) in enumerate(qr_files_group):
        try:
            # 打开二维码图片
            qr_img = Image.open(qr_file)
            # 调整二维码大小，使用LANCZOS算法保持高质量
            qr_img = qr_img.resize((qr_width, qr_height), Image.Resampling.LANCZOS)
            
            # 计算位置
            col = idx % cols
            row = idx // cols
            x = 200 + col * qr_width  # 增加边距
            y = 200 + row * qr_height  # 增加边距
            
            # 粘贴二维码到A4图片
            a4_image.paste(qr_img, (x, y))
            
            # 添加文字说明（起止编号）
            try:
                # 尝试使用系统字体，增大字体大小
                font = ImageFont.truetype("arial.ttf", 48)  # 增大字体大小
            except:
                # 如果找不到字体，使用默认字体
                font = ImageFont.load_default()
            
            text = f"{start_num}-{end_num}"
            text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:4]
            text_x = x + (qr_width - text_width) // 2
            text_y = y + qr_height + 20  # 增加与二维码的间距
            draw.text((text_x, text_y), text, fill='black', font=font)
            
        except Exception as e:
            print(f"处理二维码 {qr_file} 时出错: {e}")
    
    # 保存A4图片
    if qr_files_group:
        start_num = qr_files_group[0][1]
        end_num = qr_files_group[-1][2]
        output_file = os.path.join(output_dir, f"{start_num}-{end_num}.png")
        a4_image.save(output_file, dpi=(600, 600), quality=95)  # 略微降低quality提升速度
        return output_file
    
    return ""

def create_a4_image(qr_files: List[Tuple[str, int, int]], output_dir: str) -> None:
    """
    使用多线程并行生成A4大小的图片，每15个二维码生成一个图片
    """
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 准备工作任务
    tasks = []
    for i in range(0, len(qr_files), 15):
        end_i = min(i + 15, len(qr_files))
        group = qr_files[i:end_i]
        tasks.append((group, output_dir, i, end_i))
    
    # 使用多线程并行处理A4页面
    start_time = time.time()
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(MAX_WORKERS, 2)) as executor:  # 图像合成对内存要求较高，限制线程数
        # 提交所有任务
        future_to_idx = {
            executor.submit(process_a4_page_worker, task): i 
            for i, task in enumerate(tasks)
        }
        
        # 收集结果
        for future in concurrent.futures.as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                result = future.result()
                if result:
                    print(f"已生成图片: {result}")
            except Exception as e:
                print(f"生成A4图片时出错 (页面 {idx}): {e}")
    
    end_time = time.time()
    print(f"生成A4图片耗时: {end_time - start_time:.2f}秒")

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='从Excel文件生成二维码图片')
    parser.add_argument('excel_file', help='Excel文件路径')
    parser.add_argument('n', type=int, nargs='?', default=1, help='从第几行开始读取数据（默认：1）')
    parser.add_argument('--output_dir', default='.', help='输出目录（默认：当前路径）')
    parser.add_argument('--batch_size', type=int, default=5000, help='分批读取的批次大小（默认：5000）')
    args = parser.parse_args()
    
    try:
        total_start_time = time.time()
        
        # 1. 分批读取Excel文件
        print(f"开始从第{args.n}行读取Excel文件...")
        start_time = time.time()
        strings = read_excel_in_batches(args.excel_file, args.n, args.batch_size)
        end_time = time.time()
        print(f"读取Excel文件耗时: {end_time - start_time:.2f}秒")
        
        if not strings:
            print("没有读取到任何数据")
            return
        
        print(f"成功读取{len(strings)}条数据")
        
        # 2. 生成临时二维码文件目录
        temp_qr_dir = os.path.join(args.output_dir, 'temp_qr')
        
        # 3. 生成二维码
        print("开始生成二维码...")
        qr_files = generate_qr_codes(strings, temp_qr_dir)
        
        # 4. 生成A4图片
        print("开始生成A4图片...")
        create_a4_image(qr_files, args.output_dir)
        
        total_end_time = time.time()
        print(f"所有操作完成！总耗时: {total_end_time - total_start_time:.2f}秒")
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        raise

if __name__ == "__main__":
    main()