import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
import math
import argparse
from typing import List, Tuple

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

def generate_qr_codes(strings: List[str], output_dir: str) -> List[str]:
    """
    每10个字符串生成一个二维码
    """
    qr_files = []
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 分组处理字符串（每10个一组）
    for i in range(0, len(strings), 10):
        group = strings[i:i+10]
        data = ";".join(group)
        
        # 生成二维码文件路径
        qr_file = os.path.join(output_dir, f"qr_{i+1}_{min(i+10, len(strings))}.png")
        
        # 创建二维码
        create_qr_code(data, qr_file)
        qr_files.append((qr_file, i+1, min(i+10, len(strings))))
    
    return qr_files

def create_a4_image(qr_files: List[Tuple[str, int, int]], output_dir: str) -> None:
    """
    每15个二维码生成一个A4大小的图片
    """
    # A4纸张尺寸（像素，600 dpi，更高分辨率适合打印）
    a4_width = 4960
    a4_height = 7016
    
    # 每15个二维码一组
    for i in range(0, len(qr_files), 15):
        group = qr_files[i:i+15]
        
        # 创建A4大小的白色背景图片
        a4_image = Image.new('RGB', (a4_width, a4_height), color='white')
        draw = ImageDraw.Draw(a4_image)
        
        # 计算二维码的位置（尝试3行5列的布局）
        cols = 3
        rows = math.ceil(len(group) / cols)
        qr_width = (a4_width - 400) // cols  # 左右各留200像素边距，增加边距避免裁剪
        qr_height = (a4_height - 400) // rows  # 上下各留200像素边距，增加边距避免裁剪
        
        # 放置二维码
        for idx, (qr_file, start_num, end_num) in enumerate(group):
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
        
        # 保存A4图片，使用600 DPI提高打印质量
        if group:
            start_num = group[0][1]
            end_num = group[-1][2]
            output_file = os.path.join(output_dir, f"{start_num}-{end_num}.png")
            a4_image.save(output_file, dpi=(600, 600), quality=100)  # 增加quality参数
            print(f"已生成图片: {output_file}")

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='从Excel文件生成二维码图片')
    parser.add_argument('excel_file', help='Excel文件路径')
    parser.add_argument('n', type=int, nargs='?', default=1, help='从第几行开始读取数据（默认：1）')
    parser.add_argument('--output_dir', default='.', help='输出目录（默认：当前路径）')
    parser.add_argument('--batch_size', type=int, default=5000, help='分批读取的批次大小（默认：5000）')
    args = parser.parse_args()
    
    try:
        # 1. 分批读取Excel文件
        print(f"开始从第{args.n}行读取Excel文件...")
        strings = read_excel_in_batches(args.excel_file, args.n, args.batch_size)
        
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
        
        print("所有操作完成！")
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        raise

if __name__ == "__main__":
    main()