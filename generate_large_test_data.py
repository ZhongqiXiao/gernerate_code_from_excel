import pandas as pd
import random
import string
import os
from tqdm import tqdm  # 用于显示进度条

# 生成单个18位随机字符串（数字+大写字母）
def generate_random_string():
    characters = string.ascii_uppercase + string.digits  # 大写字母和数字
    return ''.join(random.choice(characters) for _ in range(18))

# 生成大量测试数据并写入Excel文件
def generate_large_test_data(output_file, total_rows=100000, batch_size=10000):
    print(f"开始生成{total_rows}行测试数据...")
    
    # 创建一个生成器函数，避免一次性加载所有数据到内存
    def data_generator():
        for _ in range(total_rows):
            yield generate_random_string()
    
    # 创建一个空的DataFrame
    df = pd.DataFrame(columns=['Data'])
    
    # 创建ExcelWriter对象
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 分批写入数据
        current_row = 0
        
        # 使用tqdm显示进度
        with tqdm(total=total_rows, desc="生成并写入数据") as pbar:
            while current_row < total_rows:
                # 计算当前批次的行数
                rows_remaining = total_rows - current_row
                current_batch_size = min(batch_size, rows_remaining)
                
                # 生成当前批次的数据
                batch_data = [generate_random_string() for _ in range(current_batch_size)]
                
                # 创建当前批次的DataFrame
                batch_df = pd.DataFrame({'Data': batch_data})
                
                # 写入Excel文件
                if current_row == 0:
                    # 第一批次，写入表头
                    batch_df.to_excel(writer, index=False, startrow=current_row)
                else:
                    # 后续批次，不写入表头
                    batch_df.to_excel(writer, index=False, header=False, startrow=current_row + 1)  # +1 因为pandas从1开始计数
                
                # 更新当前行数
                current_row += current_batch_size
                
                # 更新进度条
                pbar.update(current_batch_size)
    
    print(f"测试数据已成功生成并保存到: {output_file}")
    print(f"文件大小: {os.path.getsize(output_file) / (1024 * 1024):.2f} MB")

if __name__ == "__main__":
    # 输出文件路径
    output_file = "large_test_data.xlsx"
    
    # 生成10万行测试数据
    generate_large_test_data(output_file, total_rows=100000)