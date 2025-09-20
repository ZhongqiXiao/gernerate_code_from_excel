# Excel 二维码生成工具

这个工具可以从Excel文件读取数据，将每10个字符串合并生成一个二维码，并将15个二维码排列在A4大小的图片上。

## 功能特点

1. 支持从指定行开始读取Excel数据
2. 每10个字符串生成一个二维码（最后一组可以不满10个）
3. 每15个二维码生成一张A4大小的图片
4. 图片命名格式为`start-end`，表示包含的字符串范围
5. 二维码和图片均为高清质量（300dpi）
6. 支持分批读取Excel文件，避免内存溢出

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python generate_qrcode_from_excel.py <excel_file> <n> [--output_dir <output_directory>] [--batch_size <batch_size>]
```

参数说明：
- `excel_file`: Excel文件路径
- `n`: 从第几行开始读取数据
- `--output_dir`: 输出目录（默认：output）
- `--batch_size`: 分批读取的批次大小（默认：1000）

## 示例

```bash
# 从第5行开始读取数据
python generate_qrcode_from_excel.py data.xlsx 5

# 自定义输出目录和批次大小
python generate_qrcode_from_excel.py data.xlsx 1 --output_dir result --batch_size 500
```

## 注意事项

1. 程序假设Excel数据在第一个工作表的第一列
2. 如需调整二维码布局或样式，可以修改代码中的相关参数
3. 生成的临时二维码文件保存在output目录下的temp_qr子目录中
4. 如果系统中没有Arial字体，程序会使用默认字体显示编号