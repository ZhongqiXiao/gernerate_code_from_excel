# 二维码生成器

这个工具可以从Excel文件读取数据，将每10个字符串合并生成一个二维码，并将15个二维码排列在A4大小的图片上。支持命令行和图形界面两种使用方式，并已优化为多线程处理以提升大量数据的处理速度。

## 功能特点

1. 支持从指定行开始读取Excel数据
2. 每10个字符串生成一个二维码（最后一组可以不满10个）
3. 每15个二维码生成一张A4大小的图片
4. 图片命名格式为`start-end`，表示包含的字符串范围
5. 二维码和图片均为高清质量（300dpi）
6. 支持分批读取Excel文件，避免内存溢出
7. **多线程处理优化**：自动根据CPU核心数调整线程数，大幅提升批量处理速度
8. **图形界面支持**：提供简洁易用的GUI界面，支持参数输入、进度显示和操作取消

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 命令行版本

```bash
python generate_qrcode_from_excel.py <excel_file> <n> [--output_dir <output_directory>] [--batch_size <batch_size>]
```

参数说明：
- `excel_file`: Excel文件路径
- `n`: 从第几行开始读取数据（默认：1）
- `--output_dir`: 输出目录（默认：当前目录）
- `--batch_size`: 分批读取的批次大小（默认：5000）

## 示例

```bash
# 从第5行开始读取数据
python generate_qrcode_from_excel.py data.xlsx 5

# 自定义输出目录和批次大小
python generate_qrcode_from_excel.py data.xlsx 1 --output_dir result --batch_size 5000
```

### 图形界面版本

直接运行`dist`目录下的`QRCodeGenerator.exe`可执行文件，或使用以下命令启动图形界面：

```bash
python qrcode_generator_gui.py
```

图形界面支持以下功能：
- Excel文件选择
- 设置开始行数（默认为1）
- 设置输出目录（默认为当前目录）
- 设置批次大小（默认为5000）
- 实时显示生成进度
- 显示操作日志
- 支持取消正在进行的操作

## 注意事项

1. 程序假设Excel数据在第一个工作表的第一列
2. 如需调整二维码布局或样式，可以修改代码中的相关参数
3. 生成的临时二维码文件保存在输出目录下的temp_qr子目录中
4. 如果系统中没有Arial字体，程序会使用默认字体显示编号
5. 多线程优化会根据您的CPU核心数自动调整线程数量，建议在处理大量数据时使用较大的批次大小
6. 为了确保二维码在A4纸上的顺序与Excel中的数据顺序一致，程序使用有序字典来管理数据