# 二维码生成器

一个高效的批量二维码生成工具，支持从Excel文件读取数据，批量生成二维码并自动排版成A4图片。

## 功能特点

- 从Excel文件批量读取数据生成二维码
- 支持自定义开始行和批量处理大小
- 多线程处理，充分利用CPU资源
- 自动排版二维码为A4图片
- 提供图形界面（GUI）和命令行接口（CLI）
- 实时进度显示
- 详细的操作日志

## 项目结构

```
├── src/                     # 源码目录
│   ├── core/                # 核心功能模块
│   │   ├── qrcode_processor.py  # 二维码处理核心功能
│   │   └── config.py            # 配置文件
│   ├── gui/                 # 图形界面模块
│   │   └── qrcode_gui.py        # 图形界面入口
│   ├── utils/               # 工具脚本
│   │   └── generate_large_test_data.py  # 生成测试数据的工具
│   └── qrcode_cli.py        # 命令行接口入口
├── legacy/                  # 遗留代码（原始版本）
│   ├── generate_qrcode_from_excel.py  # 原始版本的二维码生成器
│   └── qrcode_generator_gui.py        # 原始版本的图形界面
├── requirements.txt         # 依赖列表
├── QRCodeGenerator.spec     # PyInstaller打包配置
├── .gitignore               # Git忽略文件配置
└── test_config.py           # 配置测试脚本
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 图形界面（GUI）

直接运行图形界面程序：

```bash
python src/gui/qrcode_gui.py
```

或者使用打包后的可执行文件：

```bash
QRCodeGenerator.exe
```

**使用步骤：**
1. 点击"浏览..."按钮选择Excel文件
2. 设置开始行、输出目录和批次大小
3. 点击"开始生成"按钮
4. 等待生成完成，查看输出目录

### 命令行接口（CLI）

```bash
python src/qrcode_cli.py [Excel文件路径] [开始行] [选项]
```

**参数说明：**
- `Excel文件路径`：必需，指定要读取的Excel文件
- `开始行`：可选，指定从第几行开始读取数据（默认为1）

**选项：**
- `--output_dir`：指定输出目录（默认为当前目录下的output文件夹）
- `--batch_size`：指定分批读取的批次大小（默认为100）

**示例：**

```bash
# 从第1行开始读取数据
python src/qrcode_cli.py data.xlsx

# 从第5行开始读取数据
python src/qrcode_cli.py data.xlsx 5

# 自定义输出目录和批次大小
python src/qrcode_cli.py data.xlsx 1 --output_dir ./results --batch_size 200
```

## 配置说明

可以在`config.py`文件中自定义以下配置：

- 线程池大小
- 批处理大小
- 二维码尺寸和纠错级别
- A4图片布局
- 字体设置
- 颜色配置
- 日志级别
- 消息模板

## 打包程序

使用PyInstaller打包成可执行文件：

```bash
pyinstaller QRCodeGenerator.spec
```

**注意：** 本项目的spec文件已针对PyInstaller v6.0及以上版本进行了优化，移除了不再支持的Windows特定参数。

打包后的程序将位于`dist`目录下。

## 测试数据生成

项目提供了生成测试数据的工具，可以生成大量模拟数据用于测试程序性能：

```bash
python src/utils/generate_large_test_data.py
```

默认生成100,000行测试数据到`test_data.xlsx`文件中。

## 注意事项

1. 确保Excel文件格式正确，数据应放在第一列
2. 大文件处理时建议适当调整批次大小以提高性能
3. 生成过程中请勿关闭程序或中断操作
4. 如遇问题可查看日志获取详细信息

## 开发环境

- Python 3.8+
- pandas (版本要求见requirements.txt)
- openpyxl 3.1.2
- qrcode 7.4.2
- Pillow 10.3.0
- tqdm 4.67.1