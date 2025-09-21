#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
二维码生成器图形用户界面
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
import time
# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

# 添加src目录到Python路径，确保本地运行和打包后都能正确找到模块
src_path = os.path.join(project_root, 'src')
if src_path not in sys.path:
    sys.path.append(src_path)

# 从core模块导入
from core.qrcode_processor import qr_processor, set_cancel_event, clear_cancel_event
from core.config import (
    APP_NAME, APP_GEOMETRY, RESIZABLE_WIDTH, RESIZABLE_HEIGHT,
    UI_FONT, DEFAULT_START_ROW, DEFAULT_OUTPUT_DIR, BATCH_SIZE_EXCEL,
    DEFAULT_QR_LENGTH, QR_PER_IMAGE, get_temp_qr_dir,
    ERROR_TITLES, ERROR_MESSAGES, WARNING_TITLES, WARNING_MESSAGES,
    INFO_MESSAGES, SUCCESS_TITLES, SUCCESS_MESSAGES
)
from core.qrcode_processor import DOCX_AVAILABLE

class QRCodeGeneratorGUI:
    def __init__(self, root):
        """初始化GUI界面"""
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry(APP_GEOMETRY)
        self.root.resizable(width=RESIZABLE_WIDTH, height=RESIZABLE_HEIGHT)
        
        # 设置中文字体
        self._set_font()
        
        # 变量
        self.excel_file_path = tk.StringVar()
        self.start_row_var = tk.StringVar(value=str(DEFAULT_START_ROW))
        self.output_dir_var = tk.StringVar(value=DEFAULT_OUTPUT_DIR)
        self.batch_size_var = tk.StringVar(value=str(BATCH_SIZE_EXCEL))
        self.qr_length_var = tk.StringVar(value=str(DEFAULT_QR_LENGTH))  # 二维码边长，单位厘米
        self.title_var = tk.StringVar(value="物料S/N清单")  # A4页面标题，默认为"物料S/N清单"
        self.output_format_var = tk.StringVar(value="image")  # 输出格式，默认为图片
        
        # 标志变量
        self.is_generating = False
        self.stop_event = threading.Event()
        self._progress_timers = []  # 存储进度条更新定时器ID
        
        # 创建界面
        self._create_widgets()
        
        # 确保中文显示
        self._ensure_chinese_display()
    
    def _set_font(self):
        """设置字体"""
        self.font = UI_FONT
        
    def _ensure_chinese_display(self):
        """确保中文显示正常"""
        # 确保中文显示，大多数情况下tkinter已经支持
        pass
        
    def _create_widgets(self):
        """创建界面组件"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding=(10, 10, 10, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.pack_propagate(False)
        
        # 第一行：Excel文件选择
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(file_frame, text="Excel文件：", font=self.font).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(file_frame, textvariable=self.excel_file_path, width=50, font=self.font).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="浏览...", command=self._browse_excel_file).pack(side=tk.RIGHT, padx=(5, 0))
        
        # 第二行：开始行、输出目录、批次大小
        settings_frame = ttk.Frame(main_frame)
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(settings_frame, text="开始行：", font=self.font).grid(row=0, column=0, padx=(0, 5), pady=5, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.start_row_var, width=10, font=self.font).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="输出目录：", font=self.font).grid(row=0, column=2, padx=(20, 5), pady=5, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.output_dir_var, width=30, font=self.font).grid(row=0, column=3, padx=5, pady=5)
        ttk.Button(settings_frame, text="浏览...", command=self._browse_output_dir).grid(row=0, column=4, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="批次大小：", font=self.font).grid(row=0, column=5, padx=(20, 5), pady=5, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.batch_size_var, width=10, font=self.font).grid(row=0, column=6, padx=5, pady=5)
        
        # 二维码边长设置
        ttk.Label(settings_frame, text="二维码边长(cm)：", font=self.font).grid(row=1, column=0, padx=(0, 5), pady=5, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.qr_length_var, width=10, font=self.font).grid(row=1, column=1, padx=5, pady=5)
        
        # 输出格式设置
        if DOCX_AVAILABLE:
            ttk.Label(settings_frame, text="输出格式：", font=self.font).grid(row=1, column=2, padx=(20, 5), pady=5, sticky=tk.W)
            output_format_frame = ttk.Frame(settings_frame)
            output_format_frame.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
            
            # 添加单选按钮组
            ttk.Radiobutton(output_format_frame, text="图片", variable=self.output_format_var, value="image", style='TRadiobutton').pack(side=tk.LEFT, padx=5)
            ttk.Radiobutton(output_format_frame, text="Word文档", variable=self.output_format_var, value="docx", style='TRadiobutton').pack(side=tk.LEFT, padx=5)
        else:
            # 如果python-docx库不可用，隐藏Word文档选项
            ttk.Label(settings_frame, text="输出格式：图片 (Word文档功能需要python-docx库)", font=self.font).grid(row=1, column=2, columnspan=2, padx=(20, 5), pady=5, sticky=tk.W)
            # 确保输出格式为图片
            self.output_format_var.set("image")
        
        # A4页面标题设置
        ttk.Label(settings_frame, text="A4页面标题：", font=self.font).grid(row=2, column=0, padx=(0, 5), pady=5, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.title_var, width=40, font=self.font).grid(row=2, column=1, columnspan=3, padx=5, pady=5)
        
        # 第三行：进度条
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X)
        
        self.progress_label = ttk.Label(progress_frame, text="准备就绪", font=self.font)
        self.progress_label.pack(side=tk.RIGHT)
        
        # 第四行：日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=10, font=self.font)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 第五行：按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.start_button = ttk.Button(button_frame, text="开始生成", command=self._start_generation)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = ttk.Button(button_frame, text="取消", command=self.cancel_generation, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        self.exit_button = ttk.Button(button_frame, text="退出", command=self.root.quit)
        self.exit_button.pack(side=tk.RIGHT, padx=5)
    
    def _browse_excel_file(self):
        """浏览Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_file_path.set(file_path)
    
    def _browse_output_dir(self):
        """浏览输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir_var.set(dir_path)
    
    def _start_generation(self):
        """开始生成二维码"""
        # 检查参数
        excel_file = self.excel_file_path.get()
        if not os.path.exists(excel_file):
            messagebox.showerror(ERROR_TITLES["FILE_ERROR"], ERROR_MESSAGES["FILE_NOT_FOUND"])
            return
        
        try:
            start_row = int(self.start_row_var.get())
            if start_row < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror(ERROR_TITLES["INPUT_ERROR"], ERROR_MESSAGES["INVALID_START_ROW"])
            return
        
        try:
            batch_size = int(self.batch_size_var.get())
            if batch_size <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(ERROR_TITLES["INPUT_ERROR"], ERROR_MESSAGES["INVALID_BATCH_SIZE"])
            return
        
        try:
            qr_length = float(self.qr_length_var.get())
            if qr_length <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(ERROR_TITLES["INPUT_ERROR"], "二维码边长必须为大于0的数字")
            return
        
        output_dir = self.output_dir_var.get()
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror(ERROR_TITLES["DIR_ERROR"], ERROR_MESSAGES["CREATE_DIR_ERROR"].format(str(e)))
                return
        
        # 禁用开始按钮，启用取消按钮
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        
        # 重置进度条
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        # 设置标志
        self.is_generating = True
        self.stop_event.clear()
        
        # 在新线程中生成二维码
        self.generation_thread = threading.Thread(
            target=self._generate_qrcodes,
            args=(excel_file, start_row, output_dir, batch_size, qr_length, self.title_var.get())
        )
        self.generation_thread.daemon = True
        self.generation_thread.start()
        
        # 检查线程是否结束
        self.root.after(100, self._check_thread)
    
    def _generate_qrcodes(self, excel_file, start_row, output_dir, batch_size, qr_length, title):
        """生成二维码的主函数"""
        try:
            # 设置取消事件
            set_cancel_event(self.stop_event)
            
            # 设置日志回调函数 - 将批次日志打印到控制台
            qr_processor.set_logger(self._log_console)
            
            # 1. 分批读取Excel文件
            self._log_gui(INFO_MESSAGES["START_EXCEL_READ"].format(start_row))
            self._log_console(INFO_MESSAGES["START_EXCEL_READ"].format(start_row))
            self._update_progress(10, "正在读取Excel文件...")
            
            start_time = time.time()
            strings = qr_processor.read_excel_in_batches(excel_file, start_row, batch_size)
            end_time = time.time()
            
            self._log_gui(INFO_MESSAGES["EXCEL_READ_TIME"].format(end_time - start_time))
            self._log_console(INFO_MESSAGES["EXCEL_READ_TIME"].format(end_time - start_time))
            
            if not strings:
                self._log_gui(ERROR_MESSAGES["NO_DATA"])
                self._log_console(ERROR_MESSAGES["NO_DATA"])
                self._update_progress(0, "没有找到数据")
                messagebox.showwarning(WARNING_TITLES["NO_DATA"], WARNING_MESSAGES["NO_DATA"])
                return
            
            self._log_gui(INFO_MESSAGES["EXCEL_READ_COMPLETE"].format(len(strings)))
            self._log_console(INFO_MESSAGES["EXCEL_READ_COMPLETE"].format(len(strings)))
            self._update_progress(30, "Excel文件读取完成")
            
            # 检查是否取消
            if self.stop_event.is_set():
                return
            
            # 2. 生成临时二维码文件目录
            temp_qr_dir = get_temp_qr_dir(output_dir)
            
            # 3. 生成二维码
            # 计算总批次数
            total_batches = (len(strings) + QR_PER_IMAGE - 1) // QR_PER_IMAGE
            self._log_gui(INFO_MESSAGES["START_QR_GENERATION"].format(total_batches))
            self._log_console(INFO_MESSAGES["START_QR_GENERATION"].format(total_batches))
            self._update_progress(40, "开始生成二维码...")
            
            # 创建进度更新回调函数
            qr_progress_range = 20  # 二维码生成占用20%的总进度
            def update_qr_progress(completed_batches):
                progress = 40 + (completed_batches / total_batches) * qr_progress_range
                self._update_progress(progress, f"正在生成二维码...({completed_batches}/{total_batches}批)")
            
            # 将进度更新回调函数传递给处理器
            qr_files = qr_processor.generate_qr_codes(strings, temp_qr_dir, progress_callback=update_qr_progress)
            
            self._update_progress(60, "二维码生成完成")
            
            # 检查是否取消
            if self.stop_event.is_set():
                return
            
            # 4. 根据用户选择的输出格式生成相应的文件
            output_format = self.output_format_var.get()
            
            if output_format == "image":
                # 生成A4图片
                self._log_gui(INFO_MESSAGES["START_IMAGE_GENERATION"])
                self._log_console(INFO_MESSAGES["START_IMAGE_GENERATION"])
                self._update_progress(70, "开始生成A4图片...")
                
                # 添加进度更新定时器
                self.a4_progress = 70  # 初始进度为70%
                self._update_a4_progress()
                
                qr_processor.create_a4_image(qr_files, output_dir, qr_length_cm=qr_length, title=title)
                
                # 取消A4图片生成进度更新定时器
                self._cancel_progress_timers()
            else:
                # 生成Word文档
                self._log_gui(INFO_MESSAGES["START_DOCX_GENERATION"])
                self._log_console(INFO_MESSAGES["START_DOCX_GENERATION"])
                self._update_progress(70, INFO_MESSAGES["START_DOCX_GENERATION"])
                
                # 添加进度更新定时器
                self.a4_progress = 70  # 初始进度为70%
                self._update_a4_progress()
                
                docx_file = qr_processor.create_docx_document(qr_files, output_dir, qr_length_cm=qr_length, title=title)
                
                if docx_file:
                    self._log_gui(INFO_MESSAGES["DOCX_FILE_GENERATED"].format(docx_file))
                    self._log_console(INFO_MESSAGES["DOCX_FILE_GENERATED"].format(docx_file))
                else:
                    self._log_gui(INFO_MESSAGES["DOCX_GENERATION_FAILED"])
                    self._log_console(INFO_MESSAGES["DOCX_GENERATION_FAILED"])
                    # 如果Word文档生成失败，尝试生成图片作为备选
                    self._log_gui(INFO_MESSAGES["TRYING_IMAGE_AS_FALLBACK"])
                    self._log_console(INFO_MESSAGES["TRYING_IMAGE_AS_FALLBACK"])
                    qr_processor.create_a4_image(qr_files, output_dir, qr_length_cm=qr_length, title=title)
                
                # 取消进度更新定时器
                self._cancel_progress_timers()
            
            self._update_progress(100, "完成")
            
            self._log_gui(INFO_MESSAGES["COMPLETE"])
            self._log_console(INFO_MESSAGES["COMPLETE"])
            self._log_gui(INFO_MESSAGES["TOTAL_TIME"].format(time.time() - start_time))
            self._log_console(INFO_MESSAGES["TOTAL_TIME"].format(time.time() - start_time))
            messagebox.showinfo(SUCCESS_TITLES["COMPLETE"], SUCCESS_MESSAGES["COMPLETE"])
            
        except Exception as e:
            error_msg = ERROR_MESSAGES["GENERAL_ERROR"].format(str(e))
            self._log_gui(error_msg)
            self._log_console(error_msg)
            self._update_progress(0, f"错误: {str(e)}")
            messagebox.showerror(ERROR_TITLES["GENERAL_ERROR"], error_msg)
        finally:
            # 清理资源
            self._cancel_progress_timers()
            clear_cancel_event()  # 清除取消事件
            
            # 恢复按钮状态
            self.is_generating = False
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.cancel_button.config(state=tk.DISABLED))
            
            # 删除临时属性
            if hasattr(self, 'qr_progress'):
                delattr(self, 'qr_progress')
            if hasattr(self, 'a4_progress'):
                delattr(self, 'a4_progress')
            if hasattr(self, '_operation_completed'):
                delattr(self, '_operation_completed')
    
    def _cancel_progress_timers(self):
        """取消所有进度条更新定时器"""
        for timer_id in self._progress_timers:
            self.root.after_cancel(timer_id)
        self._progress_timers.clear()
    
    def _update_qr_progress(self):
        """更新二维码生成进度"""
        if self.stop_event.is_set() or hasattr(self, '_operation_completed'):
            return
        
        # 每次更新增加0.5%，直到达到60%
        if self.qr_progress < 60:
            self.qr_progress += 0.5
            self._update_progress(self.qr_progress, "正在生成二维码...")
            
            # 保存定时器ID
            timer_id = self.root.after(500, self._update_qr_progress)
            self._progress_timers.append(timer_id)
    
    def _update_a4_progress(self):
        """更新A4图片生成进度"""
        if self.stop_event.is_set() or hasattr(self, '_operation_completed'):
            return
        
        # 每次更新增加0.5%，直到达到100%
        if self.a4_progress < 100:
            self.a4_progress += 0.5
            if self.a4_progress < 95:
                status = "正在生成A4图片..."
            else:
                status = "即将完成..."
            self._update_progress(self.a4_progress, status)
            
            # 保存定时器ID
            timer_id = self.root.after(500, self._update_a4_progress)
            self._progress_timers.append(timer_id)
    
    def _update_progress(self, value, status_text):
        """更新进度条和状态文本"""
        # 确保进度值在0-100之间
        value = max(0, min(100, value))
        
        # 更新进度条
        self.root.after(0, lambda: self.progress_var.set(value))
        
        # 更新状态文本
        status_text = f"{status_text} ({int(value)}%)"
        self.root.after(0, lambda: self.progress_label.config(text=status_text))
    
    def _log_gui(self, message):
        """在日志文本框中添加带时间戳的重要日志消息"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _log_console(self, message):
        """在控制台打印带时间戳的所有日志消息"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        print(log_message)
    
    def cancel_generation(self):
        """取消二维码生成"""
        if not self.is_generating:
            return
        
        if messagebox.askyesno(WARNING_TITLES["CANCEL"], WARNING_MESSAGES["CANCEL_CONFIRM"]):
            # 设置取消事件
            self.stop_event.set()
            self._log_gui(INFO_MESSAGES["CANCELLED"])
            self._log_console(INFO_MESSAGES["CANCELLED"])
            self._update_progress(0, "正在取消...")
            
            # 添加一个强制清理的定时器，确保即使线程没有响应也能恢复UI状态
            def force_cleanup():
                # 无论如何都要恢复按钮状态和标志
                self.is_generating = False
                self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.cancel_button.config(state=tk.DISABLED))
                self._update_progress(0, "已取消")
                
            # 500毫秒后强制清理，给正常取消过程一些时间
            self.root.after(500, force_cleanup)
    
    def _check_thread(self):
        """检查生成线程是否结束"""
        if self.is_generating:
            # 检查线程是否还在运行
            if not self.generation_thread.is_alive():
                # 线程已结束但标志未更新，强制更新状态
                self.is_generating = False
                self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.cancel_button.config(state=tk.DISABLED))
            else:
                self.root.after(100, self._check_thread)


def main():
    """主函数"""
    root = tk.Tk()
    app = QRCodeGeneratorGUI(root)
    
    # 设置窗口关闭事件处理
    def on_closing():
        """窗口关闭事件处理"""
        if app.is_generating:
            # 如果正在生成，先尝试取消
            app.cancel_generation()
            # 给取消操作一些时间完成，然后强制销毁窗口
            def force_close():
                # 关闭线程池
                qr_processor.shutdown()
                # 无论如何都要销毁窗口，确保程序退出
                root.destroy()
            
            # 设置一个计时器来强制关闭
            root.after(1000, force_close)
        else:
            # 关闭线程池
            qr_processor.shutdown()
            root.destroy()
    
    # 绑定窗口关闭事件
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    root.mainloop()

if __name__ == "__main__":
    main()