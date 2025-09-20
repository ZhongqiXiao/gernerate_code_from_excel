import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import sys
import threading
import concurrent.futures
from generate_qrcode_from_excel import read_excel_in_batches, generate_qr_codes, create_a4_image

class QRCodeGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("二维码生成器")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 设置中文字体
        self.font = ('SimHei', 10)
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建输入控件
        self.create_widgets()
        
        # 禁用最大化按钮
        self.root.attributes('-toolwindow', False)
    
    def create_widgets(self):
        # 标题
        title_label = ttk.Label(self.main_frame, text="Excel二维码生成器", font=('SimHei', 16, 'bold'))
        title_label.pack(pady=10)
        
        # 文件选择
        file_frame = ttk.Frame(self.main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="Excel文件:", font=self.font).pack(side=tk.LEFT, padx=5)
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=40, font=self.font)
        file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        browse_btn = ttk.Button(file_frame, text="浏览...", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT, padx=5)
        
        # 开始行设置
        row_frame = ttk.Frame(self.main_frame)
        row_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(row_frame, text="开始行数:", font=self.font).pack(side=tk.LEFT, padx=5)
        self.start_row_var = tk.StringVar(value="1")
        row_entry = ttk.Entry(row_frame, textvariable=self.start_row_var, width=10, font=self.font)
        row_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(row_frame, text="(从第几行开始读取数据，默认1)", font=self.font).pack(side=tk.LEFT, padx=5)
        
        # 输出目录设置
        output_frame = ttk.Frame(self.main_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="输出目录:", font=self.font).pack(side=tk.LEFT, padx=5)
        self.output_dir_var = tk.StringVar(value=".")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, width=40, font=self.font)
        output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        output_btn = ttk.Button(output_frame, text="浏览...", command=self.browse_output_dir)
        output_btn.pack(side=tk.RIGHT, padx=5)
        
        # 批次大小设置
        batch_frame = ttk.Frame(self.main_frame)
        batch_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(batch_frame, text="批次大小:", font=self.font).pack(side=tk.LEFT, padx=5)
        self.batch_size_var = tk.StringVar(value="5000")
        batch_entry = ttk.Entry(batch_frame, textvariable=self.batch_size_var, width=10, font=self.font)
        batch_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(batch_frame, text="(分批读取的批次大小，默认5000)", font=self.font).pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress_frame = ttk.Frame(self.main_frame)
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5)
        
        # 状态标签
        self.status_var = tk.StringVar(value="准备就绪")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, font=self.font, foreground="blue")
        self.status_label.pack(pady=5, anchor=tk.W)
        
        # 按钮区域
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.pack(pady=20)
        
        self.generate_btn = ttk.Button(btn_frame, text="开始生成", command=self.start_generation)
        self.generate_btn.pack(side=tk.LEFT, padx=10)
        
        self.cancel_btn = ttk.Button(btn_frame, text="取消", command=self.cancel_generation, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=10)
        
        self.exit_btn = ttk.Button(btn_frame, text="退出", command=self.root.destroy)
        self.exit_btn.pack(side=tk.LEFT, padx=10)
        
        # 日志文本框
        log_frame = ttk.LabelFrame(self.main_frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=10, font=self.font)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 禁止编辑日志
        self.log_text.config(state=tk.DISABLED)
        
        # 线程控制
        self.stop_event = threading.Event()
        self.generation_thread = None
        
        # 进度条定时器控制
        self.progress_timers = []
        
        # 为了支持取消操作，我们需要一个方法来监控底层的并发执行
        self.executor = None
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if filename:
            self.file_path_var.set(filename)
    
    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir_var.set(directory)
    
    def log(self, message):
        """向日志文本框中添加消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_var.set(message)
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
    
    def start_generation(self):
        """开始生成二维码的线程"""
        # 验证输入
        excel_file = self.file_path_var.get().strip()
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("错误", "请选择有效的Excel文件")
            return
        
        try:
            start_row = int(self.start_row_var.get().strip())
            if start_row < 1:
                raise ValueError("开始行数必须大于等于1")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的开始行数")
            return
        
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showerror("错误", "请选择有效的输出目录")
            return
        
        try:
            batch_size = int(self.batch_size_var.get().strip())
            if batch_size < 1:
                raise ValueError("批次大小必须大于等于1")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的批次大小")
            return
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 禁用按钮
        self.generate_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        
        # 重置线程事件
        self.stop_event.clear()
        
        # 清空进度条定时器列表
        self.progress_timers = []
        
        # 启动生成线程
        self.generation_thread = threading.Thread(
            target=self.generate_qrcodes,
            args=(excel_file, start_row, output_dir, batch_size)
        )
        self.generation_thread.daemon = True
        self.generation_thread.start()
        
        # 检查线程是否完成
        self.check_thread()
    
    def generate_qrcodes(self, excel_file, start_row, output_dir, batch_size):
        """生成二维码的核心函数"""
        try:
            # 1. 分批读取Excel文件
            self.root.after(0, lambda: self.log(f"开始从第{start_row}行读取Excel文件..."))
            self.root.after(0, lambda: self.update_progress(10))
            
            strings = read_excel_in_batches(excel_file, start_row, batch_size)
            
            if self.stop_event.is_set():
                self.root.after(0, lambda: self.log("操作已取消"))
                return
            
            if not strings:
                self.root.after(0, lambda: self.log("没有读取到任何数据"))
                return
            
            self.root.after(0, lambda: self.log(f"成功读取{len(strings)}条数据"))
            self.root.after(0, lambda: self.update_progress(30))
            
            # 2. 生成临时二维码文件目录
            temp_qr_dir = os.path.join(output_dir, 'temp_qr')
            os.makedirs(temp_qr_dir, exist_ok=True)
            
            # 3. 生成二维码（多线程处理）
            self.root.after(0, lambda: self.log("开始生成二维码...(使用多线程加速)"))
            self.root.after(0, lambda: self.update_progress(40))
            
            if self.stop_event.is_set():
                self.root.after(0, lambda: self.log("操作已取消"))
                return
            
            # 定期更新进度条
            def update_qr_progress():
                if not self.stop_event.is_set():
                    # 估算进度
                    if hasattr(self, 'qr_progress'):
                        self.qr_progress = min(60, self.qr_progress + 0.5)
                    else:
                        self.qr_progress = 40
                    self.root.after(0, lambda: self.update_progress(self.qr_progress))
                    if self.qr_progress < 60:
                        timer = self.root.after(500, update_qr_progress)
                        self.progress_timers.append(timer)
            
            timer = self.root.after(500, update_qr_progress)
            self.progress_timers.append(timer)
            
            qr_files = generate_qr_codes(strings, temp_qr_dir)
            
            # 清除二维码生成阶段的进度条定时器
            self._cancel_progress_timers()
            
            if self.stop_event.is_set():
                self.root.after(0, lambda: self.log("操作已取消"))
                return
            
            self.root.after(0, lambda: self.update_progress(70))
            
            # 4. 生成A4图片（多线程处理）
            self.root.after(0, lambda: self.log("开始生成A4图片...(使用多线程加速)"))
            
            if self.stop_event.is_set():
                self.root.after(0, lambda: self.log("操作已取消"))
                return
            
            # 定期更新进度条
            def update_a4_progress():
                if not self.stop_event.is_set():
                    # 估算进度
                    if hasattr(self, 'a4_progress'):
                        self.a4_progress = min(95, self.a4_progress + 0.5)
                    else:
                        self.a4_progress = 70
                    self.root.after(0, lambda: self.update_progress(self.a4_progress))
                    if self.a4_progress < 95:
                        timer = self.root.after(500, update_a4_progress)
                        self.progress_timers.append(timer)
            
            timer = self.root.after(500, update_a4_progress)
            self.progress_timers.append(timer)
            
            create_a4_image(qr_files, output_dir)
            
            # 清除A4图片生成阶段的进度条定时器
            self._cancel_progress_timers()
            
            if self.stop_event.is_set():
                self.root.after(0, lambda: self.log("操作已取消"))
                return
            
            # 设置最终进度为100%
            self.root.after(0, lambda: self.update_progress(100))
            self.root.after(0, lambda: self.log("所有操作完成！"))
            self.root.after(0, lambda: messagebox.showinfo("成功", "二维码生成完成！"))
            
        except Exception as e:
            error_msg = f"程序执行出错: {str(e)}"
            self.root.after(0, lambda: self.log(error_msg))
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
        finally:
            # 清除所有进度条定时器
            self._cancel_progress_timers()
            
            # 只有在发生错误或取消时才重置进度条，成功完成时保持100%
            if self.stop_event.is_set() or not hasattr(self, '_operation_completed'):
                self.root.after(0, lambda: self.update_progress(0))
            
            # 恢复按钮状态
            self.root.after(0, lambda: self.generate_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.cancel_btn.config(state=tk.DISABLED))
            
            # 清理临时属性
            if hasattr(self, 'qr_progress'):
                delattr(self, 'qr_progress')
            if hasattr(self, 'a4_progress'):
                delattr(self, 'a4_progress')
            if hasattr(self, '_operation_completed'):
                delattr(self, '_operation_completed')
    
    def _cancel_progress_timers(self):
        """取消所有进度条更新定时器"""
        for timer in self.progress_timers:
            try:
                self.root.after_cancel(timer)
            except:
                pass  # 如果定时器已经被取消，忽略异常
        self.progress_timers = []
    
    def cancel_generation(self):
        """取消生成二维码"""
        if messagebox.askyesno("确认取消", "确定要取消生成二维码吗？"):
            self.stop_event.set()
            self.log("正在取消操作...")
    
    def check_thread(self):
        """检查生成线程是否完成"""
        if self.generation_thread and self.generation_thread.is_alive():
            self.root.after(100, self.check_thread)

if __name__ == "__main__":
    # 确保中文正常显示
    root = tk.Tk()
    app = QRCodeGeneratorGUI(root)
    root.mainloop()