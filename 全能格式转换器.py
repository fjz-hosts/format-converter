import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import sys
import tempfile
import shutil
from pdf2docx import Converter
import comtypes.client
import win32com.client
import ffmpeg
from PIL import Image  # 用于图片处理
import subprocess  # 添加subprocess模块
import platform  # 添加platform模块

# 尝试导入docx2pdf的convert函数，如果失败则提供备选方案
try:
    from docx2pdf import convert
    docx2pdf_available = True
except ImportError:
    docx2pdf_available = False

def extract_ffmpeg():
    """从程序资源中提取ffmpeg到临时目录，支持PyInstaller打包的--add-data参数"""
    try:
        # 在临时目录创建一个子目录
        temp_dir = tempfile.mkdtemp()
        ffmpeg_path = os.path.join(temp_dir, "ffmpeg.exe")
        
        # 从程序资源中读取ffmpeg
        if getattr(sys, 'frozen', False):
            # 打包后的环境，处理PyInstaller --add-data的情况
            if sys.platform.startswith('win'):
                # Windows系统路径分隔符处理
                base_path = sys._MEIPASS
                source_path = os.path.join(base_path, "ffmpeg.exe")
                
                # 检查是否存在于MEIPASS目录
                if os.path.exists(source_path):
                    shutil.copy2(source_path, ffmpeg_path)
                    return ffmpeg_path
                else:
                    # 尝试其他可能的路径
                    alternative_paths = [
                        os.path.join(os.path.dirname(sys.executable), "ffmpeg.exe"),
                        os.path.join(os.path.dirname(os.path.abspath(__file__)), "ffmpeg.exe")
                    ]
                    for path in alternative_paths:
                        if os.path.exists(path):
                            shutil.copy2(path, ffmpeg_path)
                            return ffmpeg_path
                    
                    raise FileNotFoundError("无法在打包资源中找到ffmpeg.exe")
        else:
            # 开发环境，直接返回本地ffmpeg路径
            script_dir = os.path.dirname(os.path.abspath(__file__))
            dev_ffmpeg_path = os.path.join(script_dir, "ffmpeg.exe")
            if os.path.exists(dev_ffmpeg_path):
                # 复制到临时目录使用，避免开发环境文件被占用
                shutil.copy2(dev_ffmpeg_path, ffmpeg_path)
                return ffmpeg_path
            else:
                raise FileNotFoundError(f"开发环境中未找到ffmpeg.exe，路径: {dev_ffmpeg_path}")
                
        return ffmpeg_path
    except Exception as e:
        print(f"提取ffmpeg失败: {str(e)}")
        return None

class FormatConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("全能格式转换工具")
        self.root.configure(bg="#f0f2f5")
        self.root.minsize(1050, 750)  # 保持最小尺寸限制
        self.set_icon("图片1.ico")
        self.root.iconbitmap(default="")  # 避免打包时图标报错
          
        # 检查docx2pdf是否可用
        if not docx2pdf_available:
            messagebox.showwarning(
                "警告", 
                "未检测到docx2pdf库，Word转PDF功能可能无法使用。\n"
                "请运行 'pip install docx2pdf' 安装该库。"
            )
        
        # 检查PIL是否可用
        try:
            from PIL import Image
            self.pil_available = True
        except ImportError:
            self.pil_available = False
            messagebox.showwarning(
                "警告", 
                "未检测到Pillow库，图片格式转换功能无法使用。\n"
                "请运行 'pip install pillow' 安装该库。"
            )
        
        # 设置中文字体和ttk样式
        self.style = ttk.Style()
        self.style.configure(".", font=("微软雅黑", 10))
        self.style.configure("TLabel", background="#f0f2f5")
        self.style.configure("TRadiobutton", background="#f0f2f5")
        self.style.configure("TFrame", background="#f0f2f5")  # 配置ttk.Frame的背景色
        
        # 为ICO尺寸框架创建特定样式
        self.style.configure("IcoFrame.TFrame", background="#f0f2f5")
        
        # 提取或获取ffmpeg路径
        self.ffmpeg_path = extract_ffmpeg()
        # 检查ffmpeg是否可用
        self.check_ffmpeg_available()
        
        self.file_paths = []  # 改为支持多个文件路径
        self.output_dir = os.path.expanduser("~/转换输出")
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 添加Excel转换选项
        self.excel_fit_to_page = tk.BooleanVar(value=True)  # 默认为自动调整到一页
        self.excel_orientation = tk.StringVar(value="landscape")  # 默认为横向
        
        # ICO转换相关设置
        self.ico_sizes = [(16,16), (24,24), (32,32), (48,48), (64,64), 
                         (96,96), (128,128), (144,144), (192,192), (256,256)]
        self.selected_sizes = [tk.BooleanVar(value=True) for _ in self.ico_sizes]
        
        # 批量转换相关变量
        self.batch_mode = tk.BooleanVar(value=False)  # 批量模式开关
        self.current_file_index = 0  # 当前转换的文件索引
        self.total_files = 0  # 总文件数
        
        self.create_widgets()

    def set_icon(self, image_path):
        """设置应用程序图标 - 适配打包环境"""
        try:
            if getattr(sys, 'frozen', False):
                # 打包后的情况
                base_dir = sys._MEIPASS
            else:
                # 开发时的情况
                base_dir = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_dir, image_path)
            
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                print(f"图标文件未找到：{icon_path}")
        except Exception as e:
            print(f"设置图标失败：{e}")
            
    def check_ffmpeg_available(self):
        """检查ffmpeg是否可用"""
        if not self.ffmpeg_path or not os.path.exists(self.ffmpeg_path):
            messagebox.showwarning(
                "警告", 
                f"无法获取ffmpeg.exe，音视频转换功能将无法使用。\n"
                f"如果是开发环境，请确保ffmpeg.exe与程序在同一目录。"
            )
    
    def create_widgets(self):
        # 标题
        title_frame = tk.Frame(self.root, bg="#f0f2f5")
        title_frame.pack(pady=20, fill=tk.X, padx=20)
        
        title_label = tk.Label(
            title_frame, 
            text="全能格式转换工具", 
            font=("微软雅黑", 20, "bold"),
            bg="#f0f2f5",
            fg="#1a73e8"
        )
        title_label.pack()
        
        # 转换类型选择
        type_frame = tk.Frame(self.root, bg="#f0f2f5")
        type_frame.pack(pady=10, fill=tk.X, padx=20)
        
        ttk.Label(type_frame, text="转换类型:").pack(side=tk.LEFT, padx=5, pady=5)
        
        # 创建水平滚动框架
        scroll_frame = ttk.Frame(type_frame)
        scroll_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        canvas = tk.Canvas(scroll_frame, bg="#f0f2f5", height=50)
        canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        scrollbar = ttk.Scrollbar(scroll_frame, orient="horizontal", command=canvas.xview)
        scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        canvas.configure(xscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        type_options = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=type_options, anchor="nw")
        
        self.conversion_type = tk.StringVar(value="pdf_to_word")
        conversion_types = [
            ("PDF转Word", "pdf_to_word"),
            ("Word转PDF", "word_to_pdf"),
            ("Excel转PDF", "excel_to_pdf"),
            ("PPT转PDF", "ppt_to_pdf"),
            ("音频格式转换", "audio_convert"),
            ("视频格式转换", "video_convert"),
            ("图片格式转换", "image_convert")
        ]
        
        for text, value in conversion_types:
            ttk.Radiobutton(
                type_options, 
                text=text, 
                variable=self.conversion_type, 
                value=value
            ).pack(side=tk.LEFT, padx=15, pady=5)
        
        # Excel转换选项
        self.excel_options_frame = tk.Frame(self.root, bg="#f0f2f5", relief=tk.RIDGE, bd=2)
        ttk.Label(self.excel_options_frame, text="Excel转PDF设置:", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W, padx=10, pady=5)
        
        excel_settings = ttk.Frame(self.excel_options_frame, style="IcoFrame.TFrame")
        excel_settings.pack(fill=tk.X, padx=10, pady=5)
        
        # 自动调整到一页选项
        ttk.Checkbutton(
            excel_settings,
            text="自动调整表格到单页",
            variable=self.excel_fit_to_page
        ).pack(side=tk.LEFT, padx=15, pady=5)
        
        # 页面方向选项
        ttk.Label(excel_settings, text="页面方向:").pack(side=tk.LEFT, padx=(15,5), pady=5)
        ttk.Radiobutton(
            excel_settings,
            text="横向",
            variable=self.excel_orientation,
            value="landscape"
        ).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Radiobutton(
            excel_settings,
            text="纵向",
            variable=self.excel_orientation,
            value="portrait"
        ).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 格式选择（音视频、图片）
        self.format_frame = tk.Frame(self.root, bg="#f0f2f5")
        self.format_frame.pack(pady=5, fill=tk.X, padx=20)
        
        ttk.Label(self.format_frame, text="目标格式:").pack(side=tk.LEFT, padx=5, pady=5)
        
        self.target_format = tk.StringVar(value="mp3")
        self.format_options = ttk.Combobox(
            self.format_frame, 
            textvariable=self.target_format,
            state="readonly",
            values=["mp3", "wav", "flac", "m4a"]
        )
        self.format_options.pack(side=tk.LEFT, padx=5, pady=5)
        
        # ICO转换选项
        self.ico_options_frame = tk.Frame(self.root, bg="#f0f2f5", relief=tk.RIDGE, bd=2)
        ttk.Label(self.ico_options_frame, text="ICO图标尺寸 (选择需要包含的尺寸):", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W, padx=10, pady=5)
        
        # 为ICO尺寸选项添加水平滚动功能
        ico_scroll_container = ttk.Frame(self.ico_options_frame)
        ico_scroll_container.pack(fill=tk.X, padx=10, pady=5, expand=False)
        
        # 创建水平滚动条和画布
        ico_canvas = tk.Canvas(ico_scroll_container, bg="#f0f2f5", height=80)
        ico_scrollbar = ttk.Scrollbar(ico_scroll_container, orient="horizontal", command=ico_canvas.xview)
        
        # 配置画布和滚动条
        ico_canvas.configure(xscrollcommand=ico_scrollbar.set)
        ico_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        ico_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 创建放置尺寸选项的框架
        sizes_frame = ttk.Frame(ico_canvas, style="IcoFrame.TFrame")
        ico_canvas_window = ico_canvas.create_window((0, 0), window=sizes_frame, anchor="nw")
        
        # 添加尺寸选项
        for i, size in enumerate(self.ico_sizes):
            ttk.Checkbutton(
                sizes_frame,
                text=f"{size[0]}x{size[1]}",
                variable=self.selected_sizes[i]
            ).pack(side=tk.LEFT, padx=10, pady=15)
        
        # 绑定事件以确保滚动正常工作
        def on_frame_configure(event):
            ico_canvas.configure(scrollregion=ico_canvas.bbox("all"))
        
        def on_canvas_configure(event):
            ico_canvas.itemconfig(ico_canvas_window, width=event.width)
        
        sizes_frame.bind("<Configure>", on_frame_configure)
        ico_canvas.bind("<Configure>", on_canvas_configure)
        
        # 添加鼠标滚轮支持
        def on_mouse_wheel(event):
            ico_canvas.xview_scroll(-int(event.delta/120), "units")
        
        ico_canvas.bind_all("<MouseWheel>", on_mouse_wheel)
        
        # 文件选择区域
        file_frame = tk.Frame(self.root, bg="#f0f2f5")
        file_frame.pack(pady=10, fill=tk.X, padx=20)
        
        ttk.Label(file_frame, text="源文件:").pack(side=tk.LEFT, padx=5, pady=5)
        
        self.file_entry = tk.Text(
            file_frame, 
            font=("微软雅黑", 10),
            width=50,
            height=3,
            wrap=tk.WORD
        )
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10), pady=5)
        
        # 文件选择按钮框架
        file_btn_frame = tk.Frame(file_frame, bg="#f0f2f5")
        file_btn_frame.pack(side=tk.LEFT, padx=5, pady=5)
        
        browse_btn = ttk.Button(
            file_btn_frame, 
            text="选择文件", 
            command=self.browse_file
        )
        browse_btn.pack(side=tk.TOP, padx=5, pady=2)
        
        browse_folder_btn = ttk.Button(
            file_btn_frame, 
            text="选择文件夹", 
            command=self.browse_folder
        )
        browse_folder_btn.pack(side=tk.TOP, padx=5, pady=2)
        
        # 批量模式选项
        batch_frame = tk.Frame(self.root, bg="#f0f2f5")
        batch_frame.pack(pady=5, fill=tk.X, padx=20)
        
        ttk.Checkbutton(
            batch_frame,
            text="批量转换模式",
            variable=self.batch_mode
        ).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 输出路径选择区域
        output_frame = tk.Frame(self.root, bg="#f0f2f5")
        output_frame.pack(pady=10, fill=tk.X, padx=20)
        
        ttk.Label(output_frame, text="保存路径:").pack(side=tk.LEFT, padx=5, pady=5)
        
        self.output_entry = ttk.Entry(
            output_frame, 
            font=("微软雅黑", 10),
            width=50
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10), pady=5)
        self.output_entry.insert(0, self.output_dir)
        
        browse_output_btn = ttk.Button(
            output_frame, 
            text="选择路径", 
            command=self.browse_output_dir
        )
        browse_output_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 图片转换额外选项
        self.image_options_frame = tk.Frame(self.root, bg="#f0f2f5")
        # 默认隐藏，选择图片转换时显示
        
        ttk.Label(self.image_options_frame, text="图片质量 (1-100):").pack(side=tk.LEFT, padx=5, pady=5)
        
        self.image_quality = tk.IntVar(value=95)
        self.quality_scale = ttk.Scale(
            self.image_options_frame,
            from_=1,
            to=100,
            variable=self.image_quality,
            orient="horizontal",
            length=200
        )
        self.quality_scale.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.quality_label = ttk.Label(
            self.image_options_frame,
            text=str(self.image_quality.get())
        )
        self.quality_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 绑定滑块事件，实时显示质量值
        self.image_quality.trace_add("write", self.update_quality_label)
        
        # 按钮区域
        btn_frame = tk.Frame(self.root, bg="#f0f2f5")
        btn_frame.pack(pady=15)
        
        self.convert_btn = ttk.Button(
            btn_frame, 
            text="开始转换", 
            command=self.start_conversion
        )
        self.convert_btn.pack(side=tk.LEFT, padx=10)
        
        open_folder_btn = ttk.Button(
            btn_frame, 
            text="打开输出文件夹", 
            command=self.open_output_folder
        )
        open_folder_btn.pack(side=tk.LEFT, padx=10)
        
        exit_btn = ttk.Button(
            btn_frame, 
            text="退出", 
            command=self.root.destroy
        )
        exit_btn.pack(side=tk.LEFT, padx=10)
        
        # 进度条
        self.progress_frame = tk.Frame(self.root, bg="#f0f2f5")
        self.progress_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=100,
            mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # 批量转换进度显示
        self.batch_progress_label = ttk.Label(
            self.progress_frame,
            text="",
            background="#f0f2f5"
        )
        self.batch_progress_label.pack(pady=5)
        
        # 状态区域
        status_frame = tk.Frame(self.root, bg="#f0f2f5")
        status_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        ttk.Label(status_frame, text="转换日志:").pack(anchor=tk.W, pady=5)
        
        self.status_text = tk.Text(
            status_frame,
            font=("微软雅黑", 10),
            wrap=tk.WORD,
            height=10,
            state=tk.DISABLED
        )
        self.status_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.status_text, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # 底部信息
        footer_frame = tk.Frame(self.root, bg="#f0f2f5")
        footer_frame.pack(side=tk.BOTTOM, pady=10)
        
        footer_label = tk.Label(
            footer_frame,
            text="支持多种格式转换，可自定义输出文件保存路径",
            font=("微软雅黑", 9),
            bg="#f0f2f5",
            fg="#666666"
        )
        footer_label.pack()
        
        # 绑定转换类型变化事件
        self.conversion_type.trace_add("write", self.update_format_options)
        self.target_format.trace_add("write", self.update_special_options)
        self.update_format_options()
        self.update_special_options()
    
    def update_quality_label(self, *args):
        """更新图片质量显示"""
        self.quality_label.config(text=str(self.image_quality.get()))
    
    def update_format_options(self, *args):
        """根据转换类型更新格式选项"""
        conv_type = self.conversion_type.get()
        
        # 隐藏所有特殊选项框架
        self.format_frame.pack_forget()
        self.image_options_frame.pack_forget()
        self.ico_options_frame.pack_forget()
        self.excel_options_frame.pack_forget()
        
        if conv_type == "audio_convert":
            self.format_options['values'] = ["mp3", "wav", "flac", "m4a", "ogg", "aac"]
            self.target_format.set("mp3")
            self.format_frame.pack(pady=5, fill=tk.X, padx=20)
        elif conv_type == "video_convert":
            self.format_options['values'] = ["mp4", "avi", "mov", "mkv", "flv", "wmv"]
            self.target_format.set("mp4")
            self.format_frame.pack(pady=5, fill=tk.X, padx=20)
        elif conv_type == "image_convert":
            self.format_options['values'] = ["jpg", "jpeg", "png", "bmp", "gif", "tiff", "webp", "ico"]
            self.target_format.set("jpg")
            self.format_frame.pack(pady=5, fill=tk.X, padx=20)
            self.update_special_options()
        elif conv_type == "excel_to_pdf":
            # 显示Excel转换选项
            self.excel_options_frame.pack(pady=5, fill=tk.X, padx=20)
    
    def update_special_options(self, *args):
        """根据目标格式显示特殊选项（图片质量或ICO尺寸）"""
        if self.conversion_type.get() != "image_convert":
            return
            
        target_format = self.target_format.get().lower()
        
        if target_format == "ico":
            self.ico_options_frame.pack(pady=5, fill=tk.X, padx=20)
            self.image_options_frame.pack_forget()
        else:
            self.image_options_frame.pack(pady=5, fill=tk.X, padx=20)
            self.ico_options_frame.pack_forget()
    
    def browse_file(self):
        """浏览并选择文件（支持多选）"""
        conv_type = self.conversion_type.get()
        
        # 根据转换类型设置文件筛选器
        if conv_type == "pdf_to_word":
            filetypes = [("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        elif conv_type == "word_to_pdf":
            filetypes = [("Word文件", "*.docx;*.doc"), ("所有文件", "*.*")]
        elif conv_type == "excel_to_pdf":
            filetypes = [("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        elif conv_type == "ppt_to_pdf":
            filetypes = [("PPT文件", "*.pptx;*.ppt"), ("所有文件", "*.*")]
        elif conv_type == "audio_convert":
            filetypes = [("音频文件", "*.mp3;*.wav;*.flac;*.m4a;*.ogg;*.aac"), ("所有文件", "*.*")]
        elif conv_type == "video_convert":
            filetypes = [("视频文件", "*.mp4;*.avi;*.mov;*.mkv;*.flv;*.wmv"), ("所有文件", "*.*")]
        elif conv_type == "image_convert":
            filetypes = [("图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.webp;*.ico"), ("所有文件", "*.*")]
        
        file_paths = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=filetypes
        )
        
        if file_paths:
            self.file_paths = list(file_paths)
            self.update_file_list_display()
            self.update_status(f"已选择 {len(self.file_paths)} 个文件")
    
    def browse_folder(self):
        """浏览并选择文件夹，自动添加文件夹中所有支持的文件"""
        folder_path = filedialog.askdirectory(title="选择文件夹")
        if folder_path:
            conv_type = self.conversion_type.get()
            
            # 根据转换类型获取文件扩展名
            extensions = self.get_supported_extensions(conv_type)
            
            # 搜索文件夹中所有支持的文件
            file_paths = []
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if any(file.lower().endswith(ext) for ext in extensions):
                        file_paths.append(os.path.join(root, file))
            
            if file_paths:
                self.file_paths = file_paths
                self.update_file_list_display()
                self.update_status(f"从文件夹 '{folder_path}' 中找到 {len(self.file_paths)} 个支持的文件")
            else:
                messagebox.showwarning("警告", f"在文件夹 '{folder_path}' 中未找到支持的文件")
    
    def get_supported_extensions(self, conv_type):
        """根据转换类型返回支持的文件扩展名"""
        if conv_type == "pdf_to_word":
            return [".pdf"]
        elif conv_type == "word_to_pdf":
            return [".doc", ".docx"]
        elif conv_type == "excel_to_pdf":
            return [".xls", ".xlsx"]
        elif conv_type == "ppt_to_pdf":
            return [".ppt", ".pptx"]
        elif conv_type == "audio_convert":
            return [".mp3", ".wav", ".flac", ".m4a", ".ogg", ".aac"]
        elif conv_type == "video_convert":
            return [".mp4", ".avi", ".mov", ".mkv", ".flv", ".wmv"]
        elif conv_type == "image_convert":
            return [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp", ".ico"]
        return []
    
    def update_file_list_display(self):
        """更新文件列表显示"""
        self.file_entry.delete(1.0, tk.END)
        if self.file_paths:
            file_list = "\n".join([os.path.basename(path) for path in self.file_paths])
            self.file_entry.insert(1.0, file_list)
    
    def browse_output_dir(self):
        """选择输出文件保存路径"""
        selected_dir = filedialog.askdirectory(title="选择保存路径")
        if selected_dir:
            self.output_dir = selected_dir
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, selected_dir)
            self.update_status(f"已选择保存路径: {selected_dir}")
    
    def update_status(self, message):
        """更新状态文本区域"""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def update_batch_progress(self, current, total, filename):
        """更新批量转换进度显示"""
        progress_text = f"批量转换进度: {current}/{total} - 当前文件: {filename}"
        self.batch_progress_label.config(text=progress_text)
    
    def open_output_folder(self):
        """打开输出文件夹"""
        try:
            os.startfile(self.output_dir)
            self.update_status(f"已打开输出文件夹: {self.output_dir}")
        except Exception as e:
            self.update_status(f"打开文件夹失败: {str(e)}")
            messagebox.showerror("错误", f"打开文件夹失败: {str(e)}")
    
    def run_ffmpeg_silently(self, input_file, output_file, output_format):
        """静默运行ffmpeg，不显示命令行窗口"""
        try:
            # 构建ffmpeg命令
            cmd = [
                self.ffmpeg_path,
                '-i', input_file,
                '-y',  # 覆盖输出文件
                output_file
            ]
            
            # Windows系统下隐藏命令行窗口
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0  # 隐藏窗口
                
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NO_WINDOW  # 不创建窗口
                )
            else:
                # 非Windows系统
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
            
            # 等待进程完成
            stdout, stderr = process.communicate()
            
            if process.returncode != 0:
                error_msg = stderr.decode('utf-8', errors='ignore') if stderr else "未知错误"
                raise Exception(f"FFmpeg转换失败: {error_msg}")
                
            return True
            
        except Exception as e:
            raise Exception(f"FFmpeg执行错误: {str(e)}")
    
    def pdf_to_word(self, file_path):
        """PDF转Word"""
        try:
            self.update_status(f"开始PDF转Word: {os.path.basename(file_path)}")
            self.update_progress(10)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self.output_dir, f"{file_name}.docx")
            
            self.update_status(f"正在转换: {os.path.basename(file_path)}")
            self.update_progress(30)
            
            p2w = Converter(file_path)
            self.update_progress(50)
            
            p2w.convert(output_file, start=0, end=None)
            self.update_progress(80)
            
            p2w.close()
            self.update_progress(100)
            
            return output_file
            
        except Exception as e:
            raise Exception(f"PDF转Word失败: {str(e)}")
    
    def word_to_pdf(self, file_path):
        """Word转PDF - 包含错误处理和备选方案"""
        try:
            if not docx2pdf_available:
                raise Exception("docx2pdf库未安装，请先运行 'pip install docx2pdf'")
                
            self.update_status(f"开始Word转PDF: {os.path.basename(file_path)}")
            self.update_progress(30)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self.output_dir, f"{file_name}.pdf")
            
            self.update_status(f"正在转换: {os.path.basename(file_path)}")
            
            # 使用docx2pdf库的convert函数
            convert(file_path, output_file)
            self.update_progress(100)
            
            return output_file
            
        except Exception as e:
            # 如果docx2pdf失败，尝试使用Word的COM接口作为备选方案
            try:
                self.update_status(f"尝试备选方案转换: {os.path.basename(file_path)}")
                
                # 使用Word的COM接口
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(file_path)
                doc.SaveAs(output_file, FileFormat=17)  # 17是PDF格式
                doc.Close()
                word.Quit()
                
                self.update_progress(100)
                return output_file
                
            except Exception as e2:
                raise Exception(f"Word转PDF失败: 主方案[{str(e)}], 备选方案[{str(e2)}]")
    
    def excel_to_pdf(self, file_path):
        """Excel转PDF - 修复表格显示不全问题"""
        try:
            self.update_status(f"开始Excel转PDF: {os.path.basename(file_path)}")
            self.update_progress(20)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self.output_dir, f"{file_name}.pdf")
            
            self.update_status(f"正在转换: {os.path.basename(file_path)}")
            self.update_progress(40)
            
            # 创建Excel应用实例
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False  # 禁用警告
            
            # 打开工作簿
            workbook = excel.Workbooks.Open(file_path)
            
            # 遍历所有工作表并设置页面
            for worksheet in workbook.Worksheets:
                # 选择当前工作表
                worksheet.Select()
                
                # 设置页面方向
                if self.excel_orientation.get() == "landscape":
                    worksheet.PageSetup.Orientation = 2  # 2 表示横向
                else:
                    worksheet.PageSetup.Orientation = 1  # 1 表示纵向
                
                # 设置自动调整到一页
                if self.excel_fit_to_page.get():
                    worksheet.PageSetup.Zoom = False  # 禁用缩放
                    worksheet.PageSetup.FitToPagesWide = 1  # 宽度适应1页
                    worksheet.PageSetup.FitToPagesTall = False  # 高度自动
                else:
                    worksheet.PageSetup.Zoom = 100  # 使用100%缩放
            
            self.update_progress(60)
            
            # 导出为PDF
            workbook.ExportAsFixedFormat(0, output_file)  # 0 表示PDF格式
            
            # 清理资源
            workbook.Close(SaveChanges=False)  # 不保存对原文件的修改
            excel.Quit()
            
            # 释放COM对象
            del workbook
            del excel
            
            self.update_progress(100)
            return output_file
            
        except Exception as e:
            raise Exception(f"Excel转PDF失败: {str(e)}")
    
    def ppt_to_pdf(self, file_path):
        """PPT转PDF"""
        try:
            self.update_status(f"开始PPT转PDF: {os.path.basename(file_path)}")
            self.update_progress(20)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self.output_dir, f"{file_name}.pdf")
            
            self.update_status(f"正在转换: {os.path.basename(file_path)}")
            self.update_progress(40)
            
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(file_path)
            self.update_progress(60)
            
            presentation.SaveAs(output_file, 32)  # 32 表示PDF格式
            presentation.Close()
            powerpoint.Quit()
            
            self.update_progress(100)
            return output_file
            
        except Exception as e:
            raise Exception(f"PPT转PDF失败: {str(e)}")
    
    def audio_convert(self, file_path):
        """音频格式转换"""
        try:
            # 检查ffmpeg是否存在
            if not self.ffmpeg_path or not os.path.exists(self.ffmpeg_path):
                raise Exception(f"无法找到ffmpeg.exe，音频转换功能无法使用")
                
            self.update_status(f"开始音频格式转换: {os.path.basename(file_path)}")
            self.update_progress(20)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_format = self.target_format.get()
            output_file = os.path.join(self.output_dir, f"{file_name}.{output_format}")
            
            self.update_status(f"正在转换为{output_format}: {os.path.basename(file_path)}")
            self.update_progress(40)
            
            # 使用静默方式运行ffmpeg
            self.run_ffmpeg_silently(file_path, output_file, output_format)
            
            self.update_progress(100)
            return output_file
            
        except Exception as e:
            raise Exception(f"音频转换失败: {str(e)}")
    
    def video_convert(self, file_path):
        """视频格式转换"""
        try:
            # 检查ffmpeg是否存在
            if not self.ffmpeg_path or not os.path.exists(self.ffmpeg_path):
                raise Exception(f"无法找到ffmpeg.exe，视频转换功能无法使用")
                
            self.update_status(f"开始视频格式转换: {os.path.basename(file_path)}")
            self.update_progress(20)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_format = self.target_format.get()
            output_file = os.path.join(self.output_dir, f"{file_name}.{output_format}")
            
            self.update_status(f"正在转换为{output_format}: {os.path.basename(file_path)}")
            self.update_progress(40)
            
            # 使用静默方式运行ffmpeg
            self.run_ffmpeg_silently(file_path, output_file, output_format)
            
            self.update_progress(100)
            return output_file
            
        except Exception as e:
            raise Exception(f"视频转换失败: {str(e)}")
    
    def image_convert(self, file_path):
        """图片格式转换，包含ICO转换和JPG转换修复"""
        try:
            # 检查PIL是否可用
            if not self.pil_available:
                raise Exception("Pillow库未安装，请先运行 'pip install pillow'")
                
            self.update_status(f"开始图片格式转换: {os.path.basename(file_path)}")
            self.update_progress(20)
            
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_format = self.target_format.get().upper()
            output_file = os.path.join(self.output_dir, f"{file_name}.{output_format.lower()}")
            
            self.update_status(f"正在转换为{output_format}: {os.path.basename(file_path)}")
            self.update_progress(40)
            
            # 打开图片
            try:
                with Image.open(file_path) as img:
                    # 处理ICO格式
                    if output_format == "ICO":
                        # 获取用户选择的尺寸
                        selected_sizes = [self.ico_sizes[i] for i, var in enumerate(self.selected_sizes) if var.get()]
                        if not selected_sizes:
                            raise Exception("请至少选择一个ICO图标尺寸")
                            
                        # 保存多尺寸ICO
                        img.save(output_file, sizes=selected_sizes)
                    
                    else:
                        # 处理透明通道问题（针对JPG等不支持透明的格式）
                        if output_format in ["JPG", "JPEG", "BMP"] and img.mode in ["RGBA", "LA", "P"]:
                            # 对于带透明通道的图片，创建白色背景
                            if img.mode == "P":
                                # 处理调色板图像
                                img = img.convert("RGBA")
                                
                            background = Image.new("RGB", img.size, (255, 255, 255))
                            # 处理alpha通道
                            background.paste(img, mask=img.split()[-1])
                            img = background
                        elif output_format in ["JPG", "JPEG"] and img.mode in ["CMYK"]:
                            # 处理CMYK模式图片转JPG
                            img = img.convert('RGB')
                    
                        # 获取用户设置的质量值
                        quality = self.image_quality.get()
                        
                        # 保存为目标格式
                        if output_format in ["JPG", "JPEG"]:
                            # 确保图片是RGB模式
                            if img.mode != 'RGB':
                                img = img.convert('RGB')
                            img.save(output_file, output_format, quality=quality, optimize=True, progressive=True)
                        elif output_format == "PNG":
                            # PNG格式使用压缩级别参数
                            compress_level = 9 - int(quality / 11)  # 将1-100转换为0-9
                            img.save(output_file, output_format, compress_level=compress_level)
                        elif output_format == "GIF":
                            # 处理GIF动画
                            if img.is_animated:
                                frames = []
                                for frame in range(img.n_frames):
                                    img.seek(frame)
                                    frames.append(img.copy())
                                frames[0].save(output_file, format=output_format, save_all=True, append_images=frames[1:], loop=0)
                            else:
                                img.save(output_file, output_format)
                        else:
                            img.save(output_file, output_format)
            
            except Exception as e:
                raise Exception(f"图片处理错误: {str(e)}")
            
            self.update_progress(100)
            return output_file
            
        except Exception as e:
            raise Exception(f"图片转换失败: {str(e)}")
    
    def start_conversion(self):
        """开始转换过程（在新线程中执行）"""
        if not self.file_paths:
            messagebox.showwarning("警告", "请选择至少一个文件")
            return
        
        # 更新输出目录为用户选择的路径
        self.output_dir = self.output_entry.get()
        if not self.output_dir:
            self.output_dir = os.path.expanduser("~/转换输出")
            self.output_entry.insert(0, self.output_dir)
        os.makedirs(self.output_dir, exist_ok=True)
        
        self.convert_btn.config(state=tk.DISABLED)
        self.update_progress(0)
        
        conversion_thread = threading.Thread(
            target=self.perform_conversion
        )
        conversion_thread.daemon = True
        conversion_thread.start()
    
    def perform_conversion(self):
        """执行转换"""
        try:
            conv_type = self.conversion_type.get()
            successful_conversions = 0
            failed_conversions = 0
            
            self.total_files = len(self.file_paths)
            
            for i, file_path in enumerate(self.file_paths):
                self.current_file_index = i + 1
                
                # 更新批量转换进度显示
                self.update_batch_progress(self.current_file_index, self.total_files, os.path.basename(file_path))
                
                try:
                    output_file = ""
                    
                    if conv_type == "pdf_to_word":
                        output_file = self.pdf_to_word(file_path)
                    elif conv_type == "word_to_pdf":
                        output_file = self.word_to_pdf(file_path)
                    elif conv_type == "excel_to_pdf":
                        output_file = self.excel_to_pdf(file_path)
                    elif conv_type == "ppt_to_pdf":
                        output_file = self.ppt_to_pdf(file_path)
                    elif conv_type == "audio_convert":
                        output_file = self.audio_convert(file_path)
                    elif conv_type == "video_convert":
                        output_file = self.video_convert(file_path)
                    elif conv_type == "image_convert":
                        output_file = self.image_convert(file_path)
                    
                    self.update_status(f"✓ 转换成功: {os.path.basename(file_path)} -> {os.path.basename(output_file)}")
                    successful_conversions += 1
                    
                except Exception as e:
                    self.update_status(f"✗ 转换失败: {os.path.basename(file_path)} - {str(e)}")
                    failed_conversions += 1
                
                # 更新总体进度
                overall_progress = (self.current_file_index / self.total_files) * 100
                self.update_progress(overall_progress)
            
            # 清空批量进度显示
            self.batch_progress_label.config(text="")
            
            # 显示转换结果摘要
            summary = f"批量转换完成！成功: {successful_conversions} 个，失败: {failed_conversions} 个"
            self.update_status(summary)
            
            if failed_conversions == 0:
                messagebox.showinfo("成功", f"所有文件转换完成！\n成功转换 {successful_conversions} 个文件")
            else:
                messagebox.showwarning("完成", 
                    f"批量转换完成！\n"
                    f"成功: {successful_conversions} 个文件\n"
                    f"失败: {failed_conversions} 个文件\n"
                    f"请查看日志了解失败详情")
            
            self.convert_btn.config(state=tk.NORMAL)
            
        except Exception as e:
            self.update_status(f"批量转换过程出错: {str(e)}")
            messagebox.showerror("错误", f"批量转换过程出错: {str(e)}")
            self.convert_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = FormatConverter(root)
    root.mainloop()
