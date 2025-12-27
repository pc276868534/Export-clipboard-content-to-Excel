import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
import pyperclip
import re
import pandas as pd
from datetime import datetime
import os
import json
from queue import Queue

class WeChatClipboardMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("微信统计信息监控器")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 状态变量
        self.monitoring = False
        self.auto_refresh_enabled = True  # 自动刷新状态
        self.clipboard_history = []
        self.processed_data = []
        self.last_clipboard = ""
        self.processed_signatures = set()  # 用于检测重复数据
        self.existing_excel_data = []  # 存储Excel中已存在的数据
        
        # 配置
        self.excel_file = "微信统计记录.xlsx"
        self.config_file = "monitor_config.json"
        self.load_config()
        
        # 界面样式
        self.setup_ui_styles()
        
        # 消息队列
        self.message_queue = Queue()
        
        # 创建界面
        self.create_ui()
        
        # 启动消息处理循环
        self.process_messages()
        
        # 初始刷新数据显示
        self.root.after(100, self.refresh_data)
        
        # 如果Excel文件存在，自动加载
        if os.path.exists(self.excel_file):
            self.root.after(500, self.load_existing_excel)
        
        # 安排自动刷新
        self.root.after(1000, self.schedule_auto_refresh)
        
    def load_config(self):
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.excel_file = config.get('excel_file', '微信统计记录.xlsx')
                    self.monitor_keywords = config.get('monitor_keywords', ['统计'])
                    self.monitor_mode = config.get('monitor_mode', 'startswith')
                    self.case_sensitive = config.get('case_sensitive', False)
                    self.ui_theme = config.get('ui_theme', 'modern')
                    self.auto_refresh_interval = config.get('auto_refresh_interval', 5000)
            else:
                # 默认配置
                self.monitor_keywords = ['统计']
                self.monitor_mode = 'startswith'
                self.case_sensitive = False
                self.ui_theme = 'modern'
                self.auto_refresh_interval = 5000
        except:
            # 默认配置
            self.monitor_keywords = ['统计']
            self.monitor_mode = 'startswith'
            self.case_sensitive = False
            self.ui_theme = 'modern'
            self.auto_refresh_interval = 5000
    
    def save_config(self):
        """保存配置"""
        try:
            config = {
                'excel_file': self.excel_file,
                'monitor_keywords': self.monitor_keywords,
                'monitor_mode': self.monitor_mode,
                'case_sensitive': self.case_sensitive,
                'ui_theme': self.ui_theme,
                'auto_refresh_interval': self.auto_refresh_interval,
                'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def setup_ui_styles(self):
        """设置界面样式"""
        try:
            style = ttk.Style()
            
            # 根据配置的主题设置样式
            if self.ui_theme == 'modern':
                # 现代风格
                style.theme_use('clam')
                style.configure('TFrame', background='#f8f9fa')
                style.configure('TNotebook', background='#f8f9fa', borderwidth=1)
                style.configure('TNotebook.Tab', padding=[20, 12], background='#e9ecef', foreground='#495057')
                style.map('TNotebook.Tab', background=[('selected', '#ffffff'), ('active', '#f8f9fa')])
                style.configure('TButton', padding=[15, 8], font=('Microsoft YaHei', 9))
                style.map('TButton', background=[('active', '#007bff'), ('pressed', '#0056b3')])
                style.configure('TLabel', background='#f8f9fa', foreground='#495057', font=('Microsoft YaHei', 9))
                style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff')
                style.configure('Treeview.Heading', background='#e9ecef', font=('Microsoft YaHei', 9, 'bold'))
                
            elif self.ui_theme == 'classic':
                # 经典风格
                style.theme_use('default')
                style.configure('TButton', padding=[12, 6])
                style.configure('TNotebook.Tab', padding=[15, 10])
                
            elif self.ui_theme == 'minimal':
                # 简洁风格
                style.theme_use('alt')
                style.configure('TButton', padding=[10, 5], relief='flat')
                style.configure('TNotebook', borderwidth=0)
                style.configure('TNotebook.Tab', padding=[18, 10], borderwidth=0)
                style.configure('Treeview', borderwidth=1)
                
        except Exception as e:
            print(f"设置样式失败: {e}")
            # 使用默认样式
            pass
    
    def center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_ui(self):
        """创建用户界面"""
        # 设置窗口标题和图标
        self.root.title("📊 微信统计信息监控器 v3.0")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        # 设置窗口居中
        self.center_window()
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 标题区域
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        title_label = ttk.Label(title_frame, text="📊 微信统计信息监控器", 
                               font=('Microsoft YaHei', 16, 'bold'))
        title_label.pack(side=tk.LEFT)
        
        subtitle_label = ttk.Label(title_frame, text="智能剪贴板监控 · Excel实时对比 · 自定义格式匹配", 
                                 font=('Microsoft YaHei', 10), foreground='#6c757d')
        subtitle_label.pack(side=tk.LEFT, padx=(20, 0))
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 控制面板
        self.create_control_panel(main_frame)
        
        # 状态面板
        self.create_status_panel(main_frame)
        
        # 数据显示区域
        self.create_data_panel(main_frame)
        
        # 按钮面板
        self.create_button_panel(main_frame)
    
    def create_control_panel(self, parent):
        """创建控制面板"""
        control_frame = ttk.LabelFrame(parent, text="🎛️ 控制面板", padding="12")
        control_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(1, weight=1)
        
        # Excel文件选择区域
        excel_section = ttk.Frame(control_frame)
        excel_section.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        excel_section.columnconfigure(1, weight=1)
        
        ttk.Label(excel_section, text="📁 Excel文件:", font=('Microsoft YaHei', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        excel_frame = ttk.Frame(excel_section)
        excel_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        excel_frame.columnconfigure(0, weight=1)
        
        self.excel_var = tk.StringVar(value=self.excel_file)
        self.excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_var, width=60, font=('Microsoft YaHei', 9))
        self.excel_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 8))
        
        # 使用更美观的按钮
        browse_btn = ttk.Button(excel_frame, text="📂 浏览", command=self.browse_excel)
        browse_btn.grid(row=0, column=1, padx=(0, 5))
        
        load_btn = ttk.Button(excel_frame, text="📥 读取Excel", command=self.load_existing_excel)
        load_btn.grid(row=0, column=2)
        
        # 监控控制区域
        monitor_section = ttk.Frame(control_frame)
        monitor_section.grid(row=1, column=0, columnspan=2, pady=(10, 5))
        
        ttk.Label(monitor_section, text="📡 监控控制:", font=('Microsoft YaHei', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 15))
        # 创建监控按钮组
        self.start_btn = ttk.Button(monitor_section, text="▶️ 开始监控", command=self.start_monitoring)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        self.stop_btn = ttk.Button(monitor_section, text="⏹️ 停止监控", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        ttk.Button(monitor_section, text="🗑️ 清空历史", command=self.clear_history).pack(side=tk.LEFT, padx=(0, 15))
        
        # 统计信息显示
        ttk.Label(monitor_section, text="📊", font=('Microsoft YaHei', 12)).pack(side=tk.LEFT)
        self.stats_label = ttk.Label(monitor_section, text="新数据: 0 | Excel数据: 0 | 总计: 0", 
                                   font=('Microsoft YaHei', 9, 'bold'))
        self.stats_label.pack(side=tk.LEFT, padx=(5, 0))
    
    def create_status_panel(self, parent):
        """创建状态面板"""
        status_frame = ttk.LabelFrame(parent, text="📡 监控状态", padding="12")
        status_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(1, weight=1)
        
        # 状态指示器
        status_container = ttk.Frame(status_frame)
        status_container.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(status_container, text="🔴", font=('Microsoft YaHei', 12)).pack(side=tk.LEFT, padx=(0, 8))
        self.status_var = tk.StringVar(value="监控未启动")
        self.status_label = ttk.Label(status_container, textvariable=self.status_var, 
                                   font=("Microsoft YaHei", 10, "bold"))
        self.status_label.pack(side=tk.LEFT)
        
        # 最后更新时间
        update_container = ttk.Frame(status_frame)
        update_container.pack(fill=tk.X)
        
        ttk.Label(update_container, text="🕒", font=('Microsoft YaHei', 10)).pack(side=tk.LEFT, padx=(0, 8))
        self.last_update_var = tk.StringVar(value="最后更新: --")
        ttk.Label(status_frame, textvariable=self.last_update_var).pack(anchor=tk.W)
        
        # 进度条
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)
    
    def create_data_panel(self, parent):
        """创建数据显示面板"""
        # 创建Notebook
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Excel实时数据标签页
        excel_frame = ttk.Frame(self.notebook)
        self.notebook.add(excel_frame, text="📊 Excel实时数据")
        
        # Excel数据显示区域
        excel_tree_frame = ttk.Frame(excel_frame)
        excel_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        excel_columns = ('序号', '名字', '位置', '租金', '租期', '押金方式', '时间', '备注', '来源')
        self.excel_tree = ttk.Treeview(excel_tree_frame, columns=excel_columns, show='headings', height=12)
        
        # 设置Excel表格列标题和宽度
        for col in excel_columns:
            self.excel_tree.heading(col, text=col)
            if col == '序号':
                self.excel_tree.column(col, width=50, anchor=tk.CENTER)
            elif col == '租金':
                self.excel_tree.column(col, width=80, anchor=tk.CENTER)
            elif col == '时间':
                self.excel_tree.column(col, width=120, anchor=tk.CENTER)
            elif col == '备注':
                self.excel_tree.column(col, width=120, anchor=tk.CENTER)
            elif col == '来源':
                self.excel_tree.column(col, width=80, anchor=tk.CENTER)
            else:
                self.excel_tree.column(col, width=100, anchor=tk.W)
        
        # 添加Excel表格滚动条
        excel_v_scrollbar = ttk.Scrollbar(excel_tree_frame, orient=tk.VERTICAL, command=self.excel_tree.yview)
        excel_h_scrollbar = ttk.Scrollbar(excel_tree_frame, orient=tk.HORIZONTAL, command=self.excel_tree.xview)
        self.excel_tree.configure(yscrollcommand=excel_v_scrollbar.set, xscrollcommand=excel_h_scrollbar.set)
        
        self.excel_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        excel_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        excel_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        excel_tree_frame.columnconfigure(0, weight=1)
        excel_tree_frame.rowconfigure(0, weight=1)
        
        # Excel控制按钮
        excel_button_frame = ttk.Frame(excel_frame)
        excel_button_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(excel_button_frame, text="刷新Excel数据", command=self.refresh_excel_display).pack(side=tk.LEFT, padx=5)
        self.auto_refresh_btn = ttk.Button(excel_button_frame, text="自动刷新: 开启", command=self.toggle_auto_refresh)
        self.auto_refresh_btn.pack(side=tk.LEFT, padx=5)
        self.excel_status_label = ttk.Label(excel_button_frame, text="Excel记录: 0")
        self.excel_status_label.pack(side=tk.RIGHT, padx=5)
        
        # 剪贴板历史标签页
        history_frame = ttk.Frame(self.notebook)
        self.notebook.add(history_frame, text="📝 剪贴板历史")
        
        self.history_text = scrolledtext.ScrolledText(history_frame, height=15, width=80)
        self.history_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 新数据预览标签页
        result_frame = ttk.Frame(self.notebook)
        self.notebook.add(result_frame, text="🆕 新数据预览")
        
        # 创建Treeview
        tree_frame = ttk.Frame(result_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = ('序号', '名字', '位置', '租金', '租期', '押金方式', '时间', '备注')
        self.result_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12)
        
        # 设置列标题和宽度
        for col in columns:
            self.result_tree.heading(col, text=col)
            if col == '序号':
                self.result_tree.column(col, width=50, anchor=tk.CENTER)
            elif col == '租金':
                self.result_tree.column(col, width=80, anchor=tk.CENTER)
            elif col == '时间':
                self.result_tree.column(col, width=120, anchor=tk.CENTER)
            elif col == '备注':
                self.result_tree.column(col, width=120, anchor=tk.CENTER)
            else:
                self.result_tree.column(col, width=100, anchor=tk.W)
        
        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.result_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # 新数据预览按钮面板
        result_button_frame = ttk.Frame(result_frame)
        result_button_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        ttk.Button(result_button_frame, text="🗑️ 删除选中记录", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(result_button_frame, text="📥 导出到Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        
        # 新数据统计标签
        self.new_data_count_label = ttk.Label(result_button_frame, text="新数据: 0", 
                                         font=('Microsoft YaHei', 9, 'bold'))
        self.new_data_count_label.pack(side=tk.RIGHT)
    
    def create_button_panel(self, parent):
        """创建按钮面板"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        ttk.Button(button_frame, text="打开Excel文件", command=self.open_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新数据", command=self.refresh_data).pack(side=tk.LEFT, padx=5)
        
        # 关于按钮
        ttk.Button(button_frame, text="⚙️ 配置", command=self.open_config_dialog).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="使用说明", command=self.show_help).pack(side=tk.RIGHT, padx=5)
    
    def browse_excel(self):
        """浏览选择Excel文件"""
        from tkinter import filedialog
        filename = filedialog.asksaveasfilename(
            title="选择Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_var.set(filename)
            self.excel_file = filename
            self.save_config()
    
    def load_existing_excel(self):
        """读取现有的Excel文件内容"""
        if not os.path.exists(self.excel_file):
            messagebox.showinfo("文件不存在", f"Excel文件不存在:\n{self.excel_file}")
            return
        
        try:
            # 读取Excel文件
            df = pd.read_excel(self.excel_file, engine='openpyxl')
            
            if df.empty:
                messagebox.showinfo("空文件", "Excel文件中没有数据")
                return
            
            # 处理数据
            self.existing_excel_data = []
            excel_signatures = set()
            duplicate_count = 0
            
            for index, row in df.iterrows():
                # 转换为字典
                record = {
                    '序号': index + 1,
                    '名字': str(row.get('名字', '')) if pd.notna(row.get('名字')) else '',
                    '位置': str(row.get('位置', '')) if pd.notna(row.get('位置')) else '',
                    '租金': str(row.get('租金', '')) if pd.notna(row.get('租金')) else '',
                    '租期': str(row.get('租期', '')) if pd.notna(row.get('租期')) else '',
                    '押金方式': str(row.get('押金方式', '')) if pd.notna(row.get('押金方式')) else '',
                    '时间': str(row.get('时间', '')) if pd.notna(row.get('时间')) else '',
                    '备注': str(row.get('备注', '')) if pd.notna(row.get('备注')) else '',
                    '来源': 'Excel'
                }
                
                # 生成签名
                signature = self.get_data_signature(record)
                
                if signature in excel_signatures:
                    duplicate_count += 1
                    record['备注'] = 'Excel内重复记录'
                
                excel_signatures.add(signature)
                self.existing_excel_data.append(record)
            
            # 更新现有数据的签名集（用于后续去重）
            self.processed_signatures.update(excel_signatures)
            
            # 刷新Excel显示
            self.refresh_excel_display()
            
            # 显示结果
            messagebox.showinfo(
                "Excel读取完成", 
                f"成功读取Excel文件:\n{self.excel_file}\n\n"
                f"总记录数: {len(self.existing_excel_data)}\n"
                f"重复记录数: {duplicate_count}\n"
                f"有效记录数: {len(self.existing_excel_data) - duplicate_count}\n\n"
                f"已将Excel数据加载到实时预览窗口和重复检测！"
            )
            
            # 更新界面显示（预览窗口）
            self.refresh_data()
            
        except Exception as e:
            messagebox.showerror("读取失败", f"读取Excel文件失败:\n{str(e)}")
    
    def refresh_excel_display(self):
        """刷新Excel数据显示"""
        # 清空Excel树形控件
        for item in self.excel_tree.get_children():
            self.excel_tree.delete(item)
        
        # 重新添加Excel数据
        for data in self.existing_excel_data:
            self.excel_tree.insert('', tk.END, values=(
                data['序号'],
                data['名字'],
                data['位置'],
                data['租金'],
                data['租期'],
                data['押金方式'],
                data['时间'],
                data['备注'],
                data['来源']
            ))
        
        # 更新Excel状态标签
        self.excel_status_label.config(text=f"Excel记录: {len(self.existing_excel_data)}")
    
    def toggle_auto_refresh(self):
        """切换自动刷新状态"""
        self.auto_refresh_enabled = not self.auto_refresh_enabled
        if self.auto_refresh_enabled:
            self.schedule_auto_refresh()
            self.auto_refresh_btn.config(text="自动刷新: 开启")
        else:
            self.auto_refresh_btn.config(text="自动刷新: 关闭")
    
    def schedule_auto_refresh(self):
        """安排自动刷新"""
        if self.auto_refresh_enabled:
            # 自动读取最新的Excel文件内容
            if os.path.exists(self.excel_file):
                try:
                    # 读取最新Excel文件
                    df = pd.read_excel(self.excel_file, engine='openpyxl')
                    
                    if not df.empty:
                        # 重新处理数据
                        new_excel_data = []
                        for index, row in df.iterrows():
                            record = {
                                '序号': index + 1,
                                '名字': str(row.get('名字', '')) if pd.notna(row.get('名字')) else '',
                                '位置': str(row.get('位置', '')) if pd.notna(row.get('位置')) else '',
                                '租金': str(row.get('租金', '')) if pd.notna(row.get('租金')) else '',
                                '租期': str(row.get('租期', '')) if pd.notna(row.get('租期')) else '',
                                '押金方式': str(row.get('押金方式', '')) if pd.notna(row.get('押金方式')) else '',
                                '时间': str(row.get('时间', '')) if pd.notna(row.get('时间')) else '',
                                '备注': str(row.get('备注', '')) if pd.notna(row.get('备注')) else '',
                                '来源': 'Excel'
                            }
                            new_excel_data.append(record)
                        
                        # 检查是否有更新
                        if len(new_excel_data) != len(self.existing_excel_data):
                            self.existing_excel_data = new_excel_data
                            self.refresh_excel_display()
                            self.update_stats()
                except:
                    pass  # 忽略读取错误，避免影响用户体验
            
            # 安排下次刷新
            self.root.after(self.auto_refresh_interval, self.schedule_auto_refresh)
    
    def extract_statistics_from_text(self, text):
        """从文本中提取与"统计"同一行的内容"""
        if not text:
            return []
        
        results = []
        lines = text.split('\n')
        
        for i, line in enumerate(lines):
            line = line.strip()
            if self.is_statistics_message(line):
                # 解析统计消息
                parsed_data = self.parse_statistics_message(line)
                if parsed_data:
                    results.append({
                        'line_number': i + 1,
                        'line_content': line,
                        'parsed_data': parsed_data
                    })
        
        return results
    
    def monitor_clipboard_with_line_detection(self):
        """改进的剪贴板监控 - 行级别检测"""
        check_interval = 1.0
        
        while self.monitoring:
            try:
                current_clipboard = pyperclip.paste()
                
                if current_clipboard != self.last_clipboard:
                    self.last_clipboard = current_clipboard
                    
                    # 添加到历史记录
                    timestamp = datetime.now().strftime('%H:%M:%S')
                    self.clipboard_history.append({
                        'time': timestamp,
                        'content': current_clipboard[:100] + '...' if len(current_clipboard) > 100 else current_clipboard,
                        'full_content': current_clipboard
                    })
                    
                    # 限制历史记录数量
                    if len(self.clipboard_history) > 100:
                        self.clipboard_history = self.clipboard_history[-100:]
                    
                    # 提取统计行
                    statistics_lines = self.extract_statistics_from_text(current_clipboard.strip())
                    
                    # 如果有统计行，自动跳转到剪贴板历史窗口
                    if statistics_lines:
                        self.message_queue.put(('switch_to_clipboard_tab', None))
                    
                    # 处理每一条统计消息
                    for stat_info in statistics_lines:
                        parsed_data = stat_info['parsed_data']
                        line_content = stat_info['line_content']
                        line_number = stat_info['line_number']
                        
                        # 检查重复（包含Excel对比）
                        if self.is_duplicate_data(parsed_data):
                            # 查找Excel中的重复记录
                            excel_duplicate = self.find_duplicate_in_excel(parsed_data)
                            
                            # 发送重复提醒消息（包含Excel对比信息）
                            self.message_queue.put(('duplicate_found', {
                                **parsed_data,
                                '行号': line_number,
                                '行内容': line_content,
                                'excel_duplicate': excel_duplicate
                            }))
                        else:
                            # 添加到已处理数据
                            signature = self.get_data_signature(parsed_data)
                            self.processed_signatures.add(signature)
                            parsed_data['序号'] = len(self.existing_excel_data) + len(self.processed_data) + 1
                            parsed_data['来源'] = '剪贴板'
                            parsed_data['行号'] = line_number
                            parsed_data['行内容'] = line_content
                            self.processed_data.append(parsed_data)
                            
                            # 发送消息到队列
                            self.message_queue.put(('statistics_found', {
                                **parsed_data,
                                '来源': '剪贴板'
                            }))
                    
                    # 发送更新消息
                    if statistics_lines:
                        self.message_queue.put(('clipboard_updated', current_clipboard))
                
                time.sleep(check_interval)
                
            except Exception as e:
                print(f"监控剪贴板出错: {e}")
                time.sleep(check_interval)
    
    def is_statistics_message(self, text):
        """判断是否为统计消息（使用配置的关键词和匹配模式）"""
        if not text or len(text) < 1:
            return False
        
        # 清理文本首尾空白
        text = text.strip()
        
        # 使用配置的关键词进行匹配
        for keyword in self.monitor_keywords:
            if not keyword:
                continue
                
            keyword = keyword.strip()
            if not keyword:
                continue
                
            if self.monitor_mode == 'startswith':
                # 检查是否以关键词开头（支持特殊字符和字符+中文组合）
                try:
                    # 直接开头匹配
                    if text.startswith(keyword):
                        return True
                    
                    # 处理关键词后跟分隔符的情况
                    separators = ['，', ',', '：', ':', '、', ' ', '\t', '-', '—', '——']
                    for sep in separators:
                        if text.startswith(keyword + sep):
                            return True
                    
                    # 特殊处理：如果关键词包含特殊字符，使用更灵活的正则匹配
                    if any(char in keyword for char in ['*', '#', '@', '￥', '$', '€', '£', '¥', '%']):
                        # 构建正则表达式模式，支持关键词后跟任意分隔符
                        escaped_keyword = re.escape(keyword)
                        pattern = f"^{escaped_keyword}[，, ：:、\s\-———]*.*"
                        if re.match(pattern, text):
                            return True
                    
                    # 如果是纯中文关键词，也使用正则匹配
                    if re.match(r'^[\u4e00-\u9fff]+$', keyword):
                        escaped_keyword = re.escape(keyword)
                        pattern = f"^{escaped_keyword}[，, ：:、\s\-———]*.*"
                        if re.match(pattern, text):
                            return True
                
                except:
                    # 正则表达式出错时，使用简单匹配
                    if text.startswith(keyword):
                        return True
                    
            elif self.monitor_mode == 'contains':
                # 检查是否包含关键词
                if self.case_sensitive:
                    if keyword in text:
                        return True
                else:
                    if keyword.lower() in text.lower():
                        return True
                        
            elif self.monitor_mode == 'exact':
                # 检查是否精确匹配
                if self.case_sensitive:
                    if text == keyword:
                        return True
                else:
                    if text.lower() == keyword.lower():
                        return True
        
        return False
    
    def format_date_from_excel(self, date_value):
        """统一Excel中的日期格式为YYYY-MM-DD"""
        try:
            if pd.isna(date_value) or date_value == '':
                return ''
            
            # 如果是datetime类型，转换为日期格式
            if hasattr(date_value, 'date'):
                return date_value.strftime('%Y-%m-%d')
            
            # 如果是字符串，尝试解析并重新格式化
            if isinstance(date_value, str):
                # 尝试常见的日期格式
                date_formats = ['%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d']
                for fmt in date_formats:
                    try:
                        parsed_date = datetime.strptime(date_value.strip(), fmt)
                        return parsed_date.strftime('%Y-%m-%d')
                    except ValueError:
                        continue
                
                # 如果无法解析，返回原字符串但只取前10个字符（可能包含日期）
                clean_date = str(date_value).strip()[:10]
                if len(clean_date) >= 8:  # 确保看起来像日期
                    return clean_date
                else:
                    return str(date_value)
            
            return str(date_value)
        except:
            return str(date_value)
    
    def get_data_signature(self, parsed_data):
        """生成数据签名用于去重"""
        # 使用名字、位置、租金、租期、押金方式作为唯一标识
        return f"{parsed_data.get('名字', '')}|{parsed_data.get('位置', '')}|{parsed_data.get('租金', '')}|{parsed_data.get('租期', '')}|{parsed_data.get('押金方式', '')}"
    
    def is_duplicate_data(self, parsed_data):
        """检查是否为重复数据（与Excel数据和新数据对比）"""
        signature = self.get_data_signature(parsed_data)
        
        # 检查是否在已处理的签名中（包括Excel数据和新数据）
        if signature in self.processed_signatures:
            return True
        
        # 实时检查Excel数据中是否有重复
        for excel_record in self.existing_excel_data:
            excel_signature = self.get_data_signature(excel_record)
            if signature == excel_signature:
                return True
        
        return False
    
    def find_duplicate_in_excel(self, parsed_data):
        """在Excel数据中查找重复记录"""
        signature = self.get_data_signature(parsed_data)
        
        for excel_record in self.existing_excel_data:
            excel_signature = self.get_data_signature(excel_record)
            if signature == excel_signature:
                return excel_record
        
        return None
    
    def parse_statistics_message(self, text):
        """解析统计消息（适配不同的关键词）"""
        try:
            # 移除开头的关键词（使用配置中的所有关键词）
            for keyword in self.monitor_keywords:
                if not keyword:
                    continue
                    
                # 构建匹配模式
                patterns = [
                    f'^{keyword}[，, ：:]*',
                    f'^{keyword}[，, ：:]*',
                    f'^{keyword}[，, ：:]*',
                    f'^{keyword}[，, ：:]*',
                    f'^{keyword}[，, ：:]*',
                    f'^{keyword}'
                ]
                
                for pattern in patterns:
                    new_text = re.sub(pattern, '', text, count=1, flags=re.IGNORECASE if not self.case_sensitive else 0)
                    if new_text != text:  # 如果成功移除了关键词
                        text = new_text
                        break
                
                if text != text:  # 这个条件永远不会满足，只是为了break
                    break
            
            # 分割字段
            parts = re.split(r'[，, ：:]', text)
            
            if len(parts) >= 5:
                return {
                    '名字': parts[0].strip(),
                    '位置': parts[1].strip(),
                    '租金': parts[2].strip(),
                    '租期': parts[3].strip(),
                    '押金方式': parts[4].strip() if len(parts) > 4 else '',
                    '时间': datetime.now().strftime('%Y-%m-%d'),
                    '原始消息': text,
                    '备注': ''  # 确保有备注字段
                }
            else:
                return None
        except Exception as e:
            print(f"解析消息失败: {e}")
            return None
    
    def monitor_clipboard(self):
        """监控剪贴板线程"""
        check_interval = 1.0  # 检查间隔（秒）
        
        while self.monitoring:
            try:
                current_clipboard = pyperclip.paste()
                
                if current_clipboard != self.last_clipboard:
                    self.last_clipboard = current_clipboard
                    
                    # 添加到历史记录
                    timestamp = datetime.now().strftime('%H:%M:%S')
                    self.clipboard_history.append({
                        'time': timestamp,
                        'content': current_clipboard[:100] + '...' if len(current_clipboard) > 100 else current_clipboard,
                        'full_content': current_clipboard
                    })
                    
                    # 限制历史记录数量
                    if len(self.clipboard_history) > 100:
                        self.clipboard_history = self.clipboard_history[-100:]
                    
                    # 检查是否为统计消息
                    if self.is_statistics_message(current_clipboard.strip()):
                        parsed_data = self.parse_statistics_message(current_clipboard.strip())
                        if parsed_data:
                            # 检查是否重复
                            if self.is_duplicate_data(parsed_data):
                                # 发送重复提醒消息
                                self.message_queue.put(('duplicate_found', parsed_data))
                            else:
                                # 添加到已处理数据
                                signature = self.get_data_signature(parsed_data)
                                self.processed_signatures.add(signature)
                                parsed_data['序号'] = len(self.existing_excel_data) + len(self.processed_data) + 1
                                self.processed_data.append(parsed_data)
                                
                                # 发送消息到队列
                                self.message_queue.put(('statistics_found', parsed_data))
                    
                    # 发送更新消息
                    if statistics_lines:
                        self.message_queue.put(('clipboard_updated', current_clipboard))
                
                time.sleep(check_interval)
                
            except Exception as e:
                print(f"监控剪贴板出错: {e}")
                time.sleep(check_interval)
    
    def process_messages(self):
        """处理消息队列"""
        try:
            while True:
                if not self.message_queue.empty():
                    msg_type, data = self.message_queue.get_nowait()
                    
                    if msg_type == 'statistics_found':
                        self.add_statistics_to_tree(data)
                        self.show_notification(f"发现有效统计记录: {data['名字']}")
                    
                    elif msg_type == 'duplicate_found':
                        self.show_duplicate_warning(data)
                    
                    elif msg_type == 'clipboard_updated':
                        self.update_history_display()
                        self.update_stats()
                    
                    elif msg_type == 'status_update':
                        self.status_var.set(data)
                        self.last_update_var.set(f"最后更新: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    
                    elif msg_type == 'switch_to_clipboard_tab':
                        # 自动跳转到剪贴板历史标签页
                        try:
                            # 找到剪贴板历史标签页的索引（通常是第2个，索引为1）
                            self.notebook.select(1)  # 切换到剪贴板历史标签页
                        except:
                            pass  # 如果切换失败，忽略
                
                self.root.after(100, self.process_messages)
                break
                
        except:
            self.root.after(100, self.process_messages)
    
    def update_history_display(self):
        """更新历史记录显示"""
        self.history_text.delete(1.0, tk.END)
        
        for i, item in enumerate(reversed(self.clipboard_history[-20:])):  # 显示最近20条
            time_str = item['time']
            content = item['content']
            full_content = item['full_content']
            
            # 行级别检测：检查剪贴板内容中是否包含统计行
            statistics_lines = self.extract_statistics_from_text(full_content)
            has_stats_lines = len(statistics_lines) > 0
            
            if has_stats_lines:
                # 有统计行时，美化显示
                prefix = f"📊 发现 {len(statistics_lines)} 条统计"
                
                # 显示统计行预览（更清晰的格式）
                self.history_text.insert(tk.END, f"[{time_str}] {prefix}\n", "time")
                
                for j, stat_info in enumerate(statistics_lines[:3]):  # 显示前3条
                    line_num = stat_info['line_number']
                    line_content = stat_info['line_content']
                    
                    # 解析统计信息显示更友好的预览
                    parsed_data = stat_info.get('parsed_data')
                    if parsed_data:
                        name = parsed_data.get('名字', '')
                        location = parsed_data.get('位置', '')
                        rent = parsed_data.get('租金', '')
                        preview = f"📍 第{line_num}行: {name} · {location} · {rent}"
                    else:
                        # 如果解析失败，显示原始内容
                        preview = f"📍 第{line_num}行: {line_content[:50]}{'...' if len(line_content) > 50 else ''}"
                    
                    self.history_text.insert(tk.END, f"    {preview}\n", "stats_item")
                
                if len(statistics_lines) > 3:
                    self.history_text.insert(tk.END, f"    ... 还有 {len(statistics_lines) - 3} 条统计\n", "stats_more")
                
                self.history_text.insert(tk.END, "\n", "separator")
                
            else:
                # 没有统计行
                prefix = "📝 "
                self.history_text.insert(tk.END, f"[{time_str}] {prefix}{content}\n", "normal")
        
        # 设置各种消息的样式
        self.history_text.tag_config("time", foreground="#6c757d", font=("Microsoft YaHei", 8))
        self.history_text.tag_config("stats", foreground="#0066cc", font=("Microsoft YaHei", 9, "bold"))
        self.history_text.tag_config("stats_item", foreground="#28a745", font=("Microsoft YaHei", 9))
        self.history_text.tag_config("stats_more", foreground="#6c757d", font=("Microsoft YaHei", 8, "italic"))
        self.history_text.tag_config("normal", foreground="#495057", font=("Microsoft YaHei", 9))
        self.history_text.tag_config("separator", foreground="#dee2e6")
    
    def update_stats(self):
        """更新统计信息"""
        excel_count = len(self.existing_excel_data)
        new_count = len(self.processed_data)
        total_count = excel_count + new_count
        
        stats_text = f"剪贴板历史: {len(self.clipboard_history)} | Excel数据: {excel_count} | 新数据: {new_count} | 总计: {total_count}"
        self.stats_label.config(text=stats_text)
    
    def add_statistics_to_tree(self, data):
        """添加统计记录到表格"""
        remark = data.get('备注', '')
        self.result_tree.insert('', tk.END, values=(
            data['序号'],
            data['名字'],
            data['位置'],
            data['租金'],
            data['租期'],
            data['押金方式'],
            data['时间'],
            remark
        ))
    
    def start_monitoring(self):
        """开始监控"""
        if not self.monitoring:
            # 清空剪贴板内容
            try:
                pyperclip.copy('')
                self.message_queue.put(('status_update', '已清空剪贴板 - 准备开始监控...'))
            except:
                pass
            
            self.monitoring = True
            
            # 启动监控线程（使用改进的方法）
            monitor_thread = threading.Thread(target=self.monitor_clipboard_with_line_detection, daemon=True)
            monitor_thread.start()
            
            # 更新界面状态
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.progress.start()
            
            self.message_queue.put(('status_update', '监控已启动 - 正在监控剪贴板（行级别检测）...'))
    
    def stop_monitoring(self):
        """停止监控"""
        if self.monitoring:
            self.monitoring = False
            
            # 更新界面状态
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.progress.stop()
            
            self.message_queue.put(('status_update', '监控已停止'))
    
    def check_clipboard_once(self):
        """立即检查剪贴板（行级别检测）"""
        try:
            current_clipboard = pyperclip.paste()
            if current_clipboard:
                # 提取统计行
                statistics_lines = self.extract_statistics_from_text(current_clipboard.strip())
                
                if not statistics_lines:
                    messagebox.showinfo("检查结果", "剪贴板内容中没有以'统计'开头的有效消息")
                    return
                
                found_count = 0
                duplicate_count = 0
                
                for stat_info in statistics_lines:
                    parsed_data = stat_info['parsed_data']
                    line_content = stat_info['line_content']
                    line_number = stat_info['line_number']
                    
                    if self.is_duplicate_data(parsed_data):
                        duplicate_count += 1
                        response = messagebox.askyesno(
                            "重复数据提示", 
                            f"第{line_number}行发现重复的统计记录：\n\n"
                            f"内容: {line_content}\n\n"
                            f"名字: {parsed_data['名字']}\n"
                            f"位置: {parsed_data['位置']}\n"
                            f"租金: {parsed_data['租金']}\n\n"
                            f"是否仍然添加此记录？"
                        )
                        if response:
                            signature = self.get_data_signature(parsed_data)
                            self.processed_signatures.add(signature)
                            parsed_data['序号'] = len(self.existing_excel_data) + len(self.processed_data) + 1
                            parsed_data['备注'] = '重复记录（用户确认添加）'
                            parsed_data['来源'] = '剪贴板'
                            parsed_data['行号'] = line_number
                            parsed_data['行内容'] = line_content
                            self.processed_data.append(parsed_data)
                            found_count += 1
                    else:
                        signature = self.get_data_signature(parsed_data)
                        self.processed_signatures.add(signature)
                        parsed_data['序号'] = len(self.existing_excel_data) + len(self.processed_data) + 1
                        parsed_data['来源'] = '剪贴板'
                        parsed_data['行号'] = line_number
                        parsed_data['行内容'] = line_content
                        self.processed_data.append(parsed_data)
                        found_count += 1
                
                # 刷新显示（包含Excel数据）
                self.refresh_data()
                
                messagebox.showinfo(
                    "检查完成", 
                    f"剪贴板检查完成:\n\n"
                    f"总统计行数: {len(statistics_lines)}\n"
                    f"新发现记录: {found_count - duplicate_count}\n"
                    f"重复记录: {duplicate_count}\n"
                    f"实际添加: {found_count}\n\n"
                    f"（只处理了包含'统计'的行）"
                )
                
            else:
                messagebox.showwarning("剪贴板为空", "剪贴板中没有内容")
                
        except Exception as e:
            messagebox.showerror("检查失败", f"检查剪贴板时出错: {e}")
    
    def clear_history(self):
        """清空历史记录"""
        if messagebox.askyesno("确认清空", "确定要清空所有历史记录和重置重复检测吗？"):
            # 清空剪贴板内容
            try:
                pyperclip.copy('')
            except:
                pass
            
            self.clipboard_history.clear()
            self.processed_data.clear()
            self.existing_excel_data.clear()  # 清空Excel缓存数据
            self.processed_signatures.clear()  # 清空重复检测记录
            
            # 清空界面显示
            self.history_text.delete(1.0, tk.END)
            
            # 清空树形控件
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            
            self.update_stats()
            self.message_queue.put(('status_update', '已清空所有历史记录、Excel数据和剪贴板'))
            messagebox.showinfo("清空完成", "历史记录、Excel数据和重复检测已重置，剪贴板已清空")
    
    def export_to_excel(self):
        """导出到Excel（追加模式）"""
        if not self.processed_data and not self.existing_excel_data:
            messagebox.showwarning("无数据", "没有数据可导出")
            return
        
        try:
            # 合并所有数据：Excel数据 + 新数据
            all_data = self.existing_excel_data + self.processed_data
            
            if not all_data:
                messagebox.showwarning("无数据", "没有数据可导出")
                return
            
            # 确保所有记录都有完整字段
            required_fields = ['序号', '名字', '位置', '租金', '租期', '押金方式', '时间', '备注', '来源']
            for data in all_data:
                for field in required_fields:
                    if field not in data:
                        data[field] = ''
            
            # 重新编号所有数据，保持与Excel实时显示一致
            for i, data in enumerate(all_data):
                data['序号'] = i + 1
            
            # 创建DataFrame
            df = pd.DataFrame(all_data)
            
            # 选择要导出的列（包含来源信息）
            export_columns = ['序号', '名字', '位置', '租金', '租期', '押金方式', '时间', '备注']
            df_export = df[export_columns]
            
            # 导出到Excel
            df_export.to_excel(self.excel_file, index=False, engine='openpyxl')
            
            # 更新内存中的Excel数据，保持序号一致
            self.existing_excel_data = all_data.copy()
            
            # 清空已导出的新数据
            self.processed_data.clear()
            
            # 刷新Excel显示和新数据显示
            self.refresh_excel_display()
            self.refresh_data()
            
            self.save_config()
            
            # 统计信息
            excel_count = len(self.existing_excel_data)
            new_count = len(self.processed_data)
            total_count = len(all_data)
            
            messagebox.showinfo(
                "导出成功", 
                f"数据已成功导出到:\n{self.excel_file}\n\n"
                f"Excel原有记录: {excel_count}\n"
                f"新追加记录: {new_count}\n"
                f"总计记录: {total_count}\n\n"
                f"（已追加到原Excel文件末尾）"
            )
            
        except Exception as e:
            messagebox.showerror("导出失败", f"导出到Excel失败: {e}")
    
    def open_excel(self):
        """打开Excel文件"""
        if not os.path.exists(self.excel_file):
            if messagebox.askyesno("文件不存在", f"Excel文件不存在，是否创建新文件？\n{self.excel_file}"):
                self.export_to_excel()
            return
        
        try:
            os.startfile(self.excel_file)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开Excel文件: {e}")
    
    def delete_selected(self):
        """删除选中记录（只删除新数据，不能删除Excel数据）"""
        selection = self.result_tree.selection()
        if not selection:
            messagebox.showwarning("未选择", "请先选择要删除的记录")
            return
        
        if messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selection)} 条新记录吗？\n注意：只能删除新数据，Excel数据不可删除"):
            # 获取选中的序号
            selected_indices = []
            for item in selection:
                values = self.result_tree.item(item, 'values')
                selected_indices.append(int(values[0]) - 1)  # 序号从1开始，索引从0开始
            
            # 删除记录（从后往前删除，避免索引变化）
            deleted_count = 0
            for index in sorted(selected_indices, reverse=True):
                if 0 <= index < len(self.processed_data):
                    # 从重复检测签名集中移除
                    deleted_data = self.processed_data[index]
                    signature = self.get_data_signature(deleted_data)
                    self.processed_signatures.discard(signature)
                    
                    del self.processed_data[index]
                    deleted_count += 1
            
            # 重新编号和确保所有字段完整
            for i, data in enumerate(self.processed_data):
                data['序号'] = i + 1
                if '备注' not in data:
                    data['备注'] = ''
            
            # 刷新显示
            self.refresh_data()
            messagebox.showinfo("删除完成", f"已删除 {deleted_count} 条新记录")
    
    def refresh_data(self):
        """刷新数据显示"""
        # 清空新数据树形控件（只显示新数据，不显示Excel数据）
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        # 只显示新数据（processed_data），序号基于Excel记录数连续编号
        excel_count = len(self.existing_excel_data)
        for i, data in enumerate(self.processed_data):
            data['序号'] = excel_count + i + 1  # 新数据序号从Excel末尾开始
            self.add_statistics_to_tree(data)
        
        # 单独刷新Excel数据显示
        self.refresh_excel_display()
        
        # 更新统计信息
        total_data = len(self.existing_excel_data) + len(self.processed_data)
        self.message_queue.put(('status_update', f'数据显示更新 - Excel数据: {len(self.existing_excel_data)}, 新数据: {len(self.processed_data)}, 总计: {total_data}'))
    
    def show_notification(self, message):
        """显示通知"""
        # 在状态栏显示临时消息
        original_status = self.status_var.get()
        self.status_var.set(f"🎉 {message}")
        self.root.after(3000, lambda: self.status_var.set(original_status))
    
    def show_duplicate_warning(self, data):
        """显示重复数据警告（带Excel对比）"""
        # 在状态栏显示重复警告
        original_status = self.status_var.get()
        self.status_var.set(f"⚠️ 发现重复数据: {data['名字']} - {data['位置']}")
        
        # 查找Excel中的重复记录
        excel_duplicate = self.find_duplicate_in_excel(data)
        
        # 构建详细的对比信息
        message_text = "发现重复的统计记录：\n\n"
        message_text += "【新数据】\n"
        message_text += f"名字: {data['名字']}\n"
        message_text += f"位置: {data['位置']}\n"
        message_text += f"租金: {data['租金']}\n"
        message_text += f"租期: {data['租期']}\n"
        message_text += f"押金方式: {data['押金方式']}\n"
        message_text += f"时间: {data['时间']}\n"
        
        if excel_duplicate:
            message_text += "\n【Excel中的重复记录】\n"
            message_text += f"名字: {excel_duplicate['名字']}\n"
            message_text += f"位置: {excel_duplicate['位置']}\n"
            message_text += f"租金: {excel_duplicate['租金']}\n"
            message_text += f"租期: {excel_duplicate['租期']}\n"
            message_text += f"押金方式: {excel_duplicate['押金方式']}\n"
            message_text += f"时间: {excel_duplicate['时间']}\n"
            if excel_duplicate.get('备注'):
                message_text += f"备注: {excel_duplicate['备注']}\n"
        
        message_text += "\n是否仍然添加此记录？"
        
        # 显示确认对话框
        response = messagebox.askyesno("重复数据提示", message_text)
        
        if response:
            # 用户选择添加重复记录
            signature = self.get_data_signature(data)
            self.processed_signatures.add(signature)  # 即使重复也添加，避免再次提示
            data['序号'] = len(self.existing_excel_data) + len(self.processed_data) + 1
            data['备注'] = f'重复记录（用户确认）'
            if excel_duplicate:
                data['备注'] += f' - Excel序号:{excel_duplicate["序号"]}'
            self.processed_data.append(data)
            self.add_statistics_to_tree(data)
            self.update_stats()
        
        # 恢复原始状态
        self.root.after(100, lambda: self.status_var.set(original_status))
    
    def open_config_dialog(self):
        """打开配置对话框"""
        config_window = tk.Toplevel(self.root)
        config_window.title("🔧 监控配置")
        config_window.geometry("580x650")
        config_window.resizable(True, True)
        
        # 设置窗口位置居中
        config_window.transient(self.root)
        config_window.grab_set()
        
        # 创建主滚动框架
        canvas = tk.Canvas(config_window)
        scrollbar = ttk.Scrollbar(config_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 创建配置界面
        main_frame = ttk.Frame(scrollable_frame, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="🔧 监控配置", font=("Microsoft YaHei", 12, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 监控关键词配置
        keyword_frame = ttk.LabelFrame(main_frame, text="🔍 监控关键词设置", padding="8")
        keyword_frame.pack(fill=tk.X, pady=(0, 8))
        
        keyword_label = ttk.Label(keyword_frame, text="监控关键词 (支持特殊格式):", font=('Microsoft YaHei', 9, 'bold'))
        keyword_label.pack(anchor=tk.W, pady=(0, 3))
        
        keywords_text = tk.Text(keyword_frame, height=3, width=55, font=('Microsoft YaHei', 9))
        keywords_text.pack(fill=tk.X, pady=3)
        keywords_text.insert('1.0', '\n'.join(self.monitor_keywords))
        
        # 关键词格式说明 - 紧凑版本
        format_frame = ttk.LabelFrame(main_frame, text="📝 关键词格式说明", padding="6")
        format_frame.pack(fill=tk.X, pady=(0, 8))
        
        format_text = tk.Text(format_frame, height=4, width=55, wrap=tk.WORD, font=('Microsoft YaHei', 8))
        format_text.pack(fill=tk.X)
        format_text.insert('1.0', 
            "📝 普通中文: 统计, 记账, 支出\n"
            "🔤 特殊字符: ￥, *, #, @, $, €, £, ¥, %\n"
            "🔤+中文组合: *房租, #记账, @支出, %工资\n"
            "🌍 外币符号: €房租, £租金, ¥水电, $房租"
        )
        format_text.config(state='disabled')
        
        # 监控模式 - 紧凑版本
        mode_frame = ttk.LabelFrame(main_frame, text="⚙️ 监控模式设置", padding="6")
        mode_frame.pack(fill=tk.X, pady=(0, 8))
        
        mode_var = tk.StringVar(value=self.monitor_mode)
        mode_label = ttk.Label(mode_frame, text="匹配方式:", font=('Microsoft YaHei', 9, 'bold'))
        mode_label.pack(anchor=tk.W, pady=(0, 3))
        
        mode_radio_frame = ttk.Frame(mode_frame)
        mode_radio_frame.pack(fill=tk.X, pady=3)
        
        ttk.Radiobutton(mode_radio_frame, text="🎯 开头匹配", variable=mode_var, value="startswith").pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(mode_radio_frame, text="🔍 包含匹配", variable=mode_var, value="contains").pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(mode_radio_frame, text="⚡ 精确匹配", variable=mode_var, value="exact").pack(side=tk.LEFT, padx=8)
        
        # 大小写敏感
        case_var = tk.BooleanVar(value=self.case_sensitive)
        case_frame = ttk.Frame(mode_frame)
        case_frame.pack(fill=tk.X, pady=5)
        
        ttk.Checkbutton(case_frame, text="🔤 区分大小写", variable=case_var).pack(side=tk.LEFT)
        
        # 模式说明 - 简化
        mode_help = ttk.Label(mode_frame, text="💡 开头匹配支持特殊字符，包含匹配检查文本内容", 
                              font=('Microsoft YaHei', 8), foreground='#6c757d')
        mode_help.pack(anchor=tk.W, pady=(2, 0))
        
        # 界面主题和刷新间隔 - 紧凑版本
        settings_frame = ttk.LabelFrame(main_frame, text="🎨 界面设置", padding="6")
        settings_frame.pack(fill=tk.X, pady=(0, 8))
        
        # 界面主题和刷新间隔在一行
        settings_row_frame = ttk.Frame(settings_frame)
        settings_row_frame.pack(fill=tk.X, pady=3)
        
        # 界面主题
        theme_label = ttk.Label(settings_row_frame, text="主题:", font=('Microsoft YaHei', 9, 'bold'))
        theme_label.pack(side=tk.LEFT, padx=(0, 8))
        
        theme_var = tk.StringVar(value=self.ui_theme)
        theme_frame = ttk.Frame(settings_row_frame)
        theme_frame.pack(side=tk.LEFT, padx=8)
        
        ttk.Radiobutton(theme_frame, text="🌟现代", variable=theme_var, value="modern").pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(theme_frame, text="📜经典", variable=theme_var, value="classic").pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(theme_frame, text="⚪简洁", variable=theme_var, value="minimal").pack(side=tk.LEFT, padx=3)
        
        # 自动刷新间隔
        interval_frame = ttk.Frame(settings_row_frame)
        interval_frame.pack(side=tk.LEFT, padx=(15, 0))
        
        ttk.Label(interval_frame, text="⏰刷新(秒):", font=('Microsoft YaHei', 9)).pack(side=tk.LEFT)
        interval_var = tk.IntVar(value=self.auto_refresh_interval // 1000)
        interval_spin = ttk.Spinbox(interval_frame, from_=1, to=60, textvariable=interval_var, width=6)
        interval_spin.pack(side=tk.LEFT, padx=5)
        
        # 预览区域 - 紧凑版本
        preview_frame = ttk.LabelFrame(main_frame, text="🧪 关键词测试", padding="6")
        preview_frame.pack(fill=tk.X, pady=(0, 8))
        
        preview_label = ttk.Label(preview_frame, text="📋 测试文本示例:", font=('Microsoft YaHei', 9, 'bold'))
        preview_label.pack(anchor=tk.W, pady=(0, 3))
        
        preview_text = tk.Text(preview_frame, height=5, width=55, wrap=tk.WORD, font=('Microsoft YaHei', 8))
        preview_text.pack(fill=tk.X)
        preview_text.insert('1.0', 
            "￥张三,朝阳区,4500,1年,押一付三\n"
            "统计 李四,海淀区,5500,1年,押一付三\n"
            "#房租 赵六,东城区,5000,1年,押一付三\n"
            "@支出 钱七,丰台区,3800,半年,押一付一\n"
            "€租金 周九,通州区,4200,1年,押一付三"
        )
        
        # 实时统计信息和按钮 - 紧凑版本
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 统计信息
        stats_label = ttk.Label(bottom_frame, text="📊 实时匹配: 0/0 行", font=('Microsoft YaHei', 9, 'bold'))
        stats_label.pack(pady=(0, 8))
        
        # 按钮区域
        button_frame = ttk.Frame(bottom_frame)
        button_frame.pack(fill=tk.X)
        
        def test_keywords():
            """测试关键词匹配"""
            keywords = [kw.strip() for kw in keywords_text.get('1.0', tk.END).strip().split('\n') if kw.strip()]
            mode = mode_var.get()
            case_sensitive = case_var.get()
            
            content = preview_text.get('1.0', tk.END).strip()
            lines = content.split('\n')
            matched_lines = []
            match_count = 0
            
            for line in lines:
                matched = False
                for keyword in keywords:
                    if self.test_keyword_match(line, keyword, mode, case_sensitive):
                        matched_lines.append(f"✅ {line}")
                        matched = True
                        match_count += 1
                        break
                if not matched:
                    matched_lines.append(f"⭕ {line}")
            
            result_text = '\n'.join(matched_lines)
            total_lines = len(lines)
            
            # 显示更详细的测试结果
            test_window = tk.Toplevel(config_window)
            test_window.title("🧪 关键词测试结果")
            test_window.geometry("550x300")
            test_window.resizable(False, False)
            
            result_frame = ttk.Frame(test_window, padding="15")
            result_frame.pack(fill=tk.BOTH, expand=True)
            
            # 统计信息
            stats_label = ttk.Label(result_frame, 
                                   text=f"📊 测试统计：匹配 {match_count}/{total_lines} 行", 
                                   font=('Microsoft YaHei', 11, 'bold'))
            stats_label.pack(pady=(0, 8))
            
            # 关键词信息
            keywords_info = f"🔍 关键词: {', '.join(keywords[:3])}{'...' if len(keywords) > 3 else ''}"
            ttk.Label(result_frame, text=keywords_info, font=('Microsoft YaHei', 9)).pack(pady=(0, 5))
            
            mode_info = f"⚙️ 模式: {'开头匹配' if mode == 'startswith' else '包含匹配' if mode == 'contains' else '精确匹配'} | 大小写: {'是' if case_sensitive else '否'}"
            ttk.Label(result_frame, text=mode_info, font=('Microsoft YaHei', 9)).pack(pady=(0, 10))
            
            # 匹配结果
            result_text_widget = tk.Text(result_frame, height=8, width=60, wrap=tk.WORD, font=('Consolas', 9))
            result_text_widget.pack(fill=tk.BOTH, expand=True)
            result_text_widget.insert('1.0', result_text)
            result_text_widget.config(state='disabled')
            
            ttk.Button(test_window, text="关闭", command=test_window.destroy).pack(pady=8)
        
        def save_settings():
            """保存配置"""
            try:
                # 验证关键词
                keywords = [kw.strip() for kw in keywords_text.get('1.0', tk.END).strip().split('\n') if kw.strip()]
                if not keywords:
                    messagebox.showwarning("关键词为空", "请至少输入一个监控关键词！")
                    return
                
                # 更新配置
                self.monitor_keywords = keywords
                self.monitor_mode = mode_var.get()
                self.case_sensitive = case_var.get()
                self.ui_theme = theme_var.get()
                self.auto_refresh_interval = interval_var.get() * 1000
                
                # 保存配置文件
                self.save_config()
                
                # 应用界面主题
                self.apply_theme(self.ui_theme)
                
                messagebox.showinfo("✅ 配置成功", 
                                  f"已保存配置！\n"
                                  f"关键词: {len(keywords)}个 | 模式: {mode_var.get()}")
                config_window.destroy()
                
            except Exception as e:
                messagebox.showerror("❌ 配置失败", f"保存配置时出错: {e}")
        
        # 按钮样式优化
        ttk.Button(button_frame, text="🧪 测试", command=test_keywords, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="💾 保存", command=save_settings, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="❌ 取消", command=config_window.destroy, width=12).pack(side=tk.RIGHT, padx=5)
        
        # 绑定实时预览
        def update_preview(event=None):
            keywords = [kw.strip() for kw in keywords_text.get('1.0', tk.END).strip().split('\n') if kw.strip()]
            mode = mode_var.get()
            case_sensitive = case_var.get()
            
            content = preview_text.get('1.0', tk.END).strip()
            lines = [line for line in content.split('\n') if line.strip()]  # 过滤空行
            
            # 清除所有标签
            preview_text.tag_remove('match', '1.0', tk.END)
            preview_text.tag_remove('partial_match', '1.0', tk.END)
            
            match_count = 0
            for i, line in enumerate(lines, 1):
                matched = False
                for keyword in keywords:
                    if self.test_keyword_match(line, keyword, mode, case_sensitive):
                        start = f"{i}.0"
                        end = f"{i}.end"
                        preview_text.tag_add('match', start, end)
                        match_count += 1
                        matched = True
                        break
                
                # 如果没有完全匹配，检查是否有部分匹配（用于包含模式）
                if not matched and mode == 'contains':
                    for keyword in keywords:
                        if keyword.lower() in line.lower() if not case_sensitive else keyword in line:
                            start = f"{i}.0"
                            end = f"{i}.end"
                            preview_text.tag_add('partial_match', start, end)
                            break
            
            # 更新统计信息
            stats_label.config(text=f"📊 实时匹配: {match_count}/{len(lines)} 行")
            
            # 配置标签样式
            preview_text.tag_config('match', background='#d4edda', foreground='#155724', font=('Microsoft YaHei', 8, 'bold'))
            preview_text.tag_config('partial_match', background='#fff3cd', foreground='#856404')
        
        # 绑定事件
        keywords_text.bind('<KeyRelease>', update_preview)
        keywords_text.bind('<Button-1>', update_preview)
        mode_var.trace_add('write', lambda *args: update_preview())
        case_var.trace_add('write', lambda *args: update_preview())
        
        # 初始化预览
        config_window.after(100, update_preview)
        
        # 配置滚动区域
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 配置鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)
        
        # 让canvas可以接收焦点和配置滚动区域
        canvas.focus_set()
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    
    def test_keyword_match(self, text, keyword, mode, case_sensitive):
        """测试关键词匹配（与主程序保持一致）"""
        if not keyword or not text:
            return False
        
        # 清理文本首尾空白
        text = text.strip()
        
        if not case_sensitive:
            text_lower = text.lower()
            keyword_lower = keyword.lower()
        else:
            text_lower = text
            keyword_lower = keyword
        
        if mode == 'startswith':
            # 检查是否以关键词开头（支持特殊字符和字符+中文组合）
            try:
                # 直接开头匹配
                if text.startswith(keyword):
                    return True
                
                # 处理关键词后跟分隔符的情况
                separators = ['，', ',', '：', ':', '、', ' ', '\t', '-', '—', '——']
                for sep in separators:
                    if text.startswith(keyword + sep):
                        return True
                
                # 特殊处理：如果关键词包含特殊字符，使用更灵活的正则匹配
                if any(char in keyword for char in ['*', '#', '@', '￥', '$', '€', '£', '¥', '%']):
                    # 构建正则表达式模式，支持关键词后跟任意分隔符
                    escaped_keyword = re.escape(keyword)
                    pattern = f"^{escaped_keyword}[，, ：:、\s\-———]*.*"
                    if re.match(pattern, text):
                        return True
                
                # 如果是纯中文关键词，也使用正则匹配
                if re.match(r'^[\u4e00-\u9fff]+$', keyword):
                    escaped_keyword = re.escape(keyword)
                    pattern = f"^{escaped_keyword}[，, ：:、\s\-———]*.*"
                    if re.match(pattern, text):
                        return True
            
            except:
                # 正则表达式出错时，使用简单匹配
                if text.startswith(keyword):
                    return True
                    
        elif mode == 'contains':
            # 检查是否包含关键词
            return keyword_lower in text_lower
            
        elif mode == 'exact':
            # 检查是否精确匹配
            return text_lower == keyword_lower
        
        return False
    
    def apply_theme(self, theme_name):
        """应用界面主题"""
        try:
            style = ttk.Style()
            
            if theme_name == 'modern':
                style.theme_use('clam')
                style.configure('TNotebook', background='#f0f0f0')
                style.configure('TNotebook.Tab', padding=[20, 8], background='#e0e0e0')
                style.map('TNotebook.Tab', background=[('selected', '#ffffff')])
                
            elif theme_name == 'classic':
                style.theme_use('default')
                
            elif theme_name == 'minimal':
                style.theme_use('alt')
                style.configure('TButton', padding=[10, 5])
                
        except Exception as e:
            print(f"应用主题失败: {e}")
    
    def show_help(self):
        """显示使用说明"""
        help_text = """
微信统计信息监控器 - 使用说明 v3.0

新增功能：
🔥 实时Excel数据预览窗口 - 自动显示Excel文件内容
🔥 智能Excel重复对比 - 新数据与Excel实时对比检测
🔥 自动刷新Excel数据 - 5秒间隔自动更新Excel显示

功能说明：
1. 实时监控剪贴板内容（行级别检测）
2. 实时读取和显示Excel文件内容
3. 智能重复检测（新数据与Excel实时对比）
4. 严格识别以"统计"开头的微信消息（同一行内容）
5. 自动解析并提取数据
6. 支持追加导出到Excel文件

界面布局：
📊 Excel实时数据 - 实时显示Excel文件所有记录
📝 剪贴板历史 - 显示剪贴板监控历史，高亮统计行
🆕 新数据预览 - 显示新解析的数据（非重复项）

使用步骤：
1. 启动程序后自动加载Excel数据（如果文件存在）
2. 切换到"Excel实时数据"标签页查看现有数据
3. 点击"开始监控"按钮启动监控
4. 从微信聊天窗口复制内容（Ctrl+C）
5. 程序自动与Excel数据对比，发现重复时显示详细对比
6. 有效数据自动添加到"新数据预览"标签页
7. 点击"导出到Excel"追加保存数据

重复检测机制：
🔍 Excel实时对比：新数据与Excel中每条记录实时比较
📋 详细对比信息：显示新旧数据的完整字段对比
⚡ 自动检测：监控过程中自动发现并提示重复项
👤 用户确认：可选择是否添加重复记录

Excel实时数据功能：
🔄 自动刷新：每5秒自动更新Excel数据显示
📊 完整显示：显示Excel中所有记录（包含来源标识）
🎯 重复检测：作为新数据重复检测的基准
📈 状态统计：实时显示Excel记录数量

消息格式要求：
• 必须严格以"统计"开头的行
• 只处理与"统计"在同一行的内容
• 支持：统计，名字，位置，租金，租期，押金方式

智能重复检测：
• 新数据与Excel中所有记录进行实时对比
• 显示重复记录的详细信息（Excel序号、时间等）
• 支持用户选择是否添加重复记录
• 重复检测基于：名字、位置、租金、租期、押金方式

快捷操作：
• 读取Excel - 手动加载Excel文件
• 刷新Excel数据 - 立即更新Excel显示
• 自动刷新: 开启/关闭 - 控制自动刷新功能
• 开始监控 - 启动剪贴板监控
• 立即检查剪贴板 - 手动检查当前剪贴板
• 导出到Excel - 追加新数据到Excel文件

注意事项：
• 程序启动时自动加载Excel数据（如果存在）
• Excel数据会实时参与重复检测
• 自动刷新确保Excel数据与文件同步
• 重复提示会显示Excel中对应记录的详细信息
• 建议保持"自动刷新"开启以获得最佳体验

推荐工作流程：
1. 程序启动后查看"Excel实时数据"标签页
2. 确认现有数据已正确加载
3. 点击"开始监控"开始工作
4. 复制微信内容，程序自动去重
5. 在"新数据预览"中查看有效数据
6. 点击"导出到Excel"保存新数据
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("600x500")
        help_window.resizable(True, True)
        
        text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(1.0, help_text)
        text_widget.config(state=tk.DISABLED)
        
        ttk.Button(help_window, text="关闭", command=help_window.destroy).pack(pady=10)

def main():
    root = tk.Tk()
    
    # 设置样式
    style = ttk.Style()
    style.configure("Success.TButton", foreground="green")
    style.configure("Danger.TButton", foreground="red")
    
    app = WeChatClipboardMonitor(root)
    root.mainloop()

if __name__ == "__main__":
    main()