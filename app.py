import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd
import threading
import os
import sys
from datetime import datetime
import webbrowser

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入自定义模块
from base import get_weibo_list
from utils.data_processor import WeiboDataProcessor

class WeiboSpiderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🐦 微博数据爬虫分析平台")
        self.root.geometry("1400x900")
        self.root.resizable(True, True)
        
        # 设置窗口背景色
        self.root.configure(bg='#f0f0f0')
        
        # 设置现代化主题样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 设置全局字体（必须在configure_modern_styles之前）
        self.default_font = ('Microsoft YaHei UI', 10)
        self.heading_font = ('Microsoft YaHei UI', 12, 'bold')
        self.title_font = ('Microsoft YaHei UI', 18, 'bold')
        
        # 配置现代化样式
        self.configure_modern_styles()
        
        # 数据变量
        self.df = pd.DataFrame()
        self.is_crawling = False
        
        # 创建界面
        self.create_widgets()
        
        # 数据处理器
        self.processor = WeiboDataProcessor()
    
    def configure_modern_styles(self):
        """配置现代化UI样式"""
        # 主色调
        primary_color = '#2196F3'  # 现代蓝色
        secondary_color = '#FFC107'  # 警告黄色
        success_color = '#4CAF50'  # 成功绿色
        error_color = '#F44336'  # 错误红色
        bg_color = '#FFFFFF'  # 背景白色
        card_color = '#FAFAFA'  # 卡片背景色
        
        # 配置标签样式
        self.style.configure('Title.TLabel', 
                           font=self.title_font, 
                           foreground=primary_color,
                           background='#f0f0f0')
        
        self.style.configure('Heading.TLabel', 
                           font=self.heading_font, 
                           foreground='#212121',
                           background='#f0f0f0')
        
        self.style.configure('Info.TLabel', 
                           font=self.default_font,
                           foreground='#424242',
                           background='#f0f0f0')
        
        self.style.configure('Success.TLabel', 
                           font=self.default_font,
                           foreground=success_color,
                           background='#f0f0f0')
        
        self.style.configure('Error.TLabel', 
                           font=self.default_font,
                           foreground=error_color,
                           background='#f0f0f0')
        
        # 配置按钮样式
        self.style.configure('Primary.TButton',
                           font=self.default_font,
                           foreground='white',
                           background=primary_color,
                           borderwidth=0,
                           focuscolor='none',
                           relief='flat')
        
        self.style.map('Primary.TButton',
                      background=[('active', '#1976D2'),
                                 ('pressed', '#0D47A1')])
        
        self.style.configure('Success.TButton',
                           font=self.default_font,
                           foreground='white',
                           background=success_color,
                           borderwidth=0,
                           focuscolor='none',
                           relief='flat')
        
        self.style.configure('Warning.TButton',
                           font=self.default_font,
                           foreground='white',
                           background=secondary_color,
                           borderwidth=0,
                           focuscolor='none',
                           relief='flat')
        
        # 配置Frame样式
        self.style.configure('TLabelFrame',
                           background=bg_color,
                           borderwidth=1,
                           relief='solid')
        
        self.style.configure('TLabelFrame.Label',
                           font=self.heading_font,
                           foreground=primary_color,
                           background=bg_color)
        
        # 配置Notebook样式
        self.style.configure('Modern.TNotebook',
                           background='#f0f0f0',
                           borderwidth=0)
        
        self.style.configure('Modern.TNotebook.Tab',
                           font=self.default_font,
                           padding=[20, 10],
                           background=card_color,
                           foreground='#424242')
        
        self.style.map('Modern.TNotebook.Tab',
                      background=[('selected', primary_color),
                                 ('active', '#E3F2FD')],
                      foreground=[('selected', 'white'),
                                 ('active', primary_color)])
        
        # 配置进度条样式
        self.style.configure('Modern.Horizontal.TProgressbar',
                           background=primary_color,
                           borderwidth=0,
                           lightcolor=primary_color,
                           darkcolor=primary_color)
    
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 创建左侧控制面板
        self.create_control_panel(main_frame)
        
        # 创建右侧数据展示区域
        self.create_data_panel(main_frame)
        
        # 创建底部状态栏
        self.create_status_bar(main_frame)
    
    def create_control_panel(self, parent):
        # 左侧控制面板
        control_frame = ttk.LabelFrame(parent, text="⚙️ 爬取设置", padding="15")
        control_frame.grid(row=0, column=0, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        control_frame.configure(width=350)
        
        # 标题
        title_label = ttk.Label(control_frame, text="微博数据爬虫", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 关键词输入
        ttk.Label(control_frame, text="🔍 搜索关键词:", style='Heading.TLabel').grid(row=1, column=0, sticky=tk.W, pady=(0, 8))
        self.keyword_var = tk.StringVar(value="python")
        keyword_entry = ttk.Entry(control_frame, textvariable=self.keyword_var, width=30, font=self.default_font)
        keyword_entry.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 页数设置
        ttk.Label(control_frame, text="📄 爬取页数:", style='Heading.TLabel').grid(row=3, column=0, sticky=tk.W, pady=(0, 8))
        self.pages_var = tk.IntVar(value=5)
        pages_frame = ttk.Frame(control_frame)
        pages_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 添加时间间隔设置
        ttk.Label(control_frame, text="⏱️ 请求间隔(秒):", style='Heading.TLabel').grid(row=5, column=0, sticky=tk.W, pady=(0, 8))
        self.delay_var = tk.DoubleVar(value=2.0)
        delay_frame = ttk.Frame(control_frame)
        delay_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        pages_scale = ttk.Scale(pages_frame, from_=1, to=20, variable=self.pages_var, orient=tk.HORIZONTAL)
        pages_scale.grid(row=0, column=0, sticky=(tk.W, tk.E))
        pages_frame.columnconfigure(0, weight=1)
        
        self.pages_label = ttk.Label(pages_frame, text="5页", style='Info.TLabel')
        self.pages_label.grid(row=0, column=1, padx=(15, 0))
        
        pages_scale.configure(command=self.update_pages_label)
        
        # 时间间隔滑块
        delay_scale = ttk.Scale(delay_frame, from_=1.0, to=10.0, variable=self.delay_var, orient=tk.HORIZONTAL)
        delay_scale.grid(row=0, column=0, sticky=(tk.W, tk.E))
        delay_frame.columnconfigure(0, weight=1)
        
        self.delay_label = ttk.Label(delay_frame, text="2.0秒", style='Info.TLabel')
        self.delay_label.grid(row=0, column=1, padx=(15, 0))
        
        delay_scale.configure(command=self.update_delay_label)
        
        # 开始爬取按钮
        self.crawl_button = ttk.Button(control_frame, text="🚀 开始爬取", command=self.start_crawling, style='Primary.TButton')
        self.crawl_button.grid(row=7, column=0, columnspan=2, pady=15, sticky=(tk.W, tk.E), ipady=8)
        
        # 停止爬取按钮
        self.stop_button = ttk.Button(control_frame, text="⏹️ 停止爬取", command=self.stop_crawling, state=tk.DISABLED, style='Warning.TButton')
        self.stop_button.grid(row=8, column=0, columnspan=2, pady=(0, 15), sticky=(tk.W, tk.E), ipady=8)
        
        # 进度条和状态
        progress_frame = ttk.Frame(control_frame)
        progress_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', style='Modern.Horizontal.TProgressbar')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.progress_label = ttk.Label(progress_frame, text="", style='Info.TLabel')
        self.progress_label.grid(row=1, column=0, sticky=tk.W)
        
        # 统计信息
        stats_frame = ttk.LabelFrame(control_frame, text="📊 数据统计", padding="15")
        stats_frame.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.stats_text = scrolledtext.ScrolledText(stats_frame, height=8, width=35, font=self.default_font, wrap=tk.WORD,
                                                  bg='#ffffff', fg='#333333', selectbackground='#2196F3')
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        
        # 导出按钮
        export_frame = ttk.Frame(control_frame)
        export_frame.grid(row=11, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=15)
        
        self.export_csv_button = ttk.Button(export_frame, text="📄 导出CSV", command=self.export_csv, state=tk.DISABLED, style='Success.TButton')
        self.export_csv_button.grid(row=0, column=0, padx=(0, 8), sticky=(tk.W, tk.E), ipady=5)
        
        self.export_excel_button = ttk.Button(export_frame, text="📊 导出Excel", command=self.export_excel, state=tk.DISABLED, style='Success.TButton')
        self.export_excel_button.grid(row=0, column=1, padx=(8, 0), sticky=(tk.W, tk.E), ipady=5)
        
        export_frame.columnconfigure(0, weight=1)
        export_frame.columnconfigure(1, weight=1)
        
        # 清空数据按钮
        self.clear_button = ttk.Button(control_frame, text="🗑️ 清空数据", command=self.clear_data, state=tk.DISABLED)
        self.clear_button.grid(row=12, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W, tk.E), ipady=5)
        
        control_frame.columnconfigure(0, weight=1)
    
    def create_data_panel(self, parent):
        # 右侧数据展示区域
        data_frame = ttk.LabelFrame(parent, text="📋 数据展示", padding="15")
        data_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(1, weight=1)
        
        # 创建Notebook标签页
        self.notebook = ttk.Notebook(data_frame, style='Modern.TNotebook')
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # 数据表格标签页
        self.create_table_tab()
        
        # 详细信息标签页
        self.create_detail_tab()
        
        # 筛选控制
        filter_frame = ttk.LabelFrame(data_frame, text="🔍 数据筛选", padding="15")
        filter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # 作者筛选
        ttk.Label(filter_frame, text="👤 作者:", style='Heading.TLabel').grid(row=0, column=0, padx=(0, 8), sticky=tk.W)
        self.author_var = tk.StringVar(value="全部")
        self.author_combo = ttk.Combobox(filter_frame, textvariable=self.author_var, state="readonly", width=18, font=self.default_font)
        self.author_combo.grid(row=0, column=1, padx=(0, 20))
        self.author_combo.bind('<<ComboboxSelected>>', self.filter_data)
        
        # 最小互动数筛选
        ttk.Label(filter_frame, text="💬 最小互动数:", style='Heading.TLabel').grid(row=0, column=2, padx=(0, 8), sticky=tk.W)
        self.min_engagement_var = tk.IntVar(value=0)
        engagement_spinbox = ttk.Spinbox(filter_frame, from_=0, to=10000, textvariable=self.min_engagement_var, width=12, font=self.default_font)
        engagement_spinbox.grid(row=0, column=3, padx=(0, 20))
        engagement_spinbox.bind('<Return>', self.filter_data)
        engagement_spinbox.bind('<FocusOut>', self.filter_data)
        
        # 刷新按钮
        refresh_button = ttk.Button(filter_frame, text="🔄 刷新筛选", command=self.filter_data, style='Primary.TButton')
        refresh_button.grid(row=0, column=4, ipady=3)
    
    def create_table_tab(self):
        # 创建数据表格标签页
        table_frame = ttk.Frame(self.notebook)
        self.notebook.add(table_frame, text="📊 数据表格")
        
        # 创建Treeview
        columns = ('作者', '内容', '发布时间', '转发数', '评论数', '点赞数')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # 设置列标题和样式
        for col in columns:
            self.tree.heading(col, text=col)
            if col in ['转发数', '评论数', '点赞数']:
                self.tree.column(col, width=90, anchor=tk.CENTER)
            elif col == '发布时间':
                self.tree.column(col, width=140, anchor=tk.CENTER)
            elif col == '作者':
                self.tree.column(col, width=120, anchor=tk.CENTER)
            else:
                self.tree.column(col, width=350)
        
        # 配置行样式
        self.tree.tag_configure('oddrow', background='#f8f9fa')
        self.tree.tag_configure('evenrow', background='#ffffff')
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # 布局
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # 双击打开链接
        self.tree.bind('<Double-1>', self.open_weibo_link)
    
    def create_detail_tab(self):
        # 创建详细信息标签页
        detail_frame = ttk.Frame(self.notebook)
        self.notebook.add(detail_frame, text="📝 详细信息")
        
        self.detail_text = scrolledtext.ScrolledText(detail_frame, font=self.default_font, wrap=tk.WORD, 
                                                   bg='#ffffff', fg='#333333', selectbackground='#2196F3')
        self.detail_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(0, weight=1)
    
    def create_status_bar(self, parent):
        # 底部状态栏
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, style='Info.TLabel')
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        # 添加帮助按钮
        help_button = ttk.Button(status_frame, text="❓ 使用帮助", command=self.show_help, style='Primary.TButton')
        help_button.grid(row=0, column=1, sticky=tk.E, ipady=3)
        
        status_frame.columnconfigure(0, weight=1)
    
    def update_pages_label(self, value):
        pages = int(float(value))
        self.pages_label.config(text=f"{pages}页")
    
    def update_delay_label(self, value):
        delay = float(value)
        self.delay_label.config(text=f"{delay:.1f}秒")
    
    def start_crawling(self):
        keyword = self.keyword_var.get().strip()
        if not keyword:
            messagebox.showerror("错误", "请输入搜索关键词！")
            return
        
        self.is_crawling = True
        self.crawl_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress.start()
        self.progress_label.config(text="准备开始爬取...")
        self.status_var.set(f"正在爬取关键词 '{keyword}' 的微博数据...")
        
        # 在新线程中执行爬取
        threading.Thread(target=self.crawl_data, daemon=True).start()
    
    def crawl_data(self):
        try:
            keyword = self.keyword_var.get().strip()
            max_pages = self.pages_var.get()
            delay = self.delay_var.get()
            
            # 更新进度状态
            self.root.after(0, lambda: self.progress_label.config(text=f"开始爬取，请求间隔: {delay:.1f}秒"))
            
            # 调用爬取函数（传入自定义延迟）
            df = get_weibo_list(keyword, max_pages, delay)
            
            if not self.is_crawling:  # 检查是否被停止
                return
            
            if not df.empty:
                self.df = df
                # 在主线程中更新UI
                self.root.after(0, self.update_display)
                self.root.after(0, lambda: self.progress_label.config(text="爬取完成"))
                self.root.after(0, lambda: self.status_var.set(f"✅ 成功爬取 {len(df)} 条微博数据！"))
            else:
                self.root.after(0, lambda: self.progress_label.config(text="未获取到数据"))
                self.root.after(0, lambda: self.status_var.set("⚠️ 未获取到数据，请尝试更换关键词"))
                self.root.after(0, lambda: messagebox.showwarning("警告", "未获取到数据，请尝试更换关键词或检查网络连接"))
                
        except Exception as e:
            if self.is_crawling:  # 只有在未被停止时才显示错误
                error_msg = f"爬取过程中出现错误：{str(e)}"
                self.root.after(0, lambda: self.progress_label.config(text="爬取出错"))
                self.root.after(0, lambda: self.status_var.set(f"❌ {error_msg}"))
                self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
        finally:
            self.root.after(0, self.finish_crawling)
    
    def stop_crawling(self):
        self.is_crawling = False
        self.finish_crawling()
        self.progress_label.config(text="爬取已停止")
        self.status_var.set("⏹️ 爬取已停止")
    
    def finish_crawling(self):
        self.is_crawling = False
        self.crawl_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress.stop()
    
    def update_display(self):
        """更新数据显示"""
        if self.df.empty:
            return
        
        # 更新表格
        self.update_table()
        
        # 更新统计信息
        self.update_stats()
        
        # 更新详细信息
        self.update_details()
        
        # 更新筛选选项
        self.update_filters()
        
        # 启用导出和清空按钮
        self.export_csv_button.config(state=tk.NORMAL)
        self.export_excel_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)
    
    def update_table(self):
        """更新数据表格"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加新数据
        for index, row in self.df.iterrows():
            content = str(row['微博内容'])[:50] + "..." if len(str(row['微博内容'])) > 50 else str(row['微博内容'])
            values = (
                row['微博作者'],
                content,
                row['发布时间'],
                row['转发数'],
                row['评论数'],
                row['点赞数']
            )
            # 交替行颜色
            row_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert('', tk.END, values=values, tags=(index, row_tag))
    
    def update_stats(self):
        """更新统计信息"""
        if self.df.empty:
            return
        
        stats = self.processor.get_basic_stats(self.df)
        
        stats_text = "📊 数据统计信息\n" + "="*30 + "\n\n"
        stats_text += f"总微博数: {stats.get('总微博数', 0)} 条\n\n"
        stats_text += f"总转发数: {stats.get('总转发数', 0):,}\n"
        stats_text += f"总评论数: {stats.get('总评论数', 0):,}\n"
        stats_text += f"总点赞数: {stats.get('总点赞数', 0):,}\n\n"
        stats_text += f"平均转发数: {stats.get('平均转发数', 0)}\n"
        stats_text += f"平均评论数: {stats.get('平均评论数', 0)}\n"
        stats_text += f"平均点赞数: {stats.get('平均点赞数', 0)}\n\n"
        stats_text += f"最热微博转发数: {stats.get('最热微博转发数', 0)}\n"
        stats_text += f"最热微博评论数: {stats.get('最热微博评论数', 0)}\n"
        stats_text += f"最热微博点赞数: {stats.get('最热微博点赞数', 0)}\n"
        
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, stats_text)
    
    def update_details(self):
        """更新详细信息"""
        if self.df.empty:
            return
        
        detail_text = "📝 微博详细信息\n" + "="*50 + "\n\n"
        
        for index, row in self.df.head(10).iterrows():  # 只显示前10条详细信息
            detail_text += f"【{index + 1}】作者: {row['微博作者']}\n"
            detail_text += f"发布时间: {row['发布时间']}\n"
            detail_text += f"内容: {row['微博内容']}\n"
            detail_text += f"互动数据: 转发 {row['转发数']} | 评论 {row['评论数']} | 点赞 {row['点赞数']}\n"
            if 'url' in row and pd.notna(row['url']):
                detail_text += f"链接: {row['url']}\n"
            detail_text += "-" * 50 + "\n\n"
        
        if len(self.df) > 10:
            detail_text += f"... 还有 {len(self.df) - 10} 条数据，请在表格中查看\n"
        
        self.detail_text.delete(1.0, tk.END)
        self.detail_text.insert(1.0, detail_text)
    
    def update_filters(self):
        """更新筛选选项"""
        if self.df.empty:
            return
        
        # 更新作者选择列表
        authors = ['全部'] + list(self.df['微博作者'].unique())
        self.author_combo['values'] = authors
        self.author_combo.set('全部')
    
    def filter_data(self, event=None):
        """筛选数据"""
        if self.df.empty:
            return
        
        filtered_df = self.df.copy()
        
        # 按作者筛选
        author = self.author_var.get()
        if author != '全部':
            filtered_df = filtered_df[filtered_df['微博作者'] == author]
        
        # 按互动数筛选
        min_engagement = self.min_engagement_var.get()
        if min_engagement > 0:
            total_engagement = filtered_df['转发数'] + filtered_df['评论数'] + filtered_df['点赞数']
            filtered_df = filtered_df[total_engagement >= min_engagement]
        
        # 更新表格显示
        self.update_filtered_table(filtered_df)
    
    def update_filtered_table(self, df):
        """更新筛选后的表格"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加筛选后的数据
        for display_index, (index, row) in enumerate(df.iterrows()):
            content = str(row['微博内容'])[:50] + "..." if len(str(row['微博内容'])) > 50 else str(row['微博内容'])
            values = (
                row['微博作者'],
                content,
                row['发布时间'],
                row['转发数'],
                row['评论数'],
                row['点赞数']
            )
            # 交替行颜色
            row_tag = 'evenrow' if display_index % 2 == 0 else 'oddrow'
            self.tree.insert('', tk.END, values=values, tags=(index, row_tag))
    
    def open_weibo_link(self, event):
        """双击打开微博链接"""
        if not self.tree.selection():
            return
        item = self.tree.selection()[0]
        tags = self.tree.item(item, 'tags')
        if tags:
            row_index = int(tags[0])
            if 'url' in self.df.columns and row_index < len(self.df):
                url = self.df.iloc[row_index]['url']
                if pd.notna(url):
                    webbrowser.open(url)
    
    def export_csv(self):
        """导出CSV文件"""
        if self.df.empty:
            messagebox.showwarning("警告", "没有数据可导出！")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialname=f"微博数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if filename:
            try:
                self.df.to_csv(filename, index=False, encoding='utf-8-sig')
                messagebox.showinfo("成功", f"数据已导出到: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def export_excel(self):
        """导出Excel文件"""
        if self.df.empty:
            messagebox.showwarning("警告", "没有数据可导出！")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialname=f"微博数据分析_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                self.df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"数据已导出到: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def clear_data(self):
        """清空数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.df = pd.DataFrame()
            
            # 清空表格
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 清空文本区域
            self.stats_text.delete(1.0, tk.END)
            self.detail_text.delete(1.0, tk.END)
            
            # 禁用按钮
            self.export_csv_button.config(state=tk.DISABLED)
            self.export_excel_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)
            
            # 重置筛选
            self.author_combo['values'] = ['全部']
            self.author_combo.set('全部')
            self.min_engagement_var.set(0)
            
            self.status_var.set("数据已清空")
    
    def show_help(self):
        """显示帮助信息"""
        help_text = """
📋 基本使用流程：
1. 在左侧输入搜索关键词（如：python、人工智能）
2. 调整爬取页数（1-20页，建议5-10页）
3. 设置请求间隔（1-10秒，避免对服务器压力过大）
4. 点击"🚀 开始爬取"按钮
5. 查看实时进度和统计信息
6. 可以筛选、导出所需数据

⚙️ 参数设置：
- 🔍 搜索关键词：支持中文、英文、话题等
- 📄 爬取页数：每页约10-20条微博，建议不超过20页
- ⏱️ 请求间隔：默认2秒，网络较差时可增加到5-10秒

🔍 数据筛选功能：
- 👤 按作者筛选：查看特定用户的所有微博
- 💬 按互动数筛选：筛选热门微博（转发+评论+点赞）
- 🔄 实时筛选：立即预览筛选结果

📊 数据查看方式：
- 📋 表格页面：查看所有微博的核心信息
- 📝 详细信息：查看前10条微博的完整内容
- 🔗 双击表格行：直接打开微博链接

💾 数据导出：
- 📄 CSV格式：适合Excel、数据分析工具
- 📊 Excel格式：包含多个工作表和统计信息
- 🕒 自动命名：文件名包含时间戳，避免覆盖

📈 统计信息：
- 总体数据：微博总数、互动数统计
- 平均数据：平均转发、评论、点赞数
- 热门数据：最热微博的各项指标

⚠️ 重要提醒：
- 🕒 合理设置间隔：避免请求过于频繁
- 📊 适量爬取：建议每次不超过20页
- ⚖️ 合法使用：仅供学习研究，遵守法律法规
- 🌐 网络环境：确保网络连接稳定

💡 使用技巧：
- 关键词可以是话题、用户名、热点事件等
- 如遇到爬取失败，可尝试增加请求间隔
- 数据可以多次筛选，找到最需要的内容
- 导出前可先筛选，减少无用数据
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("📖 使用帮助")
        help_window.geometry("600x700")
        help_window.resizable(True, True)
        help_window.configure(bg='#f0f0f0')
        
        # 创建主框架
        main_frame = ttk.Frame(help_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="🐦 微博数据爬虫分析平台 - 使用帮助", style='Title.TLabel')
        title_label.pack(pady=(0, 20))
        
        # 帮助内容
        help_text_widget = scrolledtext.ScrolledText(main_frame, font=self.default_font, wrap=tk.WORD,
                                                   bg='#ffffff', fg='#333333', selectbackground='#2196F3')
        help_text_widget.pack(fill=tk.BOTH, expand=True)
        help_text_widget.insert(1.0, help_text)
        help_text_widget.config(state=tk.DISABLED)
        
        # 关闭按钮
        close_button = ttk.Button(main_frame, text="✅ 知道了", command=help_window.destroy, style='Primary.TButton')
        close_button.pack(pady=(15, 0), ipady=5)

def main():
    root = tk.Tk()
    app = WeiboSpiderGUI(root)
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main() 