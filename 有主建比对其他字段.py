import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import json
import threading
import re

"""
版本号:1.6
版本说明:
这个UI版1.5是用来对比两个Excel文件的，主要对比字段是主键字段，其他字段是对比字段，对比字段是多个字段用逗号分隔。
对比结果会输出到Excel文件中，输出文件名是对比结果.xlsx。

V1.5 增加了对于日期格式的处理，对于日期格式会进行标准化处理。
V1.6 修复了列名类型错误。

"""


class DataComparisonTool:
    def __init__(self, root):
        self.root = root
        self.root.title("数据对比小工具 v1.6")
        self.root.geometry("900x700")
        
        # 设置进度条样式
        self.setup_styles()
        
        # 数据存储
        self.df1 = None
        self.df2 = None
        self.file1_path = ""
        self.file2_path = ""
        self.sheet1_name = ""
        self.sheet2_name = ""
        
        self.setup_ui()
        
    def setup_styles(self):
        """设置进度条样式"""
        style = ttk.Style()
        
        # 红色进度条（0-30%）
        style.configure("red.Horizontal.TProgressbar", 
                       background="red", 
                       troughcolor="lightgray",
                       borderwidth=0,
                       lightcolor="red",
                       darkcolor="red")
        
        # 黄色进度条（30-70%）
        style.configure("yellow.Horizontal.TProgressbar", 
                       background="orange", 
                       troughcolor="lightgray",
                       borderwidth=0,
                       lightcolor="orange",
                       darkcolor="orange")
        
        # 绿色进度条（70-100%）
        style.configure("green.Horizontal.TProgressbar", 
                       background="green", 
                       troughcolor="lightgray",
                       borderwidth=0,
                       lightcolor="green",
                       darkcolor="green")
        
    def setup_ui(self):
        # 配置文件区域
        config_frame = ttk.LabelFrame(self.root, text="配置文件", padding="10")
        config_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(config_frame, text="配置文件路径:").grid(row=0, column=0, sticky="w", pady=2)
        self.config_entry = ttk.Entry(config_frame, width=60)
        self.config_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(config_frame, text="浏览", command=self.browse_config).grid(row=0, column=2, pady=2)
        ttk.Button(config_frame, text="加载配置", command=self.load_config).grid(row=0, column=3, pady=2)
        ttk.Button(config_frame, text="生成模板", command=self.generate_config_template).grid(row=0, column=4, pady=2)
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="文件选择", padding="10")
        file_frame.pack(fill="x", padx=10, pady=5)
        
        # 文件1
        ttk.Label(file_frame, text="文件1:").grid(row=0, column=0, sticky="w", pady=2)
        self.file1_entry = ttk.Entry(file_frame, width=50)
        self.file1_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="浏览", command=self.browse_file1).grid(row=0, column=2, pady=2)
        
        ttk.Label(file_frame, text="工作表:").grid(row=1, column=0, sticky="w", pady=2)
        self.sheet1_entry = ttk.Entry(file_frame, width=20)
        self.sheet1_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        self.sheet1_entry.insert(0, "默认")
        
        # 文件2
        ttk.Label(file_frame, text="文件2:").grid(row=2, column=0, sticky="w", pady=2)
        self.file2_entry = ttk.Entry(file_frame, width=50)
        self.file2_entry.grid(row=2, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="浏览", command=self.browse_file2).grid(row=2, column=2, pady=2)
        
        ttk.Label(file_frame, text="工作表:").grid(row=3, column=0, sticky="w", pady=2)
        self.sheet2_entry = ttk.Entry(file_frame, width=20)
        self.sheet2_entry.grid(row=3, column=1, padx=5, pady=2, sticky="w")
        self.sheet2_entry.insert(0, "默认")
        
        # 主键配置区域
        key_frame = ttk.LabelFrame(self.root, text="主键配置", padding="10")
        key_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(key_frame, text="主键字段（多个字段用逗号分隔）:").grid(row=0, column=0, sticky="w", pady=2)
        self.key_entry = ttk.Entry(key_frame, width=50)
        self.key_entry.grid(row=0, column=1, padx=5, pady=2)
        self.key_entry.insert(0, "税号,客商编码")
        
        # 字段映射配置区域
        mapping_frame = ttk.LabelFrame(self.root, text="字段映射配置", padding="10")
        mapping_frame.pack(fill="x", padx=10, pady=5)
        
        # 字段映射表格
        ttk.Label(mapping_frame, text="字段映射（文件1字段名:文件2字段名，多个用逗号分隔）:").grid(row=0, column=0, sticky="w", pady=2)
        self.mapping_entry = ttk.Entry(mapping_frame, width=50)
        self.mapping_entry.grid(row=0, column=1, padx=5, pady=2)
        self.mapping_entry.insert(0, "公司名称:公司名称,联系人:联系人,电话:电话")
        
        # 输出配置区域
        output_frame = ttk.LabelFrame(self.root, text="输出配置", padding="10")
        output_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(output_frame, text="输出目录:").grid(row=0, column=0, sticky="w", pady=2)
        self.output_entry = ttk.Entry(output_frame, width=50)
        self.output_entry.grid(row=0, column=1, padx=5, pady=2)
        self.output_entry.insert(0, os.path.dirname(os.path.abspath(__file__)))
        ttk.Button(output_frame, text="浏览", command=self.browse_output).grid(row=0, column=2, pady=2)
        
        # 操作按钮区域
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        self.load_button = ttk.Button(button_frame, text="加载数据", command=self.load_data)
        self.load_button.pack(side="left", padx=5)
        
        self.compare_button = ttk.Button(button_frame, text="开始比对", command=self.start_comparison, state="disabled")
        self.compare_button.pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="清空", command=self.clear_all).pack(side="left", padx=5)
        
        # 进度条区域
        progress_frame = ttk.LabelFrame(self.root, text="进度信息", padding="10")
        progress_frame.pack(fill="x", padx=10, pady=5)
        
        self.progress_var = tk.StringVar(value="准备就绪")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(anchor="w")
        
        # 添加百分比显示
        self.progress_percent = tk.StringVar(value="0%")
        ttk.Label(progress_frame, textvariable=self.progress_percent).pack(anchor="w")
        
        # 进度条改为确定模式，支持百分比
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', maximum=100)
        self.progress_bar.pack(fill="x", pady=5)
        
        # 添加详细步骤显示
        self.step_var = tk.StringVar(value="等待开始...")
        ttk.Label(progress_frame, textvariable=self.step_var, font=("Arial", 9)).pack(anchor="w")
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(self.root, text="状态信息", padding="10")
        status_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.status_text = tk.Text(status_frame, height=15, wrap="word")
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.status_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def browse_config(self):
        filename = filedialog.askopenfilename(
            title="选择配置文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.config_entry.delete(0, tk.END)
            self.config_entry.insert(0, filename)
            
    def generate_config_template(self):
        """生成配置文件模板"""
        try:
            from 生成配置文件 import create_config_template
            output_file = create_config_template()
            messagebox.showinfo("成功", f"配置文件模板已生成:\n{output_file}")
            self.config_entry.delete(0, tk.END)
            self.config_entry.insert(0, os.path.abspath(output_file))
        except ImportError:
            messagebox.showerror("错误", "无法导入配置文件生成模块")
        except Exception as e:
            messagebox.showerror("错误", f"生成配置文件失败: {str(e)}")
            
    def load_config(self):
        """从配置文件加载配置"""
        config_file = self.config_entry.get().strip()
        if not config_file:
            messagebox.showerror("错误", "请先选择配置文件")
            return
            
        try:
            self.log_message("正在加载配置文件...")
            
            # 读取文件路径配置
            file_config = pd.read_excel(config_file, sheet_name='文件路径配置')
            file1_path = file_config.iloc[0]['Excel1文件绝对路径']
            file2_path = file_config.iloc[0]['Excel2文件绝对路径']
            sheet1_name = file_config.iloc[0]['Excel1文件sheet名称']
            sheet2_name = file_config.iloc[0]['Excel2文件sheet名称']
            
            # 读取输出路径配置（新增字段）
            try:
                output_path = file_config.iloc[0]['对比结果输出绝对路径']
                if pd.notna(output_path) and str(output_path).strip():
                    self.output_entry.delete(0, tk.END)
                    self.output_entry.insert(0, str(output_path).strip())
            except:
                # 如果没有输出路径字段，保持原有输出路径不变
                pass
            
            # 读取主键配置
            key_config = pd.read_excel(config_file, sheet_name='主键配置')
            key_config = key_config.dropna(subset=['Excel1主键字段名称', 'Excel2主键字段名称'])
            keys = [f"{row['Excel1主键字段名称']}:{row['Excel2主键字段名称']}" for _, row in key_config.iterrows()]
            
            # 读取对比字段配置
            try:
                compare_config = pd.read_excel(config_file, sheet_name='对比字段配置')
                compare_config = compare_config.dropna(subset=['Excel1对比字段名称', 'Excel2对比字段名称'])
                mappings = [f"{row['Excel1对比字段名称']}:{row['Excel2对比字段名称']}" for _, row in compare_config.iterrows()]
            except:
                mappings = []
            
            # 应用配置
            if file1_path and pd.notna(file1_path):
                self.file1_entry.delete(0, tk.END)
                self.file1_entry.insert(0, str(file1_path))
                
            if file2_path and pd.notna(file2_path):
                self.file2_entry.delete(0, tk.END)
                self.file2_entry.insert(0, str(file2_path))
                
            if sheet1_name and pd.notna(sheet1_name):
                self.sheet1_entry.delete(0, tk.END)
                self.sheet1_entry.insert(0, str(sheet1_name))
            else:
                self.sheet1_entry.delete(0, tk.END)
                self.sheet1_entry.insert(0, "默认")
                
            if sheet2_name and pd.notna(sheet2_name):
                self.sheet2_entry.delete(0, tk.END)
                self.sheet2_entry.insert(0, str(sheet2_name))
            else:
                self.sheet2_entry.delete(0, tk.END)
                self.sheet2_entry.insert(0, "默认")
                
            if keys:
                self.key_entry.delete(0, tk.END)
                self.key_entry.insert(0, ','.join(keys))
                
            if mappings:
                self.mapping_entry.delete(0, tk.END)
                self.mapping_entry.insert(0, ','.join(mappings))
                
            self.log_message("配置文件加载成功")
            self.log_message(f"文件1路径: {file1_path}")
            self.log_message(f"文件2路径: {file2_path}")
            self.log_message(f"工作表1: {sheet1_name if pd.notna(sheet1_name) else '默认'}")
            self.log_message(f"工作表2: {sheet2_name if pd.notna(sheet2_name) else '默认'}")
            self.log_message(f"主键字段: {keys}")
            self.log_message(f"对比字段: {mappings}")
            messagebox.showinfo("成功", "配置文件加载成功！")
            
        except Exception as e:
            self.log_message(f"加载配置文件失败: {str(e)}")
            messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
        
    def browse_file1(self):
        filename = filedialog.askopenfilename(
            title="选择文件1",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, filename)
            
    def browse_file2(self):
        filename = filedialog.askopenfilename(
            title="选择文件2",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, filename)
            
    def browse_output(self):
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, dirname)
            
    def log_message(self, message):
        self.status_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.status_text.see(tk.END)
        self.root.update()
        
    def update_progress(self, message, percent=0, step=""):
        """更新进度信息"""
        self.progress_var.set(message)
        self.progress_percent.set(f"{percent}%")
        if step:
            self.step_var.set(step)
        
        # 更新进度条
        self.progress_bar['value'] = percent
        
        # 根据进度设置进度条颜色
        if percent < 30:
            self.progress_bar['style'] = 'red.Horizontal.TProgressbar'
        elif percent < 70:
            self.progress_bar['style'] = 'yellow.Horizontal.TProgressbar'
        else:
            self.progress_bar['style'] = 'green.Horizontal.TProgressbar'
            
        self.root.update()
        
    def load_data(self):
        try:
            file1 = self.file1_entry.get().strip()
            file2 = self.file2_entry.get().strip()
            
            if not file1 or not file2:
                messagebox.showerror("错误", "请选择两个文件")
                return
                
            # 获取工作表名称
            sheet1 = self.sheet1_entry.get().strip()
            sheet2 = self.sheet2_entry.get().strip()
            
            # 如果工作表名称为"默认"，则使用0（第一个工作表）
            sheet1 = 0 if sheet1 == "默认" else sheet1
            sheet2 = 0 if sheet2 == "默认" else sheet2
            
            self.update_progress("正在加载文件1...", 20, "读取文件1数据...")
            self.log_message("正在加载文件1...")
            
            # 读取文件时指定某些列为字符串类型，防止长数字被截断
            self.df1 = pd.read_excel(file1, sheet_name=sheet1, dtype=str)
            
            # 标准化日期格式
            self.df1 = self.standardize_date_columns(self.df1, "文件1")
            
            self.log_message(f"文件1加载成功，共 {len(self.df1)} 行，{len(self.df1.columns)} 列")
            self.log_message(f"文件1列名: {list(self.df1.columns)}")
            
            self.update_progress("正在加载文件2...", 40, "读取文件2数据...")
            self.log_message("正在加载文件2...")
            
            self.df2 = pd.read_excel(file2, sheet_name=sheet2, dtype=str)
            
            # 标准化日期格式
            self.df2 = self.standardize_date_columns(self.df2, "文件2")
            
            self.log_message(f"文件2加载成功，共 {len(self.df2)} 行，{len(self.df2.columns)} 列")
            self.log_message(f"文件2列名: {list(self.df2.columns)}")
            
            self.file1_path = file1
            self.file2_path = file2
            self.sheet1_name = sheet1
            self.sheet2_name = sheet2
            
            self.update_progress("数据加载完成", 50, "数据加载完成，可以开始比对")
            self.compare_button.config(state="normal")
            messagebox.showinfo("成功", "数据加载完成！")
            
        except Exception as e:
            self.update_progress("数据加载失败", 0, f"发生错误: {str(e)}")
            self.log_message(f"加载数据失败: {str(e)}")
            messagebox.showerror("错误", f"加载数据失败: {str(e)}")
            

    def standardize_date_columns(self, df, file_name):
        """标准化DataFrame中的日期列格式，只保留日期部分"""
        try:
            df_copy = df.copy()
            date_columns = []
            
            for col in df_copy.columns:
                # 检查列是否包含日期格式的数据
                if self.is_date_column(df_copy[col]):
                    date_columns.append(col)
                    self.log_message(f"{file_name} 发现日期列: {col}")
            
            if date_columns:
                self.log_message(f"{file_name} 正在标准化 {len(date_columns)} 个日期列...")
                
                for col in date_columns:
                    # 将日期转换为标准格式 YYYY-MM-DD
                    df_copy[col] = df_copy[col].apply(self.normalize_date)
                    self.log_message(f"  {col}: 已标准化为 YYYY-MM-DD 格式")
            
            return df_copy
            
        except Exception as e:
            self.log_message(f"标准化日期列时出错: {str(e)}")
            return df  # 如果出错，返回原始数据
    
    def is_date_column(self, series):
        """判断列是否包含日期数据"""
        try:
            # 检查前几行数据是否包含日期格式
            sample_size = min(10, len(series))
            sample_data = series.dropna().head(sample_size)
            
            if len(sample_data) == 0:
                return False
            
            date_count = 0
            for value in sample_data:
                if self.looks_like_date(str(value)):
                    date_count += 1
            
            # 如果超过50%的样本数据看起来像日期，则认为是日期列
            return date_count / len(sample_data) > 0.5
            
        except:
            return False
    
    def looks_like_date(self, value):
        """判断字符串是否看起来像日期"""
        try:
            # 常见的日期格式模式
            date_patterns = [
                r'\d{4}[-/]\d{1,2}[-/]\d{1,2}',  # YYYY-MM-DD 或 YYYY/MM/DD
                r'\d{1,2}[-/]\d{1,2}[-/]\d{4}',  # MM-DD-YYYY 或 MM/DD/YYYY
                r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}',  # YYYY-MM-DD HH:MM:SS
                r'\d{4}/\d{1,2}/\d{1,2}\s+\d{1,2}:\d{1,2}:\d{1,2}',  # YYYY/MM/DD HH:MM:SS
            ]
            
            for pattern in date_patterns:
                if re.match(pattern, str(value).strip()):
                    return True
            return False
        except:
            return False
    
    def normalize_date(self, value):
        """将日期值标准化为 YYYY-MM-DD 格式"""
        try:
            if pd.isna(value) or value == '' or str(value).strip() == '':
                return value
            
            value_str = str(value).strip()
            
            # 如果已经是 YYYY-MM-DD 格式，直接返回
            if re.match(r'^\d{4}-\d{2}-\d{2}$', value_str):
                return value_str
            
            # 处理带时间戳的日期
            if ' ' in value_str:
                date_part = value_str.split(' ')[0]
                # 将 / 替换为 -
                if '/' in date_part:
                    date_part = date_part.replace('/', '-')
                return date_part
            
            # 处理 YYYY/MM/DD 格式
            if '/' in value_str:
                parts = value_str.split('/')
                if len(parts) == 3:
                    year, month, day = parts
                    # 确保是4位年份
                    if len(year) == 2:
                        year = '20' + year if int(year) < 50 else '19' + year
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            
            # 处理 MM/DD/YYYY 格式
            if '/' in value_str and len(value_str.split('/')) == 3:
                parts = value_str.split('/')
                if len(parts[2]) == 4:  # 最后一部分是4位年份
                    month, day, year = parts
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            
            # 如果无法识别，返回原值
            return value_str
            
        except Exception as e:
            self.log_message(f"标准化日期时出错 '{value}': {str(e)}")
            return value

    def parse_keys_and_mappings(self):
        """解析主键和字段映射"""
        try:
            # 解析主键
            keys_str = self.key_entry.get().strip()
            if not keys_str:
                raise ValueError("主键字段不能为空")
            
            keys = []
            key_mappings = {}  # 存储主键字段的映射关系
            
            for key_mapping in keys_str.split(","):
                if ":" in key_mapping:
                    field1, field2 = key_mapping.split(":", 1)
                    field1 = field1.strip()
                    field2 = field2.strip()
                    keys.append(field1)  # 使用Excel1的字段名作为主键
                    key_mappings[field1] = field2  # 记录映射关系
                else:
                    field = key_mapping.strip()
                    keys.append(field)
                    key_mappings[field] = field  # 相同字段名
            
            # 解析字段映射
            mapping_str = self.mapping_entry.get().strip()
            mappings = {}
            
            if mapping_str:
                for mapping in mapping_str.split(","):
                    if ":" in mapping:
                        field1, field2 = mapping.split(":", 1)
                        mappings[field1.strip()] = field2.strip()
                    else:
                        # 如果没有冒号，假设字段名相同
                        field = mapping.strip()
                        mappings[field] = field
                        
            return keys, mappings, key_mappings
            
        except Exception as e:
            raise ValueError(f"解析配置失败: {str(e)}")
            
    def start_comparison(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showerror("错误", "请先加载数据")
            return
            
        try:
            keys, mappings, key_mappings = self.parse_keys_and_mappings()
            self.log_message(f"主键字段: {keys}")
            self.log_message(f"字段映射: {mappings}")
            self.log_message(f"主键映射关系: {key_mappings}")
            
            # 检查字段是否存在
            for key in keys:
                if key not in self.df1.columns:
                    raise ValueError(f"文件1中缺少主键字段: {key}")
                # 检查对应的Excel2字段是否存在
                excel2_key = key_mappings.get(key, key)
                if excel2_key not in self.df2.columns:
                    raise ValueError(f"文件2中缺少主键字段: {excel2_key}")
                    
            for field1, field2 in mappings.items():
                if field1 not in self.df1.columns:
                    raise ValueError(f"文件1中缺少字段: {field1}")
                if field2 not in self.df2.columns:
                    raise ValueError(f"文件2中缺少字段: {field2}")
                    
            # 在新线程中执行比对，避免界面卡死
            self.compare_button.config(state="disabled")
            self.load_button.config(state="disabled")
            
            thread = threading.Thread(target=self.perform_comparison_thread, args=(keys, mappings, key_mappings))
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            self.log_message(f"比对失败: {str(e)}")
            messagebox.showerror("错误", f"比对失败: {str(e)}")
            
    def perform_comparison_thread(self, keys, mappings, key_mappings):
        """在线程中执行比对"""
        try:
            self.update_progress("开始执行数据比对...", 60, "准备数据合并...")
            self.log_message("开始执行数据比对...")
            
            # 准备合并用的字段映射
            # 由于主键字段名可能不同，我们需要创建临时的DataFrame进行合并
            df1_temp = self.df1.copy()
            df2_temp = self.df2.copy()
            
            self.update_progress("正在处理字段映射...", 65, "重命名列和更新映射关系...")
            
            # 重命名Excel2的列，使其与Excel1的主键字段名一致
            rename_dict = {}
            for excel1_key, excel2_key in key_mappings.items():
                if excel1_key != excel2_key:
                    rename_dict[excel2_key] = excel1_key
            
            if rename_dict:
                df2_temp = df2_temp.rename(columns=rename_dict)
                self.log_message(f"重命名Excel2列: {rename_dict}")
                
                # 同时更新字段映射中的字段名
                updated_mappings = {}
                for field1, field2 in mappings.items():
                    # 如果字段2在重命名字典中，需要更新为新的字段名
                    if field2 in rename_dict:
                        updated_mappings[field1] = rename_dict[field2]
                    else:
                        updated_mappings[field1] = field2
                mappings = updated_mappings
                self.log_message(f"更新后的字段映射: {mappings}")
            
            self.update_progress("正在合并数据...", 70, "执行数据合并操作...")
            
            # 为所有字段添加文件标识，确保字段名称清晰
            # 重命名文件1的所有非主键字段
            df1_rename = {}
            for col in df1_temp.columns:
                if col not in keys:
                    # 确保列名是字符串类型
                    col_str = str(col)
                    df1_rename[col] = col_str + '_文件1'
            
            # 重命名文件2的所有非主键字段
            df2_rename = {}
            for col in df2_temp.columns:
                if col not in keys:
                    # 确保列名是字符串类型
                    col_str = str(col)
                    df2_rename[col] = col_str + '_文件2'
            
            # 应用重命名
            if df1_rename:
                df1_temp = df1_temp.rename(columns=df1_rename)
                self.log_message(f"文件1字段重命名: {df1_rename}")
            
            if df2_rename:
                df2_temp = df2_temp.rename(columns=df2_rename)
                self.log_message(f"文件2字段重命名: {df2_rename}")
            
            # 确保所有列名都是字符串类型
            df1_temp.columns = [str(col) for col in df1_temp.columns]
            df2_temp.columns = [str(col) for col in df2_temp.columns]
            
            # 以主键合并数据（现在所有非主键字段都有文件标识）
            merged = pd.merge(df1_temp, df2_temp, how='outer', on=keys, 
                             suffixes=('', ''), indicator=True)
            
            self.log_message(f"合并后总记录数: {len(merged)}")
            self.log_message(f"合并后的列名: {list(merged.columns)}")
            
            # 显示前几行数据用于调试
            self.log_message("合并后前3行数据预览:")
            for i, row in merged.head(3).iterrows():
                self.log_message(f"  行{i}: {dict(row)}")
            
            self.update_progress("正在重组列名...", 75, "优化列名显示...")
            
            # 重新组织列名，使其更清晰
            self.log_message("列名重组，明确标识字段来源:")
            column_mapping = {}
            for col in merged.columns:
                if col == '_merge':
                    continue
                elif col.endswith('_文件1'):
                    # 保持原有的_文件1后缀
                    column_mapping[col] = col
                elif col.endswith('_文件2'):
                    # 保持原有的_文件2后缀
                    column_mapping[col] = col
                elif col in keys:
                    # 主键字段，同时存在于两个文件中，保持原名
                    column_mapping[col] = col
                else:
                    # 其他字段保持原名
                    column_mapping[col] = col
            
            # 重命名列
            merged = merged.rename(columns=column_mapping)
            self.log_message(f"重组后的列名: {list(merged.columns)}")
            self.log_message(f"主键字段: {keys}")
            self.log_message(f"文件1字段: {[col for col in merged.columns if col.endswith('_文件1')]}")
            self.log_message(f"文件2字段: {[col for col in merged.columns if col.endswith('_文件2')]}")
            
            # 分类比对结果
            self.update_progress("正在分类比对结果...", 80, "分析数据差异...")
            results, field_comparison_results = self.categorize_results(merged, keys, mappings)
            
            # 生成报告
            self.update_progress("正在生成比对报告...", 90, "写入Excel文件...")
            self.generate_report(results, field_comparison_results, keys, mappings, key_mappings)
            
            self.update_progress("比对完成", 100, "所有操作已完成！")
            
        except Exception as e:
            self.log_message(f"比对执行失败: {str(e)}")
            self.update_progress("比对失败", 0, f"发生错误: {str(e)}")
            
        finally:
            # 恢复按钮状态
            self.root.after(0, lambda: self.compare_button.config(state="normal"))
            self.root.after(0, lambda: self.load_button.config(state="normal"))
        
    def categorize_results(self, merged, keys, mappings):
        """分类比对结果"""
        results = {}
        field_comparison_results = []  # 新增：记录每个字段的比对结果
        
        # 完全匹配：主键一致，所有映射字段都一致
        mask_full_match = (merged['_merge'] == 'both')
        field_masks = {}  # 记录每个字段的匹配掩码
        
        # 检查是否有映射字段需要比对
        if mappings:
            for field1, field2 in mappings.items():
                col1 = field1 + '_文件1'
                col2 = field2 + '_文件2'
                
                # 检查列是否存在，现在所有字段都有文件标识
                actual_col1 = str(field1) + '_文件1'
                actual_col2 = str(field2) + '_文件2'
                
                # 验证列是否存在
                if actual_col1 not in merged.columns:
                    self.log_message(f"警告: 文件1字段 {actual_col1} 不存在")
                    actual_col1 = None
                if actual_col2 not in merged.columns:
                    self.log_message(f"警告: 文件2字段 {actual_col2} 不存在")
                    actual_col2 = None
                
                if actual_col1 and actual_col2:
                    # 处理NaN值，避免比较错误
                    mask_field_match = (merged[actual_col1] == merged[actual_col2]) | \
                                     ((merged[actual_col1].isna()) & (merged[actual_col2].isna()))
                    field_masks[field1] = mask_field_match
                    mask_full_match &= mask_field_match
                    
                    # 记录字段比对统计
                    both_records = (merged['_merge'] == 'both')  # 两个文件都有的记录
                    field_match_count = (both_records & mask_field_match).sum()
                    field_mismatch_count = (both_records & ~mask_field_match).sum()
                    total_both_count = both_records.sum()
                    
                    field_comparison_results.append({
                        '字段名称': f"{field1} ↔ {field2}",
                        '文件1字段': field1,
                        '文件2字段': field2,
                        '匹配数量': field_match_count,
                        '不匹配数量': field_mismatch_count,
                        '两文件共有记录数': total_both_count,
                        '匹配率': f"{field_match_count/total_both_count*100:.2f}%" if total_both_count > 0 else "0%"
                    })
                    
                    self.log_message(f"字段比对: {field1}↔{field2} (实际列: {actual_col1}↔{actual_col2}), 匹配: {field_match_count}, 不匹配: {field_mismatch_count}")
                else:
                    self.log_message(f"警告: 字段 {field1}({col1}) 或 {field2}({col2}) 在合并后的数据中不存在，跳过该字段比对")
                    self.log_message(f"  合并后的列名: {list(merged.columns)}")
                    field_comparison_results.append({
                        '字段名称': f"{field1} ↔ {field2}",
                        '文件1字段': field1,
                        '文件2字段': field2,
                        '匹配数量': 0,
                        '不匹配数量': 0,
                        '两文件共有记录数': 0,
                        '匹配率': "字段不存在"
                    })
        else:
            self.log_message("没有配置对比字段，仅按主键匹配")
        
        results['完全匹配'] = merged[mask_full_match]
        
        # 不一致：主键一致，但某些映射字段不一致
        mask_disagree = (merged['_merge'] == 'both') & (~mask_full_match)
        results['不一致'] = merged[mask_disagree]
        
        # 仅在文件1
        results['仅在文件1'] = merged[merged['_merge'] == 'left_only']
        
        # 仅在文件2
        results['仅在文件2'] = merged[merged['_merge'] == 'right_only']
        
        # 为不一致的记录添加字段级别的比对结果
        if not results['不一致'].empty and mappings:
            disagree_data = results['不一致'].copy()
            
            # 为每个记录添加字段比对状态列
            for field1, field2 in mappings.items():
                col1 = field1 + '_文件1'
                col2 = field2 + '_文件2'
                status_col = f"{str(field1)}_比对状态"
                
                # 现在所有字段都有文件标识，直接使用标准格式
                actual_col1 = field1 + '_文件1'
                actual_col2 = field2 + '_文件2'
                
                if actual_col1 and actual_col2:
                    # 创建比对状态列
                    field_mask = field_masks.get(field1)
                    if field_mask is not None:
                        # 只对不一致的记录创建状态
                        status_values = []
                        
                        for idx in disagree_data.index:
                            if field_mask.loc[idx]:
                                status_values.append("匹配")
                            else:
                                val1 = disagree_data.loc[idx, actual_col1]
                                val2 = disagree_data.loc[idx, actual_col2]
                                if pd.isna(val1) and pd.isna(val2):
                                    status_values.append("匹配(都为空)")
                                elif pd.isna(val1):
                                    status_values.append("不匹配(文件1为空)")
                                elif pd.isna(val2):
                                    status_values.append("不匹配(文件2为空)")
                                else:
                                    status_values.append(f"不匹配({val1}≠{val2})")
                        
                        disagree_data[status_col] = status_values
                    else:
                        disagree_data[status_col] = "字段不存在"
            
            results['不一致'] = disagree_data
        
        # 记录统计
        for category, data in results.items():
            self.log_message(f"{category}: {len(data)} 条记录")
        
        # 返回结果和字段比对统计
        return results, field_comparison_results
        
    def generate_report(self, results, field_comparison_results, keys, mappings, key_mappings):
        """生成比对报告"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(
                self.output_entry.get().strip(),
                f"数据比对结果_{timestamp}.xlsx"
            )
            
            self.log_message(f"正在生成报告: {output_file}")
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 确保至少有一个可见的工作表
                has_visible_sheet = False
                
                # 写入各类比对结果
                for category, data in results.items():
                    if not data.empty:
                        has_visible_sheet = True
                        # 移除合并指示列
                        data_clean = data.drop(columns=['_merge'])
                        
                        # 按照指定顺序重新组织列
                        # 1. 主键字段（每个主键显示文件1和文件2的列）
                        # 2. 对比字段（每个对比字段显示文件1、文件2和对比结果列）
                        # 3. 其他字段（文件1的其他字段，然后文件2的其他字段）
                        
                        # 收集所有列
                        key_cols = []  # 主键相关列
                        compare_cols = []  # 对比字段相关列
                        other_file1_cols = []  # 文件1其他字段
                        other_file2_cols = []  # 文件2其他字段
                        status_cols = []  # 比对状态列
                        
                        for col in data_clean.columns:
                            if col in keys:
                                # 主键字段，需要找到对应的文件1和文件2列
                                key_cols.append(col)
                                # 查找对应的文件1和文件2列
                                file1_col = None
                                file2_col = None
                                for c in data_clean.columns:
                                    if c == col + '_文件1' or (c.endswith('_文件1') and c.replace('_文件1', '') == col):
                                        file1_col = c
                                    elif c == col + '_文件2' or (c.endswith('_文件2') and c.replace('_文件2', '') == col):
                                        file2_col = c
                                if file1_col and file1_col not in key_cols:
                                    key_cols.append(file1_col)
                                if file2_col and file2_col not in key_cols:
                                    key_cols.append(file2_col)
                            elif col.endswith('_比对状态'):
                                # 比对状态列
                                status_cols.append(col)
                            elif any(col.startswith(field + '_文件1') or col == field + '_文件1' for field in mappings.keys()):
                                # 对比字段的文件1列
                                compare_cols.append(col)
                            elif any(col.startswith(field + '_文件2') or col == field + '_文件2' for field in mappings.keys()):
                                # 对比字段的文件2列
                                compare_cols.append(col)
                            elif col.endswith('_文件1'):
                                # 文件1的其他字段
                                other_file1_cols.append(col)
                            elif col.endswith('_文件2'):
                                # 文件2的其他字段
                                other_file2_cols.append(col)
                            else:
                                # 其他未分类的列
                                other_file1_cols.append(col)
                        
                        # 按照指定顺序排列列
                        ordered_columns = []
                        
                        # 1. 主键字段（每个主键的文件1和文件2列）
                        for key in keys:
                            # 查找主键的文件1和文件2列
                            key_file1_col = None
                            key_file2_col = None
                            for col in data_clean.columns:
                                if col == key + '_文件1' or (col.endswith('_文件1') and col.replace('_文件1', '') == key):
                                    key_file1_col = col
                                elif col == key + '_文件2' or (col.endswith('_文件2') and col.replace('_文件2', '') == key):
                                    key_file2_col = col
                            
                            if key_file1_col:
                                ordered_columns.append(key_file1_col)
                            if key_file2_col:
                                ordered_columns.append(key_file2_col)
                        
                        # 2. 对比字段（每个对比字段的文件1、文件2和对比结果列）
                        for field1, field2 in mappings.items():
                            # 现在所有字段都有标准格式的文件标识
                            field1_col = field1 + '_文件1'
                            field2_col = field2 + '_文件2'
                            status_col = field1 + '_比对状态'
                            
                            if field1_col in data_clean.columns:
                                ordered_columns.append(field1_col)
                            if field2_col in data_clean.columns:
                                ordered_columns.append(field2_col)
                            if status_col in data_clean.columns:
                                ordered_columns.append(status_col)
                        
                        # 3. 文件1的其他字段
                        for col in other_file1_cols:
                            if col not in ordered_columns:
                                ordered_columns.append(col)
                        
                        # 4. 文件2的其他字段
                        for col in other_file2_cols:
                            if col not in ordered_columns:
                                ordered_columns.append(col)
                        
                        # 5. 其他未分类的列
                        for col in data_clean.columns:
                            if col not in ordered_columns:
                                ordered_columns.append(col)
                        
                        # 创建显示列名映射
                        display_columns = {}
                        for col in ordered_columns:
                            if col.endswith('_文件1'):
                                base_name = col.replace('_文件1', '')
                                display_columns[col] = f"{base_name}_文件1"
                            elif col.endswith('_文件2'):
                                base_name = col.replace('_文件2', '')
                                display_columns[col] = f"{base_name}_文件2"
                            elif col.endswith('_比对状态'):
                                base_name = col.replace('_比对状态', '')
                                display_columns[col] = f"{base_name}对比结果"
                            elif col in keys:
                                # 主键字段
                                display_columns[col] = f"{col} (主键)"
                            else:
                                display_columns[col] = col
                        
                        # 按照指定顺序重新排列数据
                        data_display = data_clean[ordered_columns].rename(columns=display_columns)
                        data_display.to_excel(writer, sheet_name=category, index=False)
                        
                        # 设置文本格式，防止数字截断
                        worksheet = writer.sheets[category]
                        
                        # 首先设置整列为文本格式
                        for col_num, column in enumerate(data_clean.columns, 1):
                            # 获取列对象并设置格式
                            column_letter = worksheet.cell(row=1, column=col_num).column_letter
                            worksheet.column_dimensions[column_letter].number_format = '@'
                            
                            # 同时设置每个单元格的格式
                            for row_num in range(2, len(data_clean) + 2):  # Excel行号从1开始，标题行是第1行
                                cell = worksheet.cell(row=row_num, column=col_num)
                                cell.number_format = '@'  # @表示文本格式
                                
                                # 对于可能的长数字，强制转换为字符串
                                if pd.notna(data_clean.iloc[row_num-2, col_num-1]):
                                    value = data_clean.iloc[row_num-2, col_num-1]
                                    if isinstance(value, (int, float)) and len(str(int(value))) > 10:
                                        # 强制转换为字符串，确保长数字不被截断
                                        cell.value = str(int(value))
                                        cell.number_format = '@'
                        
                        self.log_message(f"  所有列已设置为文本格式，长数字已强制转换为字符串")
                        
                        self.log_message(f"  {category} 已写入，共 {len(data)} 行，已设置文本格式")
                
                # 字段比对结果表
                if field_comparison_results:
                    has_visible_sheet = True
                    field_df = pd.DataFrame(field_comparison_results)
                    field_df.to_excel(writer, sheet_name='字段比对结果', index=False)
                    
                    # 设置字段比对结果表的文本格式
                    worksheet = writer.sheets['字段比对结果']
                    for col_num in range(1, len(field_df.columns) + 1):
                        for row_num in range(2, len(field_df) + 2):
                            cell = worksheet.cell(row=row_num, column=col_num)
                            cell.number_format = '@'
                    
                    self.log_message(f"  字段比对结果已写入，共 {len(field_df)} 个字段")
                
                # 统计信息
                has_visible_sheet = True
                stats_data = {
                    '统计项目': [
                        '文件1总数', '文件2总数', '完全匹配', '不一致', 
                        '仅在文件1', '仅在文件2'
                    ],
                    '数量': [
                        len(self.df1), len(self.df2),
                        len(results['完全匹配']), len(results['不一致']),
                        len(results['仅在文件1']), len(results['仅在文件2'])
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='比对统计', index=False)
                
                # 设置统计信息表的文本格式
                worksheet = writer.sheets['比对统计']
                for col_num in range(1, len(stats_df.columns) + 1):
                    for row_num in range(2, len(stats_df) + 2):
                        cell = worksheet.cell(row=row_num, column=col_num)
                        cell.number_format = '@'
                
                # 配置信息
                has_visible_sheet = True
                config_data = {
                    '配置项': ['主键字段', '主键映射关系', '字段映射', '文件1', '文件2', '工作表1', '工作表2', '文件1路径', '文件2路径'],
                    '配置值': [
                        ','.join(keys),
                        '; '.join([f"{k}->{v}" for k, v in key_mappings.items()]),
                        '; '.join([f"{k}->{v}" for k, v in mappings.items()]),
                        os.path.basename(self.file1_path),
                        os.path.basename(self.file2_path),
                        str(self.sheet1_name),
                        str(self.sheet2_name),
                        self.file1_path,
                        self.file2_path
                    ]
                }
                config_df = pd.DataFrame(config_data)
                config_df.to_excel(writer, sheet_name='配置信息', index=False)
                
                # 设置配置信息表的文本格式
                worksheet = writer.sheets['配置信息']
                for col_num in range(1, len(config_df.columns) + 1):
                    for row_num in range(2, len(config_df) + 2):
                        cell = worksheet.cell(row=row_num, column=col_num)
                        cell.number_format = '@'
                
            # 检查是否有可见的工作表
            if not has_visible_sheet:
                # 如果没有可见的工作表，创建一个空的汇总表
                empty_df = pd.DataFrame({'说明': ['没有数据需要显示', '所有比对结果都为空']})
                empty_df.to_excel(writer, sheet_name='数据汇总', index=False)
                self.log_message("  创建了空的数据汇总表")
            
            self.log_message(f"报告生成完成: {output_file}")
            self.root.after(0, lambda: messagebox.showinfo("成功", f"比对报告已生成:\n{output_file}"))
            
        except Exception as e:
            self.log_message(f"生成报告失败: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"生成报告失败: {str(e)}"))
            
    def clear_all(self):
        """清空所有输入和状态"""
        self.file1_entry.delete(0, tk.END)
        self.file2_entry.delete(0, tk.END)
        self.key_entry.delete(0, tk.END)
        self.key_entry.insert(0, "税号,客商编码")
        self.mapping_entry.delete(0, tk.END)
        self.mapping_entry.insert(0, "公司名称:公司名称,联系人:联系人,电话:电话")
        self.status_text.delete(1.0, tk.END)
        self.df1 = None
        self.df2 = None
        self.file1_path = ""
        self.file2_path = ""
        self.sheet1_name = ""
        self.sheet2_name = ""
        self.compare_button.config(state="disabled")
        self.update_progress("准备就绪", 0, "等待开始...")

def main():
    root = tk.Tk()
    app = DataComparisonTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
