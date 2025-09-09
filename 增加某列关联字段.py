import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import threading

class FieldMatchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("字段匹配关联工具 v1.3")
        self.root.geometry("800x750")
        
        self.setup_ui()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 创建滚动区域
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 文件1配置
        file1_frame = ttk.LabelFrame(scrollable_frame, text="文件1配置", padding="10")
        file1_frame.pack(fill="x", pady=5)
        
        ttk.Label(file1_frame, text="文件1路径:").grid(row=0, column=0, sticky="w", pady=2)
        self.file1_entry = ttk.Entry(file1_frame, width=60)
        self.file1_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file1_frame, text="浏览", command=self.browse_file1).grid(row=0, column=2, pady=2)
        
        ttk.Label(file1_frame, text="工作表名称:").grid(row=1, column=0, sticky="w", pady=2)
        self.sheet1_entry = ttk.Entry(file1_frame, width=20)
        self.sheet1_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        self.sheet1_entry.insert(0, "默认")
        
        # 文件2配置
        file2_frame = ttk.LabelFrame(scrollable_frame, text="文件2配置", padding="10")
        file2_frame.pack(fill="x", pady=5)
        
        ttk.Label(file2_frame, text="文件2路径:").grid(row=0, column=0, sticky="w", pady=2)
        self.file2_entry = ttk.Entry(file2_frame, width=60)
        self.file2_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file2_frame, text="浏览", command=self.browse_file2).grid(row=0, column=2, pady=2)
        
        ttk.Label(file2_frame, text="工作表名称:").grid(row=1, column=0, sticky="w", pady=2)
        self.sheet2_entry = ttk.Entry(file2_frame, width=20)
        self.sheet2_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 匹配字段配置
        match_frame = ttk.LabelFrame(scrollable_frame, text="匹配字段配置", padding="10")
        match_frame.pack(fill="x", pady=5)
        
        # 匹配字段表格
        ttk.Label(match_frame, text="匹配字段对（文件1字段:文件2字段）:").pack(anchor="w")
        
        # 创建匹配字段输入区域
        match_input_frame = ttk.Frame(match_frame)
        match_input_frame.pack(fill="x", pady=5)
        
        self.match_entries = []
        for i in range(3):  # 默认3个匹配字段对
            entry_frame = ttk.Frame(match_input_frame)
            entry_frame.pack(fill="x", pady=2)
            
            ttk.Label(entry_frame, text=f"匹配字段{i+1}:").pack(side="left")
            entry1 = ttk.Entry(entry_frame, width=20)
            entry1.pack(side="left", padx=5)
            ttk.Label(entry_frame, text=":").pack(side="left")
            entry2 = ttk.Entry(entry_frame, width=20)
            entry2.pack(side="left", padx=5)
            
            self.match_entries.append((entry1, entry2))
        
        # 添加/删除匹配字段按钮
        match_btn_frame = ttk.Frame(match_frame)
        match_btn_frame.pack(fill="x", pady=5)
        ttk.Button(match_btn_frame, text="添加匹配字段", command=self.add_match_field).pack(side="left", padx=5)
        ttk.Button(match_btn_frame, text="删除匹配字段", command=self.remove_match_field).pack(side="left", padx=5)
        
        # 写入字段配置
        write_frame = ttk.LabelFrame(scrollable_frame, text="写入字段配置", padding="10")
        write_frame.pack(fill="x", pady=5)
        
        ttk.Label(write_frame, text="写入字段（文件2中的字段名，多个字段用逗号分隔）:").pack(anchor="w")
        
        # 创建写入字段输入区域
        write_input_frame = ttk.Frame(write_frame)
        write_input_frame.pack(fill="x", pady=5)
        
        self.write_fields_entry = ttk.Entry(write_input_frame, width=80)
        self.write_fields_entry.pack(fill="x", pady=2)
        self.write_fields_entry.insert(0, "对应单据编号,实收金额,应收金额")
        
        ttk.Label(write_frame, text="说明: 输入文件2中需要写入到文件1的字段名，用逗号分隔", 
                 font=("Arial", 9), foreground="gray").pack(anchor="w", pady=2)
        ttk.Label(write_frame, text="示例: 对应单据编号,实收金额,应收金额", 
                 font=("Arial", 9), foreground="gray").pack(anchor="w", pady=2)
        
        # 插入位置配置
        insert_frame = ttk.LabelFrame(scrollable_frame, text="插入位置配置", padding="10")
        insert_frame.pack(fill="x", pady=5)
        
        ttk.Label(insert_frame, text="插入位置:").grid(row=0, column=0, sticky="w", pady=2)
        self.insert_position_entry = ttk.Entry(insert_frame, width=30)
        self.insert_position_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        self.insert_position_entry.insert(0, "最后一列")
        
        ttk.Label(insert_frame, text="说明: 指定文件1中插入新列的起始位置，可以是列名或列索引", 
                 font=("Arial", 9), foreground="gray").grid(row=1, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(insert_frame, text="示例: '最后一列' 或 'title' 或 '3'", 
                 font=("Arial", 9), foreground="gray").grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        
        # 输出配置
        output_frame = ttk.LabelFrame(scrollable_frame, text="输出配置", padding="10")
        output_frame.pack(fill="x", pady=5)
        
        ttk.Label(output_frame, text="输出文件路径:").grid(row=0, column=0, sticky="w", pady=2)
        self.output_entry = ttk.Entry(output_frame, width=60)
        self.output_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(output_frame, text="浏览", command=self.browse_output).grid(row=0, column=2, pady=2)
        
        ttk.Label(output_frame, text="输出文件名:").grid(row=1, column=0, sticky="w", pady=2)
        self.filename_entry = ttk.Entry(output_frame, width=40)
        self.filename_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        self.filename_entry.insert(0, "匹配结果")
        
        # 字段配置说明
        field_frame = ttk.LabelFrame(scrollable_frame, text="配置说明", padding="10")
        field_frame.pack(fill="x", pady=5)
        
        ttk.Label(field_frame, text="• 匹配字段：用于确定两个文件中哪些记录是匹配的", font=("Arial", 9)).pack(anchor="w")
        ttk.Label(field_frame, text="• 写入字段：匹配成功后，将文件2的指定字段值写入文件1的新列", font=("Arial", 9)).pack(anchor="w")
        ttk.Label(field_frame, text="• 插入位置：指定新列在文件1中的插入位置，默认为最后一列", font=("Arial", 9)).pack(anchor="w")
        ttk.Label(field_frame, text="• 新列名自动生成：原字段名_文件2", font=("Arial", 9)).pack(anchor="w")
        ttk.Label(field_frame, text="• 支持多字段匹配和多字段写入，提高匹配准确性", font=("Arial", 9)).pack(anchor="w")
        ttk.Label(field_frame, text="• 示例：title:NC账单编号, 对应单据编号,实收金额", font=("Arial", 9)).pack(anchor="w")
        
        # 操作按钮
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill="x", pady=10)
        
        self.start_button = ttk.Button(button_frame, text="开始匹配", command=self.start_matching)
        self.start_button.pack(side="left", padx=5)
        ttk.Button(button_frame, text="清空配置", command=self.clear_config).pack(side="left", padx=5)
        ttk.Button(button_frame, text="加载示例", command=self.load_example).pack(side="left", padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side="right", padx=5)
        
        # 状态显示
        status_frame = ttk.LabelFrame(scrollable_frame, text="状态信息", padding="10")
        status_frame.pack(fill="both", expand=True, pady=5)
        
        self.status_text = tk.Text(status_frame, height=8, wrap="word")
        status_scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=status_scrollbar.set)
        
        self.status_text.pack(side="left", fill="both", expand=True)
        status_scrollbar.pack(side="right", fill="y")
        
        # 打包滚动区域
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def add_match_field(self):
        """添加匹配字段"""
        entry_frame = ttk.Frame(self.match_entries[0][0].master)
        entry_frame.pack(fill="x", pady=2)
        
        ttk.Label(entry_frame, text=f"匹配字段{len(self.match_entries)+1}:").pack(side="left")
        entry1 = ttk.Entry(entry_frame, width=20)
        entry1.pack(side="left", padx=5)
        ttk.Label(entry_frame, text=":").pack(side="left")
        entry2 = ttk.Entry(entry_frame, width=20)
        entry2.pack(side="left", padx=5)
        
        self.match_entries.append((entry1, entry2))
        
    def remove_match_field(self):
        """删除匹配字段"""
        if len(self.match_entries) > 1:
            entry1, entry2 = self.match_entries.pop()
            entry1.master.destroy()
            
    def load_example(self):
        """加载示例配置"""
        # 清空现有配置
        self.clear_config()
        
        # 设置示例匹配字段
        self.match_entries[0][0].insert(0, "title")
        self.match_entries[0][1].insert(0, "NC账单编号")
        
        # 设置示例写入字段
        self.write_fields_entry.delete(0, tk.END)
        self.write_fields_entry.insert(0, "对应单据编号,实收金额,应收金额")
        
        self.log_message("已加载示例配置")
        
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
        filename = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)
            
    def log_message(self, message):
        self.status_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.status_text.see(tk.END)
        self.root.update()
        
    def clear_config(self):
        # 清空文件路径
        self.file1_entry.delete(0, tk.END)
        self.file2_entry.delete(0, tk.END)
        self.sheet1_entry.delete(0, tk.END)
        self.sheet1_entry.insert(0, "默认")
        self.sheet2_entry.delete(0, tk.END)
        self.output_entry.delete(0, tk.END)
        self.filename_entry.delete(0, tk.END)
        self.filename_entry.insert(0, "匹配结果")
        self.insert_position_entry.delete(0, tk.END)
        self.insert_position_entry.insert(0, "最后一列")
        self.write_fields_entry.delete(0, tk.END)
        self.write_fields_entry.insert(0, "对应单据编号,实收金额,应收金额")
        
        # 清空匹配字段
        for entry1, entry2 in self.match_entries:
            entry1.delete(0, tk.END)
            entry2.delete(0, tk.END)
            
        self.status_text.delete(1.0, tk.END)
        
    def start_matching(self):
        try:
            # 获取配置
            file1_path = self.file1_entry.get().strip()
            file2_path = self.file2_entry.get().strip()
            sheet1_name = self.sheet1_entry.get().strip()
            sheet2_name = self.sheet2_entry.get().strip()
            output_path = self.output_entry.get().strip()
            filename = self.filename_entry.get().strip()
            insert_position = self.insert_position_entry.get().strip()
            write_fields_str = self.write_fields_entry.get().strip()
            
            # 验证配置
            if not file1_path:
                messagebox.showerror("错误", "请选择文件1")
                return
            if not file2_path:
                messagebox.showerror("错误", "请选择文件2")
                return
            if not sheet2_name:
                messagebox.showerror("错误", "请输入文件2的工作表名称")
                return
            if not output_path and not filename:
                messagebox.showerror("错误", "请设置输出文件路径或文件名")
                return
            if not write_fields_str:
                messagebox.showerror("错误", "请输入写入字段")
                return
                
            # 检查文件是否存在
            if not os.path.exists(file1_path):
                messagebox.showerror("错误", f"文件1不存在: {file1_path}")
                return
            if not os.path.exists(file2_path):
                messagebox.showerror("错误", f"文件2不存在: {file2_path}")
                return
            
            # 获取匹配字段配置
            match_fields = []
            for entry1, entry2 in self.match_entries:
                field1 = entry1.get().strip()
                field2 = entry2.get().strip()
                if field1 and field2:
                    match_fields.append((field1, field2))
            
            if not match_fields:
                messagebox.showerror("错误", "请至少配置一个匹配字段")
                return
                
            # 解析写入字段
            write_fields = [field.strip() for field in write_fields_str.split(',') if field.strip()]
            if not write_fields:
                messagebox.showerror("错误", "请输入有效的写入字段")
                return
            
            # 生成输出文件路径
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_dir = os.path.dirname(os.path.abspath(file1_path))
                output_path = os.path.join(output_dir, f"{filename}_{timestamp}.xlsx")
            elif not output_path.endswith('.xlsx'):
                output_path += '.xlsx'
            
            self.log_message("开始执行字段匹配...")
            self.log_message(f"文件1: {file1_path}")
            self.log_message(f"文件2: {file2_path}")
            self.log_message(f"匹配字段: {match_fields}")
            self.log_message(f"写入字段: {write_fields}")
            self.log_message(f"插入位置: {insert_position}")
            self.log_message(f"输出路径: {output_path}")
            
            # 禁用开始按钮
            self.start_button.config(state="disabled")
            
            # 在新线程中执行匹配
            thread = threading.Thread(target=self.execute_matching, 
                                    args=(file1_path, file2_path, sheet1_name, sheet2_name, 
                                         match_fields, write_fields, insert_position, output_path))
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            self.log_message(f"配置验证失败: {str(e)}")
            messagebox.showerror("错误", f"配置验证失败: {str(e)}")
    
    def execute_matching(self, file1_path, file2_path, sheet1_name, sheet2_name, 
                        match_fields, write_fields, insert_position, output_path):
        """在线程中执行匹配"""
        try:
            result_df = self.match_and_merge_files(file1_path, file2_path, sheet1_name, 
                                                 sheet2_name, match_fields, write_fields, 
                                                 insert_position, output_path)
            
            # 显示结果
            matched_count = 0
            for field in write_fields:
                new_col_name = f"{field}_文件2"
                if new_col_name in result_df.columns:
                    matched_count = (result_df[new_col_name] != '').sum()
                    break
            
            if matched_count > 0:
                self.log_message(f"匹配成功！共匹配 {matched_count} 条记录")
                # 显示前3条匹配记录
                for i, (_, row) in enumerate(result_df.head(3).iterrows()):
                    match_info = []
                    for field in write_fields:
                        new_col_name = f"{field}_文件2"
                        if new_col_name in result_df.columns and row[new_col_name]:
                            match_info.append(f"{new_col_name}={row[new_col_name]}")
                    if match_info:
                        self.log_message(f"  记录{i+1}: {', '.join(match_info)}")
            else:
                self.log_message("没有找到匹配的记录")
                
            self.root.after(0, lambda: messagebox.showinfo("完成", f"处理完成！\n结果已保存到: {output_path}"))
            
        except Exception as e:
            self.log_message(f"处理失败: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"处理失败: {str(e)}"))
        finally:
            # 恢复开始按钮
            self.root.after(0, lambda: self.start_button.config(state="normal"))
    
    def match_and_merge_files(self, file1_path, file2_path, sheet1_name, sheet2_name, 
                            match_fields, write_fields, insert_position, output_path):
        """执行字段匹配和合并"""
        # 读取文件1和文件2
        self.log_message("正在读取文件1...")
        if sheet1_name == "默认":
            df1 = pd.read_excel(file1_path, dtype=str)
        else:
            df1 = pd.read_excel(file1_path, sheet_name=sheet1_name, dtype=str)
        self.log_message(f"文件1加载成功，共 {len(df1)} 行，{len(df1.columns)} 列")
        
        self.log_message("正在读取文件2...")
        df2 = pd.read_excel(file2_path, sheet_name=sheet2_name, dtype=str)
        self.log_message(f"文件2加载成功，共 {len(df2)} 行，{len(df2.columns)} 列")
        
        # 检查必要的列是否存在
        for field1, field2 in match_fields:
            if field1 not in df1.columns:
                raise ValueError(f"文件1中缺少字段: {field1}")
            if field2 not in df2.columns:
                raise ValueError(f"文件2中缺少字段: {field2}")
        
        for field in write_fields:
            if field not in df2.columns:
                raise ValueError(f"文件2中缺少写入字段: {field}")
        
        # 创建文件1的副本
        result_df = df1.copy()
        
        # 确定插入位置
        insert_index = self.get_insert_index(result_df, insert_position)
        self.log_message(f"新列将插入到位置: {insert_index}")
        
        # 创建文件2的查找字典
        lookup_dict = {}
        for idx, row in df2.iterrows():
            # 构建匹配键
            match_key = []
            for field1, field2 in match_fields:
                value = str(row[field2]).strip()
                if value and value != 'nan':
                    match_key.append(value)
                else:
                    match_key = None
                    break
            
            if match_key:
                # 构建写入值
                write_values = {}
                for field in write_fields:
                    value = str(row[field]).strip()
                    if value and value != 'nan':
                        write_values[f"{field}_文件2"] = value
                    else:
                        write_values[f"{field}_文件2"] = ''
                
                lookup_dict[tuple(match_key)] = write_values
        
        self.log_message(f"文件2中有效的匹配记录数量: {len(lookup_dict)}")
        
        # 匹配并填充数据
        matched_count = 0
        for idx, row in result_df.iterrows():
            # 构建匹配键
            match_key = []
            for field1, field2 in match_fields:
                value = str(row[field1]).strip()
                if value and value != 'nan':
                    match_key.append(value)
                else:
                    match_key = None
                    break
            
            if match_key and tuple(match_key) in lookup_dict:
                # 填充写入字段
                write_values = lookup_dict[tuple(match_key)]
                for col_name, value in write_values.items():
                    result_df.at[idx, col_name] = value
                matched_count += 1
        
        self.log_message(f"成功匹配 {matched_count} 条记录")
        
        # 重新排列列，将新列插入到指定位置
        result_df = self.reorder_columns(result_df, write_fields, insert_index)
        
        # 保存结果
        self.log_message("正在保存结果...")
        result_df.to_excel(output_path, index=False)
        
        # 设置Excel格式
        try:
            from openpyxl import load_workbook
            wb = load_workbook(output_path)
            ws = wb.active
            for col in ws.columns:
                for cell in col:
                    cell.number_format = '@'
            wb.save(output_path)
            self.log_message("已设置Excel文本格式")
        except Exception as e:
            self.log_message(f"警告: 设置Excel格式失败: {e}")
        
        return result_df
    
    def get_insert_index(self, df, insert_position):
        """获取插入位置索引"""
        if insert_position == "最后一列" or not insert_position:
            return len(df.columns)
        
        # 尝试按列名查找
        if insert_position in df.columns:
            return df.columns.get_loc(insert_position)
        
        # 尝试按索引查找
        try:
            index = int(insert_position)
            if 0 <= index <= len(df.columns):
                return index
        except ValueError:
            pass
        
        # 如果都找不到，默认插入到最后一列
        self.log_message(f"警告: 无法找到插入位置 '{insert_position}'，将插入到最后一列")
        return len(df.columns)
    
    def reorder_columns(self, df, write_fields, insert_index):
        """重新排列列，将新列插入到指定位置"""
        # 获取新列名
        new_columns = [f"{field}_文件2" for field in write_fields]
        
        # 获取原有列名
        original_columns = [col for col in df.columns if col not in new_columns]
        
        # 在指定位置插入新列
        if insert_index >= len(original_columns):
            # 插入到末尾
            reordered_columns = original_columns + new_columns
        else:
            # 插入到中间
            reordered_columns = (original_columns[:insert_index] + 
                               new_columns + 
                               original_columns[insert_index:])
        
        # 重新排列DataFrame
        result_df = df[reordered_columns]
        
        self.log_message(f"列重新排列完成，新列位置: {insert_index}")
        return result_df

def main():
    root = tk.Tk()
    app = FieldMatchTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
