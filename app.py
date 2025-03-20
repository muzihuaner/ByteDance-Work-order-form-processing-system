import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
from datetime import datetime

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("工单表格处理系统 v1.0")
        self.root.geometry("600x400")
        
        # 作者信息
        self.author_label = ttk.Label(
            root,
            text="作者：李欢 https://github.com/muzihuaner | 使用说明：1.选择文件 2.开始处理 3.自动用WPS打开结果",
            foreground="gray"
        )
        self.author_label.pack(side=tk.BOTTOM, fill=tk.X)

        # 主界面部件
        self.frame = ttk.Frame(root)
        self.frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

        self.file_label = ttk.Label(self.frame, text="选择原始Excel文件:")
        self.file_label.grid(row=0, column=0, sticky=tk.W)

        self.file_entry = ttk.Entry(self.frame, width=40)
        self.file_entry.grid(row=1, column=0, padx=5)

        self.browse_btn = ttk.Button(
            self.frame, 
            text="浏览...",
            command=self.browse_file
        )
        self.browse_btn.grid(row=1, column=1, padx=5)

        self.process_btn = ttk.Button(
            self.frame,
            text="开始处理",
            command=self.process_file,
            state=tk.DISABLED
        )
        self.process_btn.grid(row=2, column=0, pady=20, sticky=tk.W)

        self.status_label = ttk.Label(self.frame, text="")
        self.status_label.grid(row=3, column=0, columnspan=2)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.process_btn["state"] = tk.NORMAL

    def process_file(self):
        input_file = self.file_entry.get()
        if not input_file:
            messagebox.showerror("错误", "请先选择文件！")
            return

        try:
            # 生成带时间戳的输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"处理结果_{timestamp}.xlsx"
            
            # 执行数据处理
            self.status_label["text"] = "正在处理中，请稍候..."
            self.root.update()  # 更新界面显示
            
            df = pd.read_excel(input_file)
            
            # 原有数据处理逻辑
            df['房间'] = df['房间'].apply(self.extract_room)
            df['维修开始时间'] = df['维修开始时间'].apply(self.parse_datetime)
            df['维修日期'] = df['维修开始时间'].dt.date
            df['维修时间'] = df['维修开始时间'].dt.time

            new_df = pd.DataFrame({
                '工单ID': df['ID'],
                '所属物理机': df['资产编号'],
                '库房位置': df['房间'],
                '服务器厂商': df['厂商'],
                '服务器SN': df['SN'],
                '故障盘品牌': df['原件品牌'],
                '故障盘SN': df['原件SN'],
                '故障盘PN': df['原件PN'],
                '更换盘品牌': df['新件品牌'],
                '更换盘SN': df['新件SN'],
                '更换盘PN': df['新件PN'],
                '维修日期': df['维修日期'],
                '维修时间': df['维修时间'],
                '不返还字段':'不返还',
                '故障盘配件容量':'',
                '故障盘型号':'',
                'IP': df['IPv6'],
                '机柜位置': df['机架位']
            })

            new_df.to_excel(output_file, index=False, engine='openpyxl')
            
            # 尝试用WPS打开
            try:
                if os.name == 'nt':  # Windows系统
                    os.startfile(output_file)
                else:  # 其他系统
                    subprocess.call(('xdg-open', output_file))
                self.status_label["text"] = f"处理完成！已自动用WPS打开 {output_file}"
            except Exception as e:
                self.status_label["text"] = f"处理完成！但自动打开失败：{str(e)}"
                
        except Exception as e:
            messagebox.showerror("处理错误", f"处理过程中发生错误：\n{str(e)}")
            self.status_label["text"] = "处理失败"

    @staticmethod
    def extract_room(room_str):
        try:
            parts = room_str.split('_')
            return '_'.join(parts[:2])
        except:
            return None

    @staticmethod
    def parse_datetime(dt_str):
        if pd.isna(dt_str) or dt_str == "<nil>":
            return pd.NaT
        try:
            return pd.to_datetime(dt_str)
        except:
            return pd.NaT

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()