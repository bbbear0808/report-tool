import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import sys
import io
from docxtpl import DocxTemplate
from datetime import datetime
import threading
import tempfile

# ---------- 工具函数 ----------
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def find_column(df, keywords, allow_missing=False):
    for col in df.columns:
        col_lower = str(col).lower().replace(' ', '')
        for kw in keywords:
            if kw in col_lower:
                return col
    if allow_missing:
        return None
    raise ValueError(f"未找到包含 {keywords} 的列，可用列为：{list(df.columns)}")

def process_progress_sheet(df):
    progress_cols = {
        'priority': find_column(df, ['优先级', 'priority']),
        'test_func': find_column(df, ['测试功能', '功能', '模块', 'function']),
        'start_date': find_column(df, ['开测时间', '开始时间', 'start']),
        'end_date': find_column(df, ['完结时间', '结束时间', 'end']),
        'later_fix': find_column(df, ['后期修复', '延期修复']),
        'later_fix_level': find_column(df, ['后期修复bug等级', '后期修复等级']),
        'no_fix': find_column(df, ['不做修改', '不修改']),
        'no_fix_level': find_column(df, ['不做修改bug等级', '不修改等级']),
    }
    result_df = df[[v for v in progress_cols.values()]].copy()
    result_df.columns = list(progress_cols.keys())
    return result_df.to_dict(orient='records')

def process_bug_sheet(df, sheet_name):
    bug_level_col = find_column(df, ['bug等级', 'bug级别', '等级', 'level', '严重程度'])
    bug_desc_col = find_column(df, ['bug描述', '问题描述', '描述', 'description'])
    level_series = df[bug_level_col].astype(str).str.strip().str.upper()
    valid_mask = level_series != 'NA'
    df_clean = df[valid_mask].copy()
    total = len(df_clean)
    s_count = int((level_series[valid_mask] == 'S').sum())
    a_count = int((level_series[valid_mask] == 'A').sum())
    b_count = int((level_series[valid_mask] == 'B').sum())
    c_count = int((level_series[valid_mask] == 'C').sum())
    bugs = df_clean[[bug_level_col, bug_desc_col]].copy()
    bugs.columns = ['Bug等级', 'Bug描述']
    bugs_list = bugs.to_dict(orient='records')
    return {
        'sheet_name': sheet_name,
        'total_bugs': total,
        's_count': s_count,
        'a_count': a_count,
        'b_count': b_count,
        'c_count': c_count,
        'bugs': bugs_list
    }

def generate_report(context, template_path):
    doc = DocxTemplate(template_path)
    doc.render(context)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ---------- 主窗口 ----------
class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("测试报告自动生成工具")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        self.sheet_names = []
        self.progress_data = None
        self.bug_stats = []
        self.all_sheets = {}
        self.template_path = None

        self.base_dir = get_base_dir()
        self.default_template = os.path.join(self.base_dir, "template.docx")

        self.create_widgets()

    def create_widgets(self):
        # 主框架，带滚动条
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 步骤1：上传Excel
        frame1 = ttk.LabelFrame(main_frame, text="步骤1：上传 Excel 文件", padding=10)
        frame1.pack(fill=tk.X, pady=5)

        ttk.Label(frame1, text="Excel 文件：").grid(row=0, column=0, sticky=tk.W)
        self.excel_path = tk.StringVar()
        ttk.Entry(frame1, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame1, text="浏览...", command=self.browse_excel).grid(row=0, column=2)

        self.sheet_label = ttk.Label(frame1, text="已读取页签：无")
        self.sheet_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)

        ttk.Label(frame1, text="选择进度页签：").grid(row=2, column=0, sticky=tk.W)
        self.progress_sheet_combo = ttk.Combobox(frame1, state="readonly", width=40)
        self.progress_sheet_combo.grid(row=2, column=1, padx=5)
        self.progress_sheet_combo.bind("<<ComboboxSelected>>", lambda e: None)

        ttk.Button(frame1, text="提取数据", command=self.extract_data).grid(row=2, column=2)

        # 进度预览
        self.progress_text = tk.Text(frame1, height=6, state=tk.DISABLED, wrap=tk.WORD)
        self.progress_text.grid(row=3, column=0, columnspan=3, pady=5, sticky=tk.EW)

        # 步骤2：填写信息
        frame2 = ttk.LabelFrame(main_frame, text="步骤2：填写报告信息", padding=10)
        frame2.pack(fill=tk.X, pady=5)

        # 第一行
        ttk.Label(frame2, text="项目名称：").grid(row=0, column=0, sticky=tk.W)
        self.project_var = tk.StringVar()
        ttk.Entry(frame2, textvariable=self.project_var, width=25).grid(row=0, column=1, padx=5)

        ttk.Label(frame2, text="版本号：").grid(row=0, column=2, sticky=tk.W)
        self.version_var = tk.StringVar()
        ttk.Entry(frame2, textvariable=self.version_var, width=15).grid(row=0, column=3, padx=5)

        # 第二行
        ttk.Label(frame2, text="测试人员：").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.tester_var = tk.StringVar()
        ttk.Entry(frame2, textvariable=self.tester_var, width=25).grid(row=1, column=1, padx=5)

        ttk.Label(frame2, text="日期：").grid(row=1, column=2, sticky=tk.W)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(frame2, textvariable=self.date_var, width=15).grid(row=1, column=3, padx=5)

        # 测试环境
        ttk.Label(frame2, text="测试环境：").grid(row=2, column=0, sticky=tk.W)
        env_frame = ttk.Frame(frame2)
        env_frame.grid(row=2, column=1, columnspan=3, sticky=tk.W)
        self.env_ios = tk.BooleanVar()
        self.env_android = tk.BooleanVar()
        self.env_web = tk.BooleanVar()
        self.env_harmony = tk.BooleanVar()
        ttk.Checkbutton(env_frame, text="iOS", variable=self.env_ios).pack(side=tk.LEFT)
        ttk.Checkbutton(env_frame, text="Android", variable=self.env_android).pack(side=tk.LEFT)
        ttk.Checkbutton(env_frame, text="Web", variable=self.env_web).pack(side=tk.LEFT)
        ttk.Checkbutton(env_frame, text="鸿蒙", variable=self.env_harmony).pack(side=tk.LEFT)

        # 测试类型
        ttk.Label(frame2, text="测试类型：").grid(row=3, column=0, sticky=tk.W)
        self.test_type_var = tk.StringVar(value="功能测试")
        type_frame = ttk.Frame(frame2)
        type_frame.grid(row=3, column=1, columnspan=3, sticky=tk.W)
        for t in ["功能测试", "集成测试", "回归测试", "验收测试", "其他"]:
            ttk.Radiobutton(type_frame, text=t, value=t, variable=self.test_type_var).pack(side=tk.LEFT, padx=5)

        # 测试结论
        ttk.Label(frame2, text="测试结论：").grid(row=4, column=0, sticky=tk.NW)
        self.conclusion_text = tk.Text(frame2, height=5, width=60)
        self.conclusion_text.grid(row=4, column=1, columnspan=3, pady=5, sticky=tk.W)

        # 备注
        ttk.Label(frame2, text="备注：").grid(row=5, column=0, sticky=tk.NW)
        self.remark_text = tk.Text(frame2, height=3, width=60)
        self.remark_text.grid(row=5, column=1, columnspan=3, pady=5, sticky=tk.W)

        # 步骤3：生成报告
        frame3 = ttk.LabelFrame(main_frame, text="步骤3：生成报告", padding=10)
        frame3.pack(fill=tk.X, pady=5)

        ttk.Label(frame3, text="模板文件：").grid(row=0, column=0, sticky=tk.W)
        self.template_path_var = tk.StringVar()
        if os.path.exists(self.default_template):
            self.template_path_var.set(self.default_template)
        ttk.Entry(frame3, textvariable=self.template_path_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame3, text="浏览...", command=self.browse_template).grid(row=0, column=2)

        ttk.Button(frame3, text="生成并下载报告", command=self.generate_report_thread).grid(row=1, column=1, pady=10)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10,0))

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_path.set(path)
            try:
                xls = pd.ExcelFile(path)
                self.sheet_names = xls.sheet_names
                self.sheet_label.config(text=f"已读取页签：{', '.join(self.sheet_names)}")
                self.progress_sheet_combo['values'] = self.sheet_names
                if self.sheet_names:
                    self.progress_sheet_combo.current(0)
                # 读取所有页签备用
                self.all_sheets = {name: pd.read_excel(xls, sheet_name=name) for name in self.sheet_names}
                self.status_var.set(f"已加载文件：{os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("错误", f"读取Excel失败：{e}")

    def extract_data(self):
        if not self.sheet_names:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        selected_sheet = self.progress_sheet_combo.get()
        if not selected_sheet:
            messagebox.showwarning("警告", "请选择进度页签")
            return
        try:
            df_progress = self.all_sheets[selected_sheet]
            self.progress_data = process_progress_sheet(df_progress)
            # 处理Bug页签
            self.bug_stats = []
            for name in self.sheet_names:
                if name == selected_sheet:
                    continue
                try:
                    stats = process_bug_sheet(self.all_sheets[name], name)
                    self.bug_stats.append(stats)
                except Exception as e:
                    self.status_var.set(f"跳过页签 {name}：{e}")
            # 更新预览
            self.progress_text.config(state=tk.NORMAL)
            self.progress_text.delete(1.0, tk.END)
            self.progress_text.insert(tk.END, f"进度数据提取完成，共 {len(self.progress_data)} 条记录\n")
            self.progress_text.insert(tk.END, "Bug统计：\n")
            for stat in self.bug_stats:
                self.progress_text.insert(tk.END, f"  {stat['sheet_name']}: 总数{stat['total_bugs']} (S:{stat['s_count']} A:{stat['a_count']} B:{stat['b_count']} C:{stat['c_count']})\n")
            self.progress_text.config(state=tk.DISABLED)
            self.status_var.set("数据提取成功")
        except Exception as e:
            messagebox.showerror("错误", f"提取数据失败：{e}")

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.template_path_var.set(path)

    def generate_report_thread(self):
        # 在新线程中生成，避免界面卡顿
        threading.Thread(target=self.generate_report, daemon=True).start()

    def generate_report(self):
        if self.progress_data is None:
            messagebox.showwarning("警告", "请先提取数据")
            return
        project = self.project_var.get()
        if not project:
            messagebox.showwarning("警告", "请填写项目名称")
            return

        # 收集环境
        envs = []
        if self.env_ios.get(): envs.append("iOS")
        if self.env_android.get(): envs.append("Android")
        if self.env_web.get(): envs.append("Web")
        if self.env_harmony.get(): envs.append("鸿蒙")
        env_str = "、".join(envs) if envs else "未指定"

        context = {
            'project': project,
            'version': self.version_var.get(),
            'tester': self.tester_var.get(),
            'date': self.date_var.get(),
            'test_type': self.test_type_var.get(),
            'test_env': env_str,
            'conclusion': self.conclusion_text.get("1.0", tk.END).strip(),
            'remark': self.remark_text.get("1.0", tk.END).strip(),
            'progress_table': self.progress_data or [],
            'bug_stats': self.bug_stats or [],
        }

        template_path = self.template_path_var.get()
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("错误", "请选择有效的模板文件")
            return

        try:
            self.status_var.set("正在生成报告...")
            output = generate_report(context, template_path)
            # 弹出保存对话框
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx")],
                initialfile=f"测试报告_{project}_{self.date_var.get()}.docx"
            )
            if save_path:
                with open(save_path, "wb") as f:
                    f.write(output.getbuffer())
                self.status_var.set(f"报告已保存：{os.path.basename(save_path)}")
                messagebox.showinfo("成功", f"报告已生成并保存到：\n{save_path}")
            else:
                self.status_var.set("取消保存")
        except Exception as e:
            messagebox.showerror("错误", f"生成报告失败：{e}")
            self.status_var.set("生成失败")

# ---------- 启动 ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
