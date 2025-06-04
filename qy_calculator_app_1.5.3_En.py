
# qy_calculator_app 1.5.3_En
# Created by Vince on 2025/03/28
# Last modified by Vince on 2025/05/04
# Fixed bug: Cannot copy and rename calculation history data

#  * Where mountain clouds arise, wisdom flows through twin rivers
#  * Blue waters become algorithms, green peaks as abacus rods  
#  * All things follow mathematical principles, myriad machines enact changes
#  * Green codes traverse the world, intelligent light illuminates the land
#  * (Dedicated to CIGIT, CAS - Code contains green mountains, algorithms hold blue springs)
#  *



import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinterdnd2 import TkinterDnD, DND_ALL
import pandas as pd
from datetime import datetime
from extract_data import extract_data
from calculate_integration import calculate_integration
import base64
from PIL import ImageTk, Image
import io




class QYCalculatorApp(TkinterDnD.Tk):
    
    def __init__(self):
        super().__init__()
        self.title("Quantum Yield Calculator v1.5.3_En")
        self.geometry("1200x850")
        # 添加全局字体配置
        self.style = ttk.Style()
        self.style.configure('.', font=('微软雅黑', 10))  # 所有组件默认字体
        self.style.configure('TLabel', font=('微软雅黑', 10))  # 标签专用
        self.style.configure('TEntry', font=('微软雅黑', 10))  # 输入框
        self.style.configure('TButton', font=('微软雅黑', 10))  # 按钮
        self.style.configure("Accent.TButton", foreground="blue", background="#3498db", font=('微软雅黑', 12, 'bold'))

        self.sample_counter = 1
        self.file_data = {"std": None, "sample": None}
        self.initialize_ui()

    def initialize_ui(self):
        # 修改点1：统一主容器名称
        self.main_frame = ttk.Frame(self)  # 保持原名称不变
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        # 上部容器（左侧+右侧面板）
        upper_frame = ttk.Frame(self.main_frame)  # 使用原容器名称
        upper_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧面板（保持原有参数）
        self.setup_left_panel()  # 恢复原始调用方式
        
        # 右侧面板（保持原有参数）
        self.setup_right_panel()  # 恢复原始调用方式
        
        # 历史记录面板
        self.setup_history_panel()

        # 上下文菜单配置
        self.setup_context_menu()

    def setup_left_panel(self):
        """左侧面板初始化"""
        self.left_panel = ttk.LabelFrame(self.main_frame, text="Data Source Configuration", width=450)
        self.left_panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=5, pady=5)
        # 文件拖拽区
        self.setup_file_areas()
        
        # 积分参数设置
        self.setup_integration_params()

    def setup_right_panel(self):
        """右侧面板初始化"""
        self.right_panel = ttk.LabelFrame(self.main_frame, text="QY Calculator", width=750)
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 输入字段
        self.setup_qy_inputs()
        
        # 操作按钮
        self.setup_action_buttons()
        
        # ========== 新增信息展示区 ==========
        info_frame = ttk.LabelFrame(self.right_panel, text="References")
        info_frame.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)

        # 创建Notebook标签页
        self.info_notebook = ttk.Notebook(info_frame)
        self.info_notebook.pack(fill=tk.BOTH, expand=True)

        # 公式标签页
        formula_tab = ttk.Frame(self.info_notebook)
        self._setup_formula_tab(formula_tab)
        self.info_notebook.add(formula_tab, text="Calc. Formula")

        # 溶剂折射率标签页
        solvent_tab = ttk.Frame(self.info_notebook)
        self._setup_solvent_tab(solvent_tab)
        self.info_notebook.add(solvent_tab, text="Solvent RI")

        # 开发者信息标签页
        about_tab = ttk.Frame(self.info_notebook)
        self._setup_about_tab(about_tab)
        self.info_notebook.add(about_tab, text="About")
        

    def setup_file_areas(self):
        """文件拖拽区配置"""
        # ========== 标液文件区域 ==========
        self.std_frame = ttk.LabelFrame(self.left_panel, text="QY standards")
        self.std_frame.pack(fill=tk.X, pady=5)
        
        # 文件操作组件
        std_control_frame = ttk.Frame(self.std_frame)
        std_control_frame.pack(fill=tk.X)
        
        ttk.Button(std_control_frame, text="BrowseFiles", width=10,
                  command=lambda: self.select_file("std")).pack(side=tk.LEFT, padx=2)
        ttk.Button(std_control_frame, text="Calc.Integral", width=10,
                  command=lambda: self.calculate_integral("std")).pack(side=tk.LEFT, padx=2)
        
        # 拖拽区域
        self.std_drop_area = ttk.Label(
            self.std_frame,
            text="Drag QY standards file here",
            relief="groove",
            anchor="center"
        )
        self.std_drop_area.pack(pady=5, fill=tk.X)
        self.std_drop_area.drop_target_register(DND_ALL)
        self.std_drop_area.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, "std"))
        
        # 文件状态显示
        self.std_file_label = ttk.Label(self.std_frame, text="⚠️ No file selected")
        self.std_file_label.pack(pady=2)

        # 新增标液积分结果显示（第1处修改）
        self.std_integral_label = ttk.Label(
            self.std_frame, 
            text="Integral: Not Calculated",
            foreground="#2c3e50",
            font=('Arial', 12)
        )
        self.std_integral_label.pack(pady=2)
        # ========== 样品文件区域 ==========
        self.sample_frame = ttk.LabelFrame(self.left_panel, text="Sample File")
        self.sample_frame.pack(fill=tk.X, pady=5)

        # 文件操作组件
        sample_control_frame = ttk.Frame(self.sample_frame)
        sample_control_frame.pack(fill=tk.X)
        
        ttk.Button(sample_control_frame, text="BrowseFiles", width=10,
                  command=lambda: self.select_file("sample")).pack(side=tk.LEFT, padx=2)
        ttk.Button(sample_control_frame, text="Calc.Integral", width=10,
                  command=lambda: self.calculate_integral("sample")).pack(side=tk.LEFT, padx=2)

        # 拖拽区域
        self.sample_drop_area = ttk.Label(
            self.sample_frame,
            text="Drag sample file here",
            relief="groove",
            anchor="center"
        )
        self.sample_drop_area.pack(pady=5, fill=tk.X)
        self.sample_drop_area.drop_target_register(DND_ALL)
        self.sample_drop_area.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, "sample"))
        
        # 文件状态显示
        self.sample_file_label = ttk.Label(self.sample_frame, text="⚠️ No file selected")
        self.sample_file_label.pack(pady=2)
    
        # 新增样品积分结果显示（第2处修改）
        self.sample_integral_label = ttk.Label(
            self.sample_frame, 
            text="Integral: Not Calc.",
            foreground="#2c3e50",
            font=('微软雅黑', 12)
        )
        self.sample_integral_label.pack(pady=2)
        
        # 状态标签
        self.status_label = ttk.Label(self.left_panel, text="Ready✔️")
        self.status_label.pack(pady=5)

    def setup_integration_params(self):
        """积分参数设置（独立参数版本）"""
        param_frame = ttk.LabelFrame(self.left_panel, text="Integration Parameters")
        param_frame.pack(fill=tk.X, pady=5)

        # ======== 标液参数 ========
        ttk.Label(param_frame, text="▣ Std.Params", font=('微软雅黑', 10, 'bold')).grid(row=0, column=0, columnspan=6, pady=3, sticky=tk.W)
        
        ttk.Label(param_frame, text="Method:", font=('微软雅黑', 10)).grid(row=1, column=0, padx=2)
        self.std_method = ttk.Combobox(param_frame, values=['trapz', 'simpson'], width=8)
        self.std_method.set('trapz')
        self.std_method.grid(row=1, column=1, padx=2)
        
        ttk.Label(param_frame, text="Start WL:", font=('微软雅黑', 10)).grid(row=1, column=2, padx=2)
        self.std_start_wl = ttk.Entry(param_frame, width=8)
        self.std_start_wl.insert(0, "380.0")
        self.std_start_wl.grid(row=1, column=3, padx=2)
        
        ttk.Label(param_frame, text="End WL:", font=('微软雅黑', 10)).grid(row=1, column=4, padx=2)
        self.std_end_wl = ttk.Entry(param_frame, width=8)
        self.std_end_wl.insert(0, "620.0")
        self.std_end_wl.grid(row=1, column=5, padx=2)

        # ======== 样品参数 ========
        ttk.Label(param_frame, text="▣ Sample Params", font=('微软雅黑', 10, 'bold')).grid(row=2, column=0, columnspan=6, pady=3, sticky=tk.W)
        
        ttk.Label(param_frame, text="Method:", font=('微软雅黑', 10)).grid(row=3, column=0, padx=2)
        self.sample_method = ttk.Combobox(param_frame, values=['trapz', 'simpson'], width=8)
        self.sample_method.set('trapz')
        self.sample_method.grid(row=3, column=1, padx=2)
        
        ttk.Label(param_frame, text="Start WL:", font=('微软雅黑', 10)).grid(row=3, column=2, padx=2)
        self.sample_start_wl = ttk.Entry(param_frame, width=8)
        self.sample_start_wl.insert(0, "380.0")
        self.sample_start_wl.grid(row=3, column=3, padx=2)
        
        ttk.Label(param_frame, text="End WL:", font=('微软雅黑', 10)).grid(row=3, column=4, padx=2)
        self.sample_end_wl = ttk.Entry(param_frame, width=8)
        self.sample_end_wl.insert(0, "620.0")
        self.sample_end_wl.grid(row=3, column=5, padx=2)

    def _setup_formula_tab(self, parent):
        """计算公式标签页"""
        formula_text = """Relative Quantum Yield Calculation Formula:
        QYs = QYr × (Is/Ir) × (Ar/As) × (Ns²/Nr²)

        其中：
        QY = Quantum Yield
        I = Integrated Intensity
        A = Absorbance
        N = Solvent Refractive Index
        s = Test Sample
        r = Ref.Sample"""
        
        formula_label = ttk.Label(
            parent,
            text=formula_text,
            font=('Consolas', 12),  # 使用等宽字体
            justify=tk.LEFT,
            anchor=tk.NW
        )
        formula_label.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)

    def _setup_solvent_tab(self, parent):
        """溶剂折射率标签页"""
        solvents = [
            ("Pure Water", 1.3330),
            ("Ethanol", 1.3616),
            ("Tetrahydrofuran(THF)", 1.4073),
            ("N,N-Dimethylformamide(DMF)", 1.4305),
            ("Trichloromethane", 1.4467),
            ("Acetonitrile", 1.3441),
            ("Dimethyl Sulfoxide(DMSO)", 1.4770),
            ("Methanol", 1.3285),
            ("Ethyl Acetate", 1.37239)
        ]
        
        tree = ttk.Treeview(
            parent,
            columns=("solvent", "n"),
            show="headings",
            height=6
        )
    
        tree.heading("solvent", text="Solvent")
        tree.heading("n", text="Refractive Index(n)")
        tree.column("solvent", width=100, anchor=tk.CENTER)
        tree.column("n", width=80, anchor=tk.CENTER)
        
        for solvent in solvents:
            tree.insert("", tk.END, values=solvent)
    
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 绑定右键菜单
        solvent_menu = tk.Menu(self, tearoff=0)
        solvent_menu.add_command(label="Copy", command=lambda: self.copy_refractive_index(tree))
    
        def on_right_click(event):
            item = tree.identify_row(event.y)
            if item:
                tree.selection_set(item)
                solvent_menu.post(event.x_root, event.y_root)
        
        tree.bind("<Button-3>", on_right_click)

    def copy_refractive_index(self, tree):
        selected = tree.selection()
        if selected:
            value = tree.item(selected[0], 'values')[1]
            self.clipboard_clear()
            self.clipboard_append(str(value))
            self.status_label.config(text="Copied to clipboard")

    def _setup_about_tab(self, parent):
        """关于我标签页"""
        main_horizontal = ttk.Frame(parent)
        main_horizontal.pack(fill=tk.BOTH, expand=True, pady=10)
        
        text_container = ttk.Frame(main_horizontal)
        text_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, anchor=tk.NW)
        # ====================

        about_text = """Welcome to contact the author for bug reports or feature suggestions!
        Developer: Vicne
        Email: 1604178443@qq.com
        Version: v1.5.3_En
        Release Date: 2025-04-03
        Note: Calculation results are for reference only. Use actual measurements as the standard.
        This software is open source for personal use only and cannot be used commercially.
        """

          
        text_label = ttk.Label(
            text_container,
            text=about_text,
            font=('微软雅黑', 12),  
            justify=tk.LEFT,
            anchor=tk.NW
        )
        text_label.pack(fill=tk.BOTH, expand=True, anchor=tk.NW)

        # 开源链接部分（修改变量名为 self.github_link_label）
        link_frame1 = ttk.Frame(text_container)
        link_frame1.pack(fill=tk.X, anchor=tk.NW)
        ttk.Label(link_frame1, 
                text="Open Source：", 
                font=('微软雅黑', 10),
                anchor=tk.W).pack(side=tk.LEFT)

        github_url = "https://github.com/Vicnecxx/qy_calculator"
        self.github_link_label = ttk.Label(  # 使用唯一变量名
            link_frame1,
            text=github_url,
            font=('微软雅黑', 10, 'underline'),  # 添加下划线
            foreground="blue",
            cursor="hand2",
            anchor=tk.W
        )
        self.github_link_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.github_link_label.bind("<Button-1>", lambda e: self.open_url(github_url))

        # CIGIT 链接部分（修改变量名为 self.cigit_link_label）
        link_frame2 = ttk.Frame(text_container)
        link_frame2.pack(fill=tk.X, anchor=tk.NW)

        cigit_url = "http://www.cigit.cas.cn/"
        link_text = ''' Algorithms greet green peaks, twin tides drive tech unveiled 
                 — Chongqing Institute of Green and Intelligent Technology, CAS'''
        self.cigit_link_label = ttk.Label(  # 使用唯一变量名
            link_frame2,
            text=link_text,
            font=('微软雅黑', 12),  # 补全下划线样式
            foreground="blue",
            cursor="hand2",
            anchor=tk.W
        )
        self.cigit_link_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cigit_link_label.bind("<Button-1>", lambda e: self.open_url(cigit_url))

    def open_url(self, url):
        import webbrowser
        # 自动补全协议头
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url
        try:
            webbrowser.open_new(url)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open link:\n{str(e)}")

    def setup_qy_inputs(self):
        """输入字段配置"""
        input_frame = ttk.Frame(self.right_panel)
        input_frame.pack(fill=tk.X, pady=10)

        self.inputs = {
            'qy_ref': ttk.Entry(input_frame, width=18),
            'i_ref': ttk.Entry(input_frame, width=18),
            'i_sample': ttk.Entry(input_frame, width=18),
            'a_ref': ttk.Entry(input_frame, width=18),
            'a_sample': ttk.Entry(input_frame, width=18),
            'n_sample': ttk.Entry(input_frame, width=18),
            'n_ref': ttk.Entry(input_frame, width=18)
        }

        # 设置默认值
        defaults = {
            'qy_ref': "0.54",
            'a_ref': "0.050",
            'n_sample': "1.3330",
            'n_ref': "1.3330"
        }
        for key, value in defaults.items():
            self.inputs[key].insert(0, value)

        # 标签配置
        labels = [
            ("Reference QY (qy_ref)", 0, 0),
            ("Reference integrated intensity (i_ref)", 0, 1),
            ("Sample integrated intensity (i_sample)", 0, 2),
            ("Reference absorbance (a_ref)", 1, 0),
            ("Sample absorbance (a_sample)", 1, 1),
            ("Smp.Solvent RI (n_sample)", 2, 0),
            ("Std.Solvent RI (n_ref)", 2, 1)
        ]

        for text, row, col in labels:
            key = text.split('(')[1].split(')')[0].strip()
            ttk.Label(
                input_frame, 
                text=text.split('(')[0].strip()+" :",
                font=('微软雅黑', 12)  # 显式指定字体
            ).grid(row=row, column=col*2, padx=5, pady=5, sticky=tk.E)
            self.inputs[key].grid(row=row, column=col*2+1, padx=5, pady=5)

    def setup_action_buttons(self):
        """操作按钮配置"""
        btn_frame = ttk.Frame(self.right_panel)
        btn_frame.pack(pady=15)

        ttk.Button(btn_frame, text="Calculate", command=self.calculate_qy, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Clear", command=self.clear_inputs).pack(side=tk.LEFT, padx=10)
        
        # 结果标签
        self.result_label = ttk.Label(self.right_panel, 
                                    text="Waiting for calculation...", 
                                    font=('微软雅黑', 12, 'bold'),
                                    foreground="#2c3e50")
        self.result_label.pack(pady=10)

    def setup_history_panel(self):
        """历史记录面板"""
        self.history_panel = ttk.LabelFrame(self, text="Calculation History")
        self.history_panel.pack(fill=tk.BOTH, expand=True, padx=5, pady=5, side=tk.BOTTOM)

        columns = ("name", "timestamp", "qy_ref", "qy_sample", "method", "range")
        self.history_tree = ttk.Treeview(
            self.history_panel,
            columns=columns,
            show="headings",
            height=8,
            selectmode="browse"
        )

        col_names = ["Sample Name", "Time", "Ref.QY", "Sample QY", "Method", "Wavelength Range"]
        col_widths = [120, 150, 80, 100, 80, 120]
        for col, name, width in zip(columns, col_names, col_widths):
            self.history_tree.heading(col, text=name)
            self.history_tree.column(col, width=width, anchor='center')

        scrollbar = ttk.Scrollbar(self.history_panel, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.history_panel.grid_rowconfigure(0, weight=1)
        self.history_panel.grid_columnconfigure(0, weight=1)

    def setup_context_menu(self):
        """上下文菜单配置"""
        self.file_menu = tk.Menu(self, tearoff=0)
        self.file_menu.add_command(label="Clear File", command=self.clear_file, font=('Arial', 10))
        
        self.history_menu = tk.Menu(self, tearoff=0)
        self.history_menu.add_command(label="Rename", command=self.rename_sample, font=('Arial', 10))
        self.history_menu.add_command(label="Delete Record", command=self.delete_record, font=('Arial', 10))
        self.history_menu.add_separator()
        self.history_menu.add_command(label="Clear History", command=self.clear_history, font=('Arial', 10))
        self.history_menu.add_command(label="Export Table", command=self.export_history, font=('Arial', 10))
        
        self.std_drop_area.bind("<Button-3>", lambda e: self.show_file_menu(e, "std"))
        self.sample_drop_area.bind("<Button-3>", lambda e: self.show_file_menu(e, "sample"))
        self.history_tree.bind("<Button-3>", self.show_history_menu)

    def show_file_menu(self, event, file_type):
        if self.file_data[file_type] is not None:
            self.current_file_type = file_type
            self.file_menu.post(event.x_root, event.y_root)

    def show_history_menu(self, event):
        item = self.history_tree.identify_row(event.y)
        if item:
            self.history_tree.selection_set(item)
            self.history_menu.post(event.x_root, event.y_root)

    def clear_file(self):
        """清除文件时重置显示"""
        if hasattr(self, 'current_file_type'):
            file_type = self.current_file_type
            self.file_data[file_type] = None
            getattr(self, f"{file_type}_file_label").config(text="No file selected")
            self.status_label.config(text=f"{'Standards' if file_type=='std' else 'Sample'} file cleared")
            if file_type == "std":
                self.inputs['i_ref'].delete(0, tk.END)
            else:
                self.inputs['i_sample'].delete(0, tk.END)
            # 清除对应积分结果
            if file_type == "std":
                self.std_integral_label.config(text="Integral: Not calculated", foreground="#2c3e50")
            else:
                self.sample_integral_label.config(text="Integral: Not calculated", foreground="#2c3e50")

    def on_drop(self, event, file_type):
        file_path = event.data.strip("{}")
        if file_path.endswith('.xls'):
            self.process_file(file_path, file_type)
        else:
            self.status_label.config(text="Only .xls files are supported")

    def select_file(self, file_type):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls")])
        if file_path:
            self.process_file(file_path, file_type)

    def process_file(self, file_path, file_type):
        try:
            df = extract_data(file_path)
            self.file_data[file_type] = df
            
            short_path = file_path.split('/')[-1] if '/' in file_path else file_path
            getattr(self, f"{file_type}_file_label").config(text=short_path)
            self.status_label.config(text=f"{'Standard' if file_type=='std' else 'Sample'} file loaded")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="File processing failed")

    def calculate_integral(self, file_type):
        """手动计算积分"""
        try:
            if self.file_data[file_type] is None:
                raise ValueError("Please upload file first")
            
            method = self.std_method.get() if file_type == "std" else self.sample_method.get()
            start = float(self.std_start_wl.get()) if file_type == "std" else float(self.sample_start_wl.get())
            end = float(self.std_end_wl.get()) if file_type == "std" else float(self.sample_end_wl.get())
            target_entry = 'i_ref' if file_type == "std" else 'i_sample'

            df = self.file_data[file_type]
            integral = calculate_integration(df, start_wl=start, end_wl=end, method=method)
            
            self.inputs[target_entry].delete(0, tk.END)
            self.inputs[target_entry].insert(0, f"{integral:.4f}")
            self.status_label.config(text=f"{'Standard solution' if file_type=='std' else 'Sample'} integration completed")
            if file_type == "std":
                self.std_integral_label.config(text=f"Integral: {integral:.4f}", foreground="#27ae60")
            else:
                self.sample_integral_label.config(text=f"Integral: {integral:.4f}", foreground="#27ae60")
            
        except Exception as e:
            messagebox.showerror("Calculation Error", str(e))
            self.status_label.config(text="Integration failed")
            if file_type == "std":
                self.std_integral_label.config(text=f"Calculation failed: {str(e)}", foreground="#e74c3c")
            else:
                self.sample_integral_label.config(text=f"Calculation failed: {str(e)}", foreground="#e74c3c")

    def calculate_qy(self):
        """量子产率计算"""
        try:
            params = {key: float(entry.get()) for key, entry in self.inputs.items()}
            params['n_ref'] = params['n_ref'] if params['n_ref'] else 1.3330  # 默认折射率

            qy_sample = (params['qy_ref'] *
                        (params['i_sample'] / params['i_ref']) *
                        (params['a_ref'] / params['a_sample']) *
                        (params['n_sample']**2 / params['n_ref']**2)) * 100  # 乘以100转换为百分比

            sample_name = f"Sample{self.sample_counter}"
            self.sample_counter += 1

            self.history_tree.insert("", tk.END, values=(
                sample_name,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                f"{params['qy_ref']:.2f}",
                f"{qy_sample:.2f}%",  # 添加百分号
                self.sample_method.get(),
                f"{self.sample_start_wl.get()}-{self.sample_end_wl.get()}nm"
            ))

            self.result_label.config(
                text=f"Quantum yield result: {qy_sample:.2f}%",
                foreground="#27ae60" if qy_sample > 0 else "#e74c3c"
            )

        except ValueError:
            messagebox.showerror("Input Error", "Please check input number format")
            self.result_label.config(text="Invalid input format", foreground="#e74c3c")
        except ZeroDivisionError:
            messagebox.showerror("Calculation Error", "Division by zero not allowed")
            self.result_label.config(text="Error: Division by zero", foreground="#e74c3c")

    def clear_inputs(self):
        """清空所有输入字段并恢复默认值"""
        self.file_data = {"std": None, "sample": None}
        self.std_file_label.config(text="No file selected")
        self.sample_file_label.config(text="No file selected")
        
        for entry in self.inputs.values():
            entry.delete(0, tk.END)
        
        defaults = {
            'qy_ref': "0.54",
            'a_ref': "0.050",
            'n_sample': "1.3330",
            'n_ref': "1.3330"
        }
        for key, value in defaults.items():
            self.inputs[key].insert(0, value)
        
        self.status_label.config(text="Inputs Reset")
        self.result_label.config(text="Awaiting Calculation...", foreground="#2c3e50")
        self.std_integral_label.config(text="Integral: Not Calc.", foreground="#2c3e50")
        self.sample_integral_label.config(text="Integral: Not Calc.", foreground="#2c3e50")

    def rename_sample(self):
        selected = self.history_tree.selection()
        if selected:
            current_name = self.history_tree.item(selected, 'values')[0]
            new_name = simpledialog.askstring("Rename", "New Name:", initialvalue=current_name)
            if new_name:
                values = list(self.history_tree.item(selected, 'values'))
                values[0] = new_name
                self.history_tree.item(selected, values=values)

    def delete_record(self):
        selected = self.history_tree.selection()
        if selected:
            self.history_tree.delete(selected)

    def clear_history(self):
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self.sample_counter = 1

    def export_history(self):
        """导出历史记录"""
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if path:
            data = []
            for item in self.history_tree.get_children():
                values = list(self.history_tree.item(item)['values'])
                values[3] = float(values[3].strip('%'))
                data.append(values)
            
            df = pd.DataFrame(data, columns=["Sample Name", "Time", "Ref.QY", "Sample QY", "Method", "λ Range"])
            if path.endswith('.csv'):
                df.to_csv(path, index=False)
            else:
                df.to_excel(path, index=False)
            messagebox.showinfo("SUCCESS", f"Data saved to:\n{path}")

if __name__ == "__main__":
    app = QYCalculatorApp()
    app.mainloop()