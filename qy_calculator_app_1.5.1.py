# -*- coding: utf-8 -*-
# qy_calculator_app 1.5.1
# Created by Vince on 2025/03/28
# Last modified by Vince on 2025/03/28
# 修复了计算历史数据无法复制修改名称的bug

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from tkinterdnd2 import TkinterDnD, DND_ALL
import pandas as pd
from datetime import datetime
from extract_data import extract_data
from calculate_integration import calculate_integration

class QYCalculatorApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("碳点相对量子产率计算平台 1.4.1")
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
        self.left_panel = ttk.LabelFrame(self.main_frame, text="数据源配置", width=450)
        self.left_panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=5, pady=5)
        # 文件拖拽区
        self.setup_file_areas()
        
        # 积分参数设置
        self.setup_integration_params()

    def setup_right_panel(self):
        """右侧面板初始化"""
        self.right_panel = ttk.LabelFrame(self.main_frame, text="相对量子产率计算", width=750)
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 输入字段
        self.setup_qy_inputs()
        
        # 操作按钮
        self.setup_action_buttons()
        
        # ========== 新增信息展示区 ==========
        info_frame = ttk.LabelFrame(self.right_panel, text="参考资料")
        info_frame.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)

        # 创建Notebook标签页
        self.info_notebook = ttk.Notebook(info_frame)
        self.info_notebook.pack(fill=tk.BOTH, expand=True)

        # 公式标签页
        formula_tab = ttk.Frame(self.info_notebook)
        self._setup_formula_tab(formula_tab)
        self.info_notebook.add(formula_tab, text="计算公式")

        # 溶剂折射率标签页
        solvent_tab = ttk.Frame(self.info_notebook)
        self._setup_solvent_tab(solvent_tab)
        self.info_notebook.add(solvent_tab, text="溶剂折射率")

        # 开发者信息标签页
        about_tab = ttk.Frame(self.info_notebook)
        self._setup_about_tab(about_tab)
        self.info_notebook.add(about_tab, text="软件信息")
        
        # 欢迎加入我们交流学习
        welcome_tab = ttk.Frame(self.info_notebook)
        self._setup_welcome_tab(welcome_tab)
        self.info_notebook.add(welcome_tab, text="欢迎交流")

    def setup_file_areas(self):
        """文件拖拽区配置"""
        # ========== 标液文件区域 ==========
        self.std_frame = ttk.LabelFrame(self.left_panel, text="参考标液文件")
        self.std_frame.pack(fill=tk.X, pady=5)
        
        # 文件操作组件
        std_control_frame = ttk.Frame(self.std_frame)
        std_control_frame.pack(fill=tk.X)
        
        ttk.Button(std_control_frame, text="浏览文件", width=8,
                  command=lambda: self.select_file("std")).pack(side=tk.LEFT, padx=2)
        ttk.Button(std_control_frame, text="计算积分", width=8,
                  command=lambda: self.calculate_integral("std")).pack(side=tk.LEFT, padx=2)
        
        # 拖拽区域
        self.std_drop_area = ttk.Label(
            self.std_frame,
            text="拖拽标液.xls文件至此",
            relief="groove",
            anchor="center"
        )
        self.std_drop_area.pack(pady=5, fill=tk.X)
        self.std_drop_area.drop_target_register(DND_ALL)
        self.std_drop_area.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, "std"))
        
        # 文件状态显示
        self.std_file_label = ttk.Label(self.std_frame, text="未选择文件")
        self.std_file_label.pack(pady=2)

        # 新增标液积分结果显示（第1处修改）
        self.std_integral_label = ttk.Label(
            self.std_frame, 
            text="积分结果：未计算",
            foreground="#2c3e50",
            font=('微软雅黑', 12)
        )
        self.std_integral_label.pack(pady=2)
        # ========== 样品文件区域 ==========
        self.sample_frame = ttk.LabelFrame(self.left_panel, text="待测样品文件")
        self.sample_frame.pack(fill=tk.X, pady=5)

        # 文件操作组件
        sample_control_frame = ttk.Frame(self.sample_frame)
        sample_control_frame.pack(fill=tk.X)
        
        ttk.Button(sample_control_frame, text="浏览文件", width=8,
                  command=lambda: self.select_file("sample")).pack(side=tk.LEFT, padx=2)
        ttk.Button(sample_control_frame, text="计算积分", width=8,
                  command=lambda: self.calculate_integral("sample")).pack(side=tk.LEFT, padx=2)

        # 拖拽区域
        self.sample_drop_area = ttk.Label(
            self.sample_frame,
            text="拖拽样品.xls文件至此",
            relief="groove",
            anchor="center"
        )
        self.sample_drop_area.pack(pady=5, fill=tk.X)
        self.sample_drop_area.drop_target_register(DND_ALL)
        self.sample_drop_area.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, "sample"))
        
        # 文件状态显示
        self.sample_file_label = ttk.Label(self.sample_frame, text="未选择文件")
        self.sample_file_label.pack(pady=2)
    
        # 新增样品积分结果显示（第2处修改）
        self.sample_integral_label = ttk.Label(
            self.sample_frame, 
            text="积分结果：未计算",
            foreground="#2c3e50",
            font=('微软雅黑', 12)
        )
        self.sample_integral_label.pack(pady=2)
        
        # 状态标签
        self.status_label = ttk.Label(self.left_panel, text="就绪")
        self.status_label.pack(pady=5)

    def setup_integration_params(self):
        """积分参数设置（独立参数版本）"""
        param_frame = ttk.LabelFrame(self.left_panel, text="积分参数设置")
        param_frame.pack(fill=tk.X, pady=5)

        # ======== 标液参数 ========
        ttk.Label(param_frame, text="▣ 标液参数", font=('微软雅黑', 10, 'bold')).grid(row=0, column=0, columnspan=6, pady=3, sticky=tk.W)
        
        ttk.Label(param_frame, text="积分方法:", font=('微软雅黑', 10)).grid(row=1, column=0, padx=2)
        self.std_method = ttk.Combobox(param_frame, values=['trapz', 'simpson'], width=8)
        self.std_method.set('trapz')
        self.std_method.grid(row=1, column=1, padx=2)
        
        ttk.Label(param_frame, text="起始波长:", font=('微软雅黑', 10)).grid(row=1, column=2, padx=2)
        self.std_start_wl = ttk.Entry(param_frame, width=8)
        self.std_start_wl.insert(0, "380.0")
        self.std_start_wl.grid(row=1, column=3, padx=2)
        
        ttk.Label(param_frame, text="结束波长:", font=('微软雅黑', 10)).grid(row=1, column=4, padx=2)
        self.std_end_wl = ttk.Entry(param_frame, width=8)
        self.std_end_wl.insert(0, "620.0")
        self.std_end_wl.grid(row=1, column=5, padx=2)

        # ======== 样品参数 ========
        ttk.Label(param_frame, text="▣ 样品参数", font=('微软雅黑', 10, 'bold')).grid(row=2, column=0, columnspan=6, pady=3, sticky=tk.W)
        
        ttk.Label(param_frame, text="积分方法:", font=('微软雅黑', 10)).grid(row=3, column=0, padx=2)
        self.sample_method = ttk.Combobox(param_frame, values=['trapz', 'simpson'], width=8)
        self.sample_method.set('trapz')
        self.sample_method.grid(row=3, column=1, padx=2)
        
        ttk.Label(param_frame, text="起始波长:", font=('微软雅黑', 10)).grid(row=3, column=2, padx=2)
        self.sample_start_wl = ttk.Entry(param_frame, width=8)
        self.sample_start_wl.insert(0, "380.0")
        self.sample_start_wl.grid(row=3, column=3, padx=2)
        
        ttk.Label(param_frame, text="结束波长:", font=('微软雅黑', 10)).grid(row=3, column=4, padx=2)
        self.sample_end_wl = ttk.Entry(param_frame, width=8)
        self.sample_end_wl.insert(0, "620.0")
        self.sample_end_wl.grid(row=3, column=5, padx=2)

    def _setup_formula_tab(self, parent):
        """计算公式标签页"""
        formula_text = """相对量子产率计算公式：
        QYs = QYr × (Is/Ir) × (Ar/As) × (Ns²/Nr²)

        其中：
        QY = 量子产率
        I = 积分强度
        A = 吸光度
        N = 溶剂折射率
        s = 待测样品
        r = 参考标液"""
        
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
            ("水", 1.3330),
            ("乙醇", 1.3616),
            ("四氢呋喃(THF)", 1.4073),
            ("N，N-二甲基甲酰胺(DMF)", 1.4305),
            ("三氯甲烷", 1.4467),
            ("乙腈", 1.3441),
            ("二甲基亚砜(DMSO)", 1.4770),
            ("甲醇", 1.3285),
            ("乙酸乙酯", 1.37239)
        ]
        
        tree = ttk.Treeview(
            parent,
            columns=("solvent", "n"),
            show="headings",
            height=6
        )
    
        tree.heading("solvent", text="溶剂")
        tree.heading("n", text="折射率(n)")
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
        solvent_menu.add_command(label="复制", command=lambda: self.copy_refractive_index(tree))
    
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
            self.status_label.config(text="已复制到剪贴板")    

    def _setup_about_tab(self, parent):
        """关于我标签页"""
        main_horizontal = ttk.Frame(parent)
        main_horizontal.pack(fill=tk.BOTH, expand=True, pady=10)
        
        text_container = ttk.Frame(main_horizontal)
        text_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, anchor=tk.NW)
        # ========== 文本内容 ==========
        about_text ="""相关信息：
        该软件由Vince编写，基于Python3开发。
        版本号　　：v1.5.1
        发布日期　：2025-03-29
        注意事项　：计算结果仅供参考， 请以实际测量结果为准！
        软件适用　：可直接用于日立F-7000荧光分光度计导出的数据！
        使用范围　：仅供个人学习交流，禁止用于商业用途！
        """
        
        text_label = ttk.Label(
            text_container,
            text=about_text,
            font=('微软雅黑', 12),  
            justify=tk.LEFT,
            anchor=tk.NW
        )
        text_label.pack(fill=tk.BOTH, expand=True, anchor=tk.NW)

        # 开源链接（独立容器保证对齐）
        link_frame = ttk.Frame(text_container)
        link_frame.pack(fill=tk.X, anchor=tk.NW)
        ttk.Label(link_frame, 
                 text="软件开源：", 
                 font=('微软雅黑', 12),
                 anchor=tk.W).pack(side=tk.LEFT)
        
        url = "https://github.com/Vicnecxx/qy_calculator"
        self.link_label = ttk.Label(
            link_frame,
            text=url,
            font=('微软雅黑', 12, 'underline'),
            foreground="blue",
            cursor="hand2",
            anchor=tk.W
        )
        self.link_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.link_label.bind("<Button-1>", lambda e: self.open_url(url))

    def open_url(self, url):
        import webbrowser
        webbrowser.open_new(url)

    def _setup_welcome_tab(self, parent):
        """欢迎交流标签页"""
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        welcome_label = ttk.Label(
            main_frame,
            text="欢迎软件Bug反馈、功能建议或加入群聊交流学习！",
            font=('微软雅黑', 14, 'bold'),
            foreground="#2c3e50",
            anchor=tk.CENTER
        )
        welcome_label.pack(pady=10)

        qr_container = ttk.Frame(main_frame)
        qr_container.pack(fill=tk.BOTH, expand=True)

        # 第一张二维码（联系作者）
        contact_frame = ttk.Frame(qr_container)
        contact_frame.pack(side=tk.LEFT, expand=True, padx=(20, 100))
        
        try:
            from PIL import ImageTk, Image
            qr_img1 = Image.open("contact.jpg").resize((150, 150), Image.LANCZOS)
            self.contact_qr = ImageTk.PhotoImage(qr_img1)
            ttk.Label(contact_frame, image=self.contact_qr).pack(pady=5)
            ttk.Label(contact_frame, 
                    text="Vicne", 
                    font=('微软雅黑', 12),
                    foreground="#3498db").pack()
        except Exception as e:
            ttk.Label(contact_frame, 
                    text="[联系二维码加载失败]", 
                    foreground="red").pack()

        # 第二张二维码（交流群）
        group_frame = ttk.Frame(qr_container)
        group_frame.pack(side=tk.RIGHT, expand=True, padx=(40, 20))
        
        try:
            qr_img2 = Image.open("group.jpg").resize((150, 150), Image.LANCZOS)
            self.group_qr = ImageTk.PhotoImage(qr_img2)
            ttk.Label(group_frame, image=self.group_qr).pack(pady=5)
            ttk.Label(group_frame, 
                    text="碳点QQ交流群", 
                    font=('微软雅黑', 12),
                    foreground="#3498db").pack()
        except Exception as e:
            ttk.Label(group_frame, 
                    text="[交流群二维码加载失败]", 
                    foreground="red").pack()

        # 提示文字
        tip_label = ttk.Label(
            main_frame,
            text="请使用手机扫描二维码添加联系方式或加入交流群",
            font=('微软雅黑', 10),
            foreground="#7f8c8d"
        )
        tip_label.pack(pady=15)

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
            'a_ref': "0.49",
            'n_sample': "1.333",
            'n_ref': "1.3333"
        }
        for key, value in defaults.items():
            self.inputs[key].insert(0, value)

        # 标签配置
        labels = [
            ("标液量子产率QY (qy_ref)", 0, 0),
            ("标液积分强度 (i_ref)", 0, 1),
            ("样品积分强度 (i_sample)", 0, 2),
            ("标液吸光度 (a_ref)", 1, 0),
            ("样品吸光度 (a_sample)", 1, 1),
            ("样品溶液折射率 (n_sample)", 2, 0),
            ("标液溶液折射率 (n_ref)", 2, 1)
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

        ttk.Button(btn_frame, text="计算QY", command=self.calculate_qy, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="清空所有输入", command=self.clear_inputs).pack(side=tk.LEFT, padx=10)
        
        # 结果标签
        self.result_label = ttk.Label(self.right_panel, 
                                    text="等待计算...", 
                                    font=('微软雅黑', 12, 'bold'),
                                    foreground="#2c3e50")
        self.result_label.pack(pady=10)

    def setup_history_panel(self):
        """历史记录面板"""
        self.history_panel = ttk.LabelFrame(self, text="计算历史记录")
        self.history_panel.pack(fill=tk.BOTH, expand=True, padx=5, pady=5, side=tk.BOTTOM)

        columns = ("name", "timestamp", "qy_ref", "qy_sample", "method", "range")
        self.history_tree = ttk.Treeview(
            self.history_panel,
            columns=columns,
            show="headings",
            height=8,
            selectmode="browse"
        )

        col_names = ["样品名称", "时间", "参考QY", "样品QY", "积分方法", "波长范围"]
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
        self.file_menu.add_command(label="清除文件", command=self.clear_file, font=('微软雅黑'))
        
        self.history_menu = tk.Menu(self, tearoff=0)
        self.history_menu.add_command(label="修改名称", command=self.rename_sample, font=('微软雅黑', 10))
        self.history_menu.add_command(label="删除记录", command=self.delete_record, font=('微软雅黑', 10))
        self.history_menu.add_separator()
        self.history_menu.add_command(label="清空历史", command=self.clear_history, font=('微软雅黑', 10))
        self.history_menu.add_command(label="导出Excel", command=self.export_history, font=('微软雅黑', 10))
        
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
            getattr(self, f"{file_type}_file_label").config(text="未选择文件")
            self.status_label.config(text=f"{'标液' if file_type=='std' else '样品'}文件已清除")
            if file_type == "std":
                self.inputs['i_ref'].delete(0, tk.END)
            else:
                self.inputs['i_sample'].delete(0, tk.END)
            # 清除对应积分结果
            if file_type == "std":
                self.std_integral_label.config(text="积分结果：未计算", foreground="#2c3e50")
            else:
                self.sample_integral_label.config(text="积分结果：未计算", foreground="#2c3e50")

    def on_drop(self, event, file_type):
        file_path = event.data.strip("{}")
        if file_path.endswith('.xls'):
            self.process_file(file_path, file_type)
        else:
            self.status_label.config(text="仅支持.xls文件")

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
            self.status_label.config(text=f"{'标液' if file_type=='std' else '样品'}文件已加载")
            
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.status_label.config(text="文件处理失败")

    def calculate_integral(self, file_type):
        """手动计算积分"""
        try:
            if self.file_data[file_type] is None:
                raise ValueError("请先上传文件")
            
            method = self.std_method.get() if file_type == "std" else self.sample_method.get()
            start = float(self.std_start_wl.get()) if file_type == "std" else float(self.sample_start_wl.get())
            end = float(self.std_end_wl.get()) if file_type == "std" else float(self.sample_end_wl.get())
            target_entry = 'i_ref' if file_type == "std" else 'i_sample'

            df = self.file_data[file_type]
            integral = calculate_integration(df, start_wl=start, end_wl=end, method=method)
            
            self.inputs[target_entry].delete(0, tk.END)
            self.inputs[target_entry].insert(0, f"{integral:.4f}")
            self.status_label.config(text=f"{'标液' if file_type=='std' else '样品'}积分计算完成")
            if file_type == "std":
                self.std_integral_label.config(text=f"积分结果：{integral:.4f}", foreground="#27ae60")
            else:
                self.sample_integral_label.config(text=f"积分结果：{integral:.4f}", foreground="#27ae60")
            
        except Exception as e:
            messagebox.showerror("计算错误", str(e))
            self.status_label.config(text="积分计算失败")
            if file_type == "std":
                self.std_integral_label.config(text=f"计算失败：{str(e)}", foreground="#e74c3c")
            else:
                self.sample_integral_label.config(text=f"计算失败：{str(e)}", foreground="#e74c3c")

    def calculate_qy(self):
        """量子产率计算"""
        try:
            params = {key: float(entry.get()) for key, entry in self.inputs.items()}
            params['n_ref'] = params['n_ref'] if params['n_ref'] else 1.3333

            qy_sample = (params['qy_ref'] *
                        (params['i_sample'] / params['i_ref']) *
                        (params['a_ref'] / params['a_sample']) *
                        (params['n_sample']**2 / params['n_ref']**2)) * 100  # 乘以100转换为百分比

            sample_name = f"样品{self.sample_counter}"
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
                text=f"量子产率计算结果：{qy_sample:.2f}%", 
                foreground="#27ae60" if qy_sample > 0 else "#e74c3c"
            )

        except ValueError:
            messagebox.showerror("输入错误", "请检查输入数值格式")
            self.result_label.config(text="输入格式错误", foreground="#e74c3c")
        except ZeroDivisionError:
            messagebox.showerror("计算错误", "除数不能为零")
            self.result_label.config(text="计算错误：除数为零", foreground="#e74c3c")

    def clear_inputs(self):
        """清空所有输入字段并恢复默认值"""
        self.file_data = {"std": None, "sample": None}
        self.std_file_label.config(text="未选择文件")
        self.sample_file_label.config(text="未选择文件")
        
        for entry in self.inputs.values():
            entry.delete(0, tk.END)
        
        defaults = {
            'qy_ref': "0.54",
            'a_ref': "0.49",
            'n_sample': "1.333",
            'n_ref': "1.3333"
        }
        for key, value in defaults.items():
            self.inputs[key].insert(0, value)
        
        self.status_label.config(text="输入已重置")
        self.result_label.config(text="等待计算...", foreground="#2c3e50")
        self.std_integral_label.config(text="积分结果：未计算", foreground="#2c3e50")
        self.sample_integral_label.config(text="积分结果：未计算", foreground="#2c3e50")

    def rename_sample(self):
        selected = self.history_tree.selection()
        if selected:
            current_name = self.history_tree.item(selected, 'values')[0]
            new_name = simpledialog.askstring("修改名称", "输入新名称:", initialvalue=current_name)
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
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv")]
        )
        if path:
            data = []
            for item in self.history_tree.get_children():
                values = list(self.history_tree.item(item)['values'])
                values[3] = float(values[3].strip('%'))
                data.append(values)
            
            df = pd.DataFrame(data, columns=["样品名称", "时间", "参考QY", "样品QY", "积分方法", "波长范围"])
            if path.endswith('.csv'):
                df.to_csv(path, index=False)
            else:
                df.to_excel(path, index=False)
            messagebox.showinfo("导出成功", f"数据已保存到:\n{path}")

if __name__ == "__main__":
    app = QYCalculatorApp()
    app.mainloop()
