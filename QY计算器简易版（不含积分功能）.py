import tkinter as tk
from tkinter import ttk
import calculate_qy

# 创建主窗口
root = tk.Tk()
root.title("量子产率计算器")

# 创建输入框和标签
qy_ref_label = ttk.Label(root, text="请输入参考物质的量子产率 (QY_ref):")
qy_ref_label.grid(row=0, column=0, padx=5, pady=5)
qy_ref_entry = ttk.Entry(root)
qy_ref_entry.grid(row=0, column=1, padx=5, pady=5)

i_sample_label = ttk.Label(root, text="请输入待测样品的荧光强度积分值 (I_sample):")
i_sample_label.grid(row=1, column=0, padx=5, pady=5)
i_sample_entry = ttk.Entry(root)
i_sample_entry.grid(row=1, column=1, padx=5, pady=5)

i_ref_label = ttk.Label(root, text="请输入参考物质的荧光强度积分值 (I_ref):")
i_ref_label.grid(row=2, column=0, padx=5, pady=5)
i_ref_entry = ttk.Entry(root)
i_ref_entry.grid(row=2, column=1, padx=5, pady=5)

a_ref_label = ttk.Label(root, text="请输入参考物质在激发波长处的吸光度 (A_ref):")
a_ref_label.grid(row=3, column=0, padx=5, pady=5)
a_ref_entry = ttk.Entry(root)
a_ref_entry.grid(row=3, column=1, padx=5, pady=5)

a_sample_label = ttk.Label(root, text="请输入待测样品在激发波长处的吸光度 (A_sample):")
a_sample_label.grid(row=4, column=0, padx=5, pady=5)
a_sample_entry = ttk.Entry(root)
a_sample_entry.grid(row=4, column=1, padx=5, pady=5)

n_sample_label = ttk.Label(root, text="请输入待测样品的溶剂折射率 (n_sample):")
n_sample_label.grid(row=5, column=0, padx=5, pady=5)
n_sample_entry = ttk.Entry(root)
n_sample_entry.grid(row=5, column=1, padx=5, pady=5)

n_ref_label = ttk.Label(root, text="请输入参考物质的溶剂折射率 (默认1.3333，按Enter跳过):")
n_ref_label.grid(row=6, column=0, padx=5, pady=5)
n_ref_entry = ttk.Entry(root)
n_ref_entry.grid(row=6, column=1, padx=5, pady=5)

# 创建计算按钮
calculate_button = ttk.Button(root, text="计算量子产率", command=calculate_qy)
calculate_button.grid(row=7, column=0, columnspan=2, pady=10)

# 创建结果显示标签
result_label = ttk.Label(root, text="")
result_label.grid(row=8, column=0, columnspan=2, pady=10)

# 运行主循环
root.mainloop()
