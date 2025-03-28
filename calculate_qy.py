
#相对量子产率计算
def calculate_qy():
    """
    计算样品的相对量子产率（Quantum Yield, QY）。
    该函数从用户界面获取所需的参考和样品的光学性质数据，
    并计算样品的量子产率。如果输入不合法（如非数字），则显示错误信息。
    """
    try:
        # 获取用户界面输入的参考样品的量子产率
        qy_ref = float(qy_ref_entry.get())
        # 获取用户界面输入的样品和参考的强度
        i_sample = float(i_sample_entry.get())
        i_ref = float(i_ref_entry.get())
        # 获取用户界面输入的参考和样品的吸光度
        a_ref = float(a_ref_entry.get())
        a_sample = float(a_sample_entry.get())
        # 获取用户界面输入的样品的折射率
        n_sample = float(n_sample_entry.get())
        # 获取用户界面输入的参考的折射率，如果为空，则使用默认值
        n_ref_input = n_ref_entry.get()
        
        # 根据用户输入或使用默认值确定参考的折射率
        if n_ref_input.strip() == "":
            n_ref = 1.3333
        else:
            n_ref = float(n_ref_input)

        # 计算样品和参考的强度比
        ratio_intensity = i_sample / i_ref
        # 计算参考和样品的吸光度比
        ratio_absorbance = a_ref / a_sample
        # 计算样品和参考的折射率平方比
        ratio_refractive = (n_sample ** 2) / (n_ref ** 2)
        # 计算样品的相对量子产率
        qy_sample = qy_ref * ratio_intensity * ratio_absorbance * ratio_refractive

        # 更新用户界面显示计算结果
        result_label.config(text=f"待测样品的相对量子产率 QY_sample = {qy_sample:.4f}")

    except ValueError:
        # 如果输入值不是数字，则显示错误信息
        result_label.config(text="输入无效，请检查输入值。")

