
# -*- coding: utf-8 -*- 
# 荧光数据文件提取函数

import pandas as pd
#文件提取函数，默认不打印荧光图谱数据参数，若需要打印请设置参数为print_params =True，函数接口为文件路径，返回值是数据表格
def extract_data(file_path, print_params =False):
    # 读取整个Excel表格
    try:
        df_raw = pd.read_excel(file_path, header=None, engine="xlrd")
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

    # 提取仪器参数
    params = {
        "Data mode:": None, "EX WL:": None, "EM  Start WL:": None,
        "EM  End WL:": None, "Scan speed:": None, "Delay:": None,
        "EX Slit:": None, "EM Slit:": None, "PMT Voltage:": None,
        "Response:": None, "Corrected spectra:": None, "Shutter control:": None
    }

    # 遍历行查找参数
    data_start_row = None
    for idx, row in df_raw.iterrows():
        if row[0] in params:
            params[row[0]] = row[1]
        # 定位数据起始行（"nm"和"Data"列头）
        if row[0] == "nm" and row[1] == "Data":
            data_start_row = idx + 1
            break

    # 检查是否找到数据起始行
    if data_start_row is None:
        print("错误: 数据起始行未被成功定位。")
        return None

    # 提取数据点
    df_data = df_raw.iloc[data_start_row:, [0, 1]]
    df_data.columns = ["nm", "Data"]
    df_data = df_data.dropna(subset=["nm"])  # 删除空行

    # 打印参数
    if print_params:
        print("光谱仪器参数:")
        for key, value in params.items():
            if value is None:
                print(f"错误: 参数 {key} 没有被成功提取。")
            else:
                print(f"{key}\t{value}")

    return df_data.reset_index(drop=True)

# 运行测试
if __name__ == "__main__":
    file_path = "D:\\\Xuan\Desktop\\相对量子产率计算软件\\test.xls"
    df_data = extract_data(file_path, print_params=True)
    if df_data is not None:
        print("\n数据表格:")
        print(df_data)