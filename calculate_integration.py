# -*- coding: utf-8 -*-
# CALCULATE_INTEGRATION.PY


import pandas as pd
import numpy as np
from extract_data import extract_data

def calculate_integration(df, start_wl=380, end_wl=620, method='trapz'):
    """
    荧光光谱数据积分计算函数
    
    参数：
    df : DataFrame - 包含nm和Data列的数据框
    start_wl : int/float - 积分起始波长（默认380）
    end_wl : int/float - 积分结束波长（默认620）
    method : str - 积分方法：trapz（梯形法，默认,同Oringin计算结果相同）| simpson（辛普森法）
    
    返回：
    integral_value : float - 积分结果
    """
    # 验证输入数据
    if not isinstance(df, pd.DataFrame) or not {'nm', 'Data'}.issubset(df.columns):
        raise ValueError("输入数据格式错误，需要包含nm和Data列的DataFrame")
    
    # 获取有效波长范围
    min_wl = df['nm'].min()
    max_wl = df['nm'].max()
    
    # 自动修正超出范围的请求波长
    start_wl = max(start_wl, min_wl)
    end_wl = min(end_wl, max_wl)
    
    # 筛选积分区间数据
    mask = (df['nm'] >= start_wl) & (df['nm'] <= end_wl)
    interval_df = df[mask].sort_values('nm')
    
    if len(interval_df) < 2:
        raise ValueError("积分区间内数据点不足，至少需要2个数据点")
    
    # 准备积分数据
    x = interval_df['nm'].values
    y = interval_df['Data'].values
    
    # 选择积分方法
    if method == 'trapz':
        integral = np.trapz(y, x)
    elif method == 'simpson':
        from scipy.integrate import simpson
        integral = simpson(y, x)
    else:
        raise ValueError("不支持的积分方法，请选择'trapz'或'simpson'")
    
    return round(integral, 4)
