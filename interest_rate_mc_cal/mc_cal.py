import pandas as pd
def discount_cashflows(result_df, monthly_df, cashflow_start_col=5):
    """
    使用矩阵运算对按列存储的现金流进行折现,三条曲线得到pv,pv_down,pv_up
    
    参数:
    1.result_df: cashflow_cal返回资产现金流表，默认601个月
    2.monthly_df: 月度折现率表（包含600个月的折现因子）
    3.cashflow_start_col: 现金流起始列索引，需要根据cashflow_cal中metadata字典配置数量（账户，代码，名称等）设置
    4.output_file：main中写入excel文件的路径
    
    返回:
    1.result:包含三个情景折现值的DataFramepv,包含pv,pv_down,pv_up
    """
    
    # 提取现金流列（从cashflow_start_col开始到最后一列）
    cashflow_cols = result_df.columns[cashflow_start_col:]
    
    # 确保现金流列数量为601（month=1~601）
    if len(cashflow_cols) != 601:
        raise ValueError(f"现金流列数量应为601，但实际为{len(cashflow_cols)}")
    
    # 提取后600个月的现金流列（month=2~601）
    cashflow_cols = cashflow_cols[1:]  # 跳过第1个月
    
    # 提取三个情景的折现因子（month=1~600）
    discount_factors = monthly_df[[
        'rate_discount',
        'rate_down_discount',
        'rate_up_discount'
    ]].values
    
    # 将现金流数据转换为NumPy数组（形状：[资产数, 600个月]）
    cashflow_matrix = result_df[cashflow_cols].values
    
    # 矩阵乘法计算折现值（每个资产在每个情景下的总折现值）
    discounted_matrix = cashflow_matrix @ discount_factors
    
    # 创建结果DataFrame
    result = pd.DataFrame({
        'account_1': result_df.iloc[:, 0],
        'account_2': result_df.iloc[:, 1],
        'product_type': result_df.iloc[:, 2],
        'bond_code': result_df.iloc[:, 3],
        'bond_name': result_df.iloc[:, 4],
        'pv': discounted_matrix[:, 0],
        'pv_down': discounted_matrix[:, 1],
        'pv_up': discounted_matrix[:, 2]
    })
    
    return result


if __name__=="__main__":
    date = '20241231'
    result_df = pd.read_excel("cashflow_analysis-2412.xlsx",sheet_name = "总现金流")
    monthly_df = pd.read_excel("curve2.xlsx")
    result = discount_cashflows(result_df, monthly_df, cashflow_start_col=5)
    output_file = date+'mc.xlsx'
    result.to_excel(output_file)
    print(f"\n最低基本已计算完成并保存到: {output_file}")