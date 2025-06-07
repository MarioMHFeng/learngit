import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
from tqdm import tqdm

"""
将输入的现金流根据压力参数进行处理，得到三条折现率曲线
参数：
1.param_df:压力参数，通过data转换成df，data根据监管文设定
2.file_path：季度末最后一个交易日基础利率曲线，
    下载地址：https://yield.chinabond.com.cn/cbweb-mn/yc/bxjInit?locale=zh_CN，参数设定60日，
3.其他插值主要参数: - ultimate_rate: 基础终极利率（%前数字），默认4.5
                    - premium_base_1: 短期（<=20年）基础溢价（%前数字），默认0.45
                    - premium_base_2: 长期（>=41年）基础溢价（%前数字），默认0
返回：
1.monthly_df：三条折现率曲线
"""



#压力参数
# 构建原始参数表
def interpolate_stress_params(param_df, term_col='期限', up_col='利率向上压力参数', 
                             down_col='利率向下压力参数', start_term=20, end_term=40):
    """
    对压力参数表进行插值处理（默认对20-40年进行线性插值）
    
    参数:
    - param_df: 原始参数表DataFrame，包含期限和压力参数列
    - term_col: 期限列的列名
    - up_col: 利率向上压力参数列的列名
    - down_col: 利率向下压力参数列的列名
    - start_term: 插值起始期限（包含）
    - end_term: 插值结束期限（包含）
    
    返回:
    - 插值后的完整参数表DataFrame
    """
    # 确保参数表按期限排序
    param_df = param_df.sort_values(term_col).reset_index(drop=True)
    
    # 获取插值基准点（起始期限和结束期限对应的参数值）
    start_row = param_df[param_df[term_col] == start_term].iloc[0]
    end_row = param_df[param_df[term_col] == end_term].iloc[0]
    
    # 生成需要插值的期限列表
    new_terms = list(range(start_term + 1, end_term))
    
    # 执行线性插值
    interpolated_data = []
    for term in new_terms:
        # 计算插值比例
        ratio = (term - start_term) / (end_term - start_term)
        
        # 线性插值计算向上和向下压力参数
        up_param = start_row[up_col] + ratio * (end_row[up_col] - start_row[up_col])
        down_param = start_row[down_col] + ratio * (end_row[down_col] - start_row[down_col])
        
        interpolated_data.append({
            term_col: term,
            up_col: up_param,
            down_col: down_param
        })
    
    # 创建插值后的DataFrame并与原始数据合并
    interpolated_df = pd.DataFrame(interpolated_data)
    combined_df = pd.concat([param_df, interpolated_df], ignore_index=True)
    combined_df = combined_df.sort_values(term_col).reset_index(drop=True)

    #转换单位
    combined_df[up_col] = combined_df[up_col] / 100
    combined_df[down_col] = combined_df[down_col] / 100

    # 按期限排序并重置索引
    return combined_df

def load_rate_curve(file_path, sheet_name = "Sheet1",term_col='标准期限(年)', rate_col='平均值(%)'):
    """
    加载并预处理利率曲线数据
    
    参数:
    - file_path: Excel文件路径
    - term_col: 期限列的列名
    - rate_col: 利率列的列名
    
    返回:
    - 预处理后的利率曲线DataFrame
    """
    try:
        # 读取数据
        rate_curve_df = pd.read_excel(file_path,sheet_name = sheet_name)
        
        # 检查必要列
        required_cols = [term_col, rate_col]
        missing_cols = [col for col in required_cols if col not in rate_curve_df.columns]
        if missing_cols:
            raise ValueError(f"数据缺少必要的列: {', '.join(missing_cols)}")
        
        # 处理缺失值
        if rate_curve_df[required_cols].isnull().any().any():
            print(f"⚠️ 发现缺失值，已删除包含缺失值的行")
            rate_curve_df = rate_curve_df.dropna(subset=required_cols)
        
        # 筛选整数年期限并转换为数值类型
        rate_curve_df = rate_curve_df[rate_curve_df[term_col].apply(lambda x: isinstance(x, (int, float)) and x.is_integer())]
        rate_curve_df[term_col] = rate_curve_df[term_col].astype(int)
        
        # 重命名列
        rate_curve_df = rate_curve_df.rename(columns={
            term_col: 'date',
            rate_col: 'rate'
        })
        
        # 筛选有效期限（大于0）
        rate_curve_df = rate_curve_df[rate_curve_df['date'] > 0]
        
        # 按期限排序并重置索引
        rate_curve_df = rate_curve_df.sort_values('date').reset_index(drop=True)
        
        # 验证关键期限点是否存在
        critical_terms = [1, 20, 40]
        missing_terms = [t for t in critical_terms if t not in rate_curve_df['date'].values]
        if missing_terms:
            print(f"⚠️ 警告：缺少关键期限点: {missing_terms}")
        
        print(f"✅ 成功加载利率曲线，包含 {len(rate_curve_df)} 个期限点")
        print(f"   期限范围: {rate_curve_df['date'].min()}年 ~ {rate_curve_df['date'].max()}年")
        
        return rate_curve_df
    
    except Exception as e:
        print(f"❌ 加载利率曲线失败: {str(e)}")
        return None

def validate_rate_curve(rate_curve_df):
    """验证利率曲线数据的有效性"""
    if rate_curve_df is None or len(rate_curve_df) == 0:
        return False, "利率曲线数据为空"
    
    # 检查是否有重复期限
    if rate_curve_df['date'].duplicated().any():
        return False, "利率曲线包含重复的期限"
    
    # 检查利率是否为正数
    if (rate_curve_df['rate'] < 0).any():
        return False, "利率曲线包含负利率"
    
    return True, "利率曲线验证通过"


def apply_stress_to_curve(rate_curve_df, stress_param_df, base_rate_col='rate', 
                         term_col='date', up_param_col='利率向上压力参数', 
                         down_param_col='利率向下压力参数'):
    """
    应用压力参数生成rate_up和rate_down（不合并参数表）
    
    参数:
    - rate_curve_df: 利率曲线DataFrame
    - stress_param_df: 压力参数DataFrame
    - base_rate_col: 基准利率列名
    - term_col: 期限列名
    
    返回:
    - 包含原始数据及rate_up、rate_down的DataFrame
    """
    # 创建压力参数映射字典（期限 → 压力参数）
    up_param_map = dict(zip(stress_param_df['期限'], stress_param_df[up_param_col]))
    down_param_map = dict(zip(stress_param_df['期限'], stress_param_df[down_param_col]))
    
    # 应用压力测试（直接根据期限查找参数）
    rate_curve_df = rate_curve_df.copy()  # 避免修改原始数据
    rate_curve_df['rate_up'] = rate_curve_df.apply(
        lambda row: row[base_rate_col] * (1 + up_param_map.get(row[term_col], 0)), 
        axis=1
    )
    rate_curve_df['rate_down'] = rate_curve_df.apply(
        lambda row: row[base_rate_col] * (1 + down_param_map.get(row[term_col], 0)), 
        axis=1
    )
    
    return rate_curve_df

def interpolate_rate_curve(rate_curve_df, stress_param_df, ultimate_rate=4.5, premium_base_1=0.45,premium_base_2=0):
    """
    对利率曲线进行两次插值、溢价调整并转远期计算
    计算 rate_up 和 rate_down 时使用调整后的 ultimate_rate
    
    参数:
    - rate_curve_df: 包含 rate、rate_up、rate_down 列的 DataFrame
    - stress_param_df: 压力参数表，用于获取调整系数
    - ultimate_rate: 基础终极利率（%前数字）
    - premium_base_1: 短期（<=20年）基础溢价（%前数字）
    - premium_base_2: 长期（>=41年）基础溢价（%前数字）
    
    返回:
    - 处理后的利率曲线 DataFrame
    """
    rate_curve_df = rate_curve_df.sort_values('date').copy()
    
    # 创建压力参数映射字典
    up_param_map = dict(zip(stress_param_df['期限'], stress_param_df['利率向上压力参数']))
    down_param_map = dict(zip(stress_param_df['期限'], stress_param_df['利率向下压力参数']))
    
    # 对每列（rate、rate_up、rate_down）执行完整处理
    for col_prefix in ['rate', 'rate_up', 'rate_down']:
        # 获取当前处理的列名
        base_col = col_prefix
        rate_1_col = f"{col_prefix}_1"
        rate_2_col = f"{col_prefix}_2"
        rate_3_col = f"{col_prefix}_3"
        rate_4_col = f"{col_prefix}_4"
        
        # 确保基础列存在
        if base_col not in rate_curve_df.columns:
            print(f"⚠️ 警告：DataFrame 中不存在列 '{base_col}'，跳过处理")
            continue
        
        # 计算当前列的 ultimate_rate（仅 rate_up 和 rate_down 需要调整）
        if col_prefix == 'rate_up':
            # 向上压力：ultimate_rate 乘上 (1 + 对应期限的向上压力参数)
            adjusted_ultimate_rate = ultimate_rate * (1 + up_param_map.get(40, 0))
        elif col_prefix == 'rate_down':
            # 向下压力：ultimate_rate 乘上 (1 + 对应期限的向下压力参数)
            adjusted_ultimate_rate = ultimate_rate * (1 + down_param_map.get(40, 0))
        else:
            # 基础情况：使用原始 ultimate_rate
            adjusted_ultimate_rate = ultimate_rate
        
        # 第一次插值（rate_1）
        rate_curve_df[rate_1_col] = None
        rate_20 = rate_curve_df[rate_curve_df['date'] == 20][base_col].values[0]
        
        for i, row in rate_curve_df.iterrows():
            term = row['date']
            rate = row[base_col]
            
            if term <= 20:
                rate_1 = rate
            elif 20 < term < 41:
                rate_1 = rate_20 + (adjusted_ultimate_rate - rate_20) * (term - 20) / 20
            else:
                rate_1 = adjusted_ultimate_rate
            
            rate_curve_df.loc[i, rate_1_col] = rate_1
        
        # 第二次插值（rate_2）
        rate_curve_df[rate_2_col] = None
        
        for i, row in rate_curve_df.iterrows():
            term = row['date']
            rate = row[base_col]
            rate_1 = row[rate_1_col]
            
            if term < 20 or term >= 41:
                rate_2 = rate_1
            else:
                rate_2 = rate_1 * (term - 20) / 20 + rate * (40 - term) / 20
            
            rate_curve_df.loc[i, rate_2_col] = rate_2
        
        # 溢价调整（rate_3）
        rate_curve_df[rate_3_col] = None
        
        for i, row in rate_curve_df.iterrows():
            term = row['date']
            rate_2 = row[rate_2_col]
            
            if term <= 20:
                rate_3 = rate_2 + premium_base_1
            elif term >= 41:
                rate_3 = rate_2 + premium_base_2
            else:
                rate_3 = rate_2 + (premium_base_1-premium_base_2) * (40 - term) / 20
            
            rate_curve_df.loc[i, rate_3_col] = rate_3
        
        # 转远期计算（rate_4）
        rate_curve_df[rate_4_col] = None
        
        # 第一年：rate_4等于rate_3
        first_row = rate_curve_df.iloc[0]
        rate_curve_df.loc[first_row.name, rate_4_col] = round(first_row[rate_3_col], 2)
        
        # 后续年份：根据公式计算远期利率
        for i in range(1, len(rate_curve_df)):
            current_row = rate_curve_df.iloc[i]
            prev_row = rate_curve_df.iloc[i-1]
            
            current_date = current_row['date']
            prev_date = prev_row['date']
            current_rate3 = current_row[rate_3_col]
            prev_rate3 = prev_row[rate_3_col]
            
            # 计算公式：((1+current_rate3/100)^current_date / (1+prev_rate3/100)^prev_date - 1) * 100
            forward_rate = (((1 + current_rate3 / 100) ** current_date) / 
                            ((1 + prev_rate3 / 100) ** prev_date) - 1) * 100
            
            # 保留两位小数
            rate_curve_df.loc[current_row.name, rate_4_col] = round(forward_rate, 2)
    
    return rate_curve_df

def annual_to_monthly(rate_curve_df, rate_cols=['rate_4', 'rate_up_4', 'rate_down_4'], max_years=50):
    """
    将年度利率曲线转换为月度利率曲线，并计算对应的折现率
    
    参数:
    - rate_curve_df: 包含年度利率列的DataFrame
    - rate_cols: 需要转换的年度利率列名列表
    - max_years: 最大转换年数（默认50年=600个月）
    
    返回:
    - 包含月度利率和折现率的DataFrame
    """
    # 创建月度期限列表
    total_months = max_years * 12
    months = list(range(1, total_months + 1))
    monthly_df = pd.DataFrame({'month': months})
    
    # 从年度曲线中提取年份信息
    rate_curve_df['year'] = rate_curve_df['date'].astype(int)

    #排序输出列
    result_columns = ['month']
    monthly_rate_columns = []
    discount_factor_columns = []
    
    # 对每个需要转换的利率列进行处理
    for rate_col in rate_cols:
        # 创建对应的月度利率列名和折现率列名
        monthly_rate_col = rate_col.replace('_4', '_monthly')
        discount_factor_col = rate_col.replace('_4', '_discount')
        
        # 添加到结果列组
        monthly_rate_columns.append(monthly_rate_col)
        discount_factor_columns.append(discount_factor_col)

        # 初始化月度利率列和折现率列
        monthly_df[monthly_rate_col] = None
        monthly_df[discount_factor_col] = None
        
        # 累积折现因子（初始值为1）
        cumulative_discount = 1.0
        
        # 对每个月进行处理
        for _, row in monthly_df.iterrows():
            month = row['month']
            year = (month - 1) // 12 + 1  # 计算对应的年份
            
            # 查找该年的年度利率
            year_row = rate_curve_df[rate_curve_df['year'] == year]
            
            if not year_row.empty:
                annual_rate = year_row[rate_col].values[0]
                
                # 月度利率计算公式: (1 + 年利率/100)^(1/12) - 1
                monthly_rate = ((1 + annual_rate / 100) ** (1/12) - 1) 
                monthly_df.loc[monthly_df['month'] == month, monthly_rate_col] = monthly_rate
            
                # 计算当月折现因子: 1 / (1 + 月利率/100)
                monthly_discount = 1 / (1 + monthly_rate)
                
                # 累积折现因子: 累乘每个月的折现因子
                cumulative_discount *= monthly_discount
                monthly_df.loc[monthly_df['month'] == month, discount_factor_col] = cumulative_discount

    result_columns += monthly_rate_columns + discount_factor_columns
    
    return monthly_df[result_columns]

if __name__=="__main__":
    data = {
    '期限': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 40, 41,42,43,44,45,46,47,48,49,50],
    '利率向上压力参数': [97, 76, 68, 65, 66, 61, 55, 53, 52, 50, 49, 47, 45, 42, 41, 39, 38, 38, 38, 37, 17, 17,17,17,17,17,17,17,17,17,17],
    '利率向下压力参数': [-71, -66, -61, -54, -48, -45, -42, -39, -36, -34, -32, -30, -28, -27, -25, -24, -23, -23, -23, -23, -11, -11,-11,-11,-11,-11,-11,-11,-11,-11,-11]
    }
    param_df = pd.DataFrame(data)
    rate_curve_df = load_rate_curve('curve_20241231.xlsx',sheet_name = "Export")
    combined_df = interpolate_stress_params(param_df, term_col='期限', up_col='利率向上压力参数', 
                                            down_col='利率向下压力参数', start_term=20, end_term=40)
    rate_curve_df=apply_stress_to_curve(rate_curve_df, combined_df, base_rate_col='rate', 
                             term_col='date', up_param_col='利率向上压力参数', 
                             down_param_col='利率向下压力参数')
    rate_curve_df = interpolate_rate_curve(rate_curve_df, combined_df, ultimate_rate=4.5, premium_base_1=0.45,premium_base_2=0)
    monthly_df = annual_to_monthly(rate_curve_df, rate_cols=['rate_4', 'rate_up_4', 'rate_down_4'])
    output_file = input("请输入输出Excel文件名 (例如: curve2.xlsx): ").strip() or "curve2.xlsx"
    monthly_df.to_excel(output_file)