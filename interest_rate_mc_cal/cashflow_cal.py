import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
from tqdm import tqdm

'''
根据输入的债券基础信息计算现金流
参数：
1.file_path：包含债券基础信息的表格，默认直接从持仓明细获得的数据无需处理
    'principal'：本金
    'issue_date'：起息日，一般为YYYY-MM-DD或excel序列值（如44000）
    'maturity_date'：到期日，一般为YYYY-MM-DD或excel序列值（如44000）
    'coupon_rate':百分号前的数据，4
    'payment_freq'：年付/半年付/季付/月付/一次性还本付息/到期支付，通过freq_map自动处理，可修改freq_map增加枚举值
    其他字段，例如账户，债券代码等，通过metadata_fields配置，5个字段，如需增加，调整mc_cal的cashflow_start_col参数
2.start_date：评估日，建议输入YYYYMMDD格式
3.months：评估时间长度，默认601个月

返回：
1.result_df:总现金流
2.principal_result_df：本金现金流
3.coupon_result_df：票息现金流
4.output_file：输出3个sheet的excel表，分别为三个现金流df，默认cashflow_analysis.xlsx
'''

def parse_date(date_str):
    """
    将日期字符串/Excel 序列值转换为 datetime 对象（支持多种格式）
    - 字符串格式：YYYYMMDD、YYYY-MM-DD、YYYY/MM/DD 等
    - Excel 序列值：数值型（如 44000 对应 2020-01-01）
    """
    if pd.isna(date_str):  # 处理空值
        return pd.NaT
    
    # 处理 Excel 日期序列值（数值类型）
    if isinstance(date_str, (int, float)):
        try:
            # Excel 通常从 1900-01-01 开始（序列值 1），但需处理 1900 年非闰年的 bug
            # 参考：https://support.microsoft.com/en-us/help/214330
            return pd.to_datetime(date_str, origin='1899-12-30', unit='D')
        except:
            pass  # 若转换失败，尝试按字符串处理
    
    # 处理日期字符串
    if isinstance(date_str, str):
        date_str = date_str.strip()
        # 支持的格式：YYYYMMDD、YYYY-MM-DD、MM/DD/YYYY 等
        for fmt in ['%Y%m%d', '%Y-%m-%d', '%m/%d/%Y', '%Y/%m/%d']:
            try:
                return pd.to_datetime(date_str, format=fmt)
            except:
                pass
        # 尝试自动解析（处理模糊格式）
        try:
            return pd.to_datetime(date_str, errors='coerce')  # 容错处理
        except:
            pass
    
    # 处理其他类型（如 pd.Timestamp）
    try:
        return pd.to_datetime(date_str)
    except:
        raise ValueError(f"无法解析的日期格式：{date_str}")

def get_last_day_of_month(date):
    """获取月份的最后一天"""
    next_month = date.replace(day=28) + relativedelta(days=4)
    return next_month - relativedelta(days=next_month.day)

def is_same_month(date1, date2):
    """判断两个日期是否在同一个月"""
    return date1.year == date2.year and date1.month == date2.month

def generate_cashflows(file_path, start_date_str, months=360):
    """从Excel读取债券数据，生成现金流表（修复数组广播错误）"""
    try:
        # 读取Excel文件
        bond_data = pd.read_excel(file_path)
        
        # 检查必要的列是否存在
        required_columns = ['principal', 'issue_date', 'maturity_date', 'coupon_rate', 'payment_freq']
        missing_columns = [col for col in required_columns if col not in bond_data.columns]
        
        if missing_columns:
            raise ValueError(f"Excel文件缺少必要的列: {', '.join(missing_columns)}")
        
        freq_map = {
            '年付': 1,    # 每年支付1次
            '半年付': 2,   # 每半年支付1次
            '季付': 4,     # 每季度支付1次
            '月付': 12,    # 每月支付1次  
            '一次性还本付息': 99,  # 到期一次性支付（期末）
            '到期支付': 98,     # 到期支付
            # 可根据需要添加更多映射
        }
        
        # 将付款频率映射为数值
        bond_data['payment_freq'] = bond_data['payment_freq'].map(freq_map).fillna(99)
        
        # 转换日期列
        bond_data['issue_date'] = bond_data['issue_date'].apply(parse_date)
        bond_data['maturity_date'] = bond_data['maturity_date'].apply(parse_date)
        bond_data['maturity_date_end'] = bond_data['maturity_date'].apply(get_last_day_of_month)
        
        # 计算持有年数
        bond_data['years_held'] = (bond_data['maturity_date'] - bond_data['issue_date']).dt.days / 365
        
        start_date = parse_date(start_date_str)
        
        # 生成未来N个月的月末日期
        date_list = [get_last_day_of_month(start_date + relativedelta(months=i)) 
                    for i in range(months)]
        date_indices = {date.strftime('%Y%m'): i for i, date in enumerate(date_list)}
        
        # 创建结果数组，初始化为0
        num_bonds = len(bond_data)
        result = np.zeros((num_bonds, months), dtype=float)  # 本金+利息
        principal_result = np.zeros((num_bonds, months), dtype=float)  # 本金+到期一次还本付息
        coupon_result = np.zeros((num_bonds, months), dtype=float)  # 票息
        
        # 创建日期矩阵（每个债券对应所有计算日期）
        date_matrix = np.array(date_list, dtype='datetime64[D]')
        date_matrix = np.tile(date_matrix, (num_bonds, 1))
        
        # 创建债券属性矩阵（确保所有矩阵形状一致）
        principal_matrix = np.tile(bond_data['principal'].values.reshape(-1, 1), (1, months))
        coupon_rate_matrix = np.tile((bond_data['coupon_rate']/100).values.reshape(-1, 1), (1, months))
        payment_freq_matrix = np.tile(bond_data['payment_freq'].values.reshape(-1, 1), (1, months))
        maturity_date_matrix = np.tile(bond_data['maturity_date'].values.reshape(-1, 1), (1, months))
        maturity_date_end_matrix = np.tile(bond_data['maturity_date_end'].values.reshape(-1, 1), (1, months))
        years_held_matrix = np.tile(bond_data['years_held'].values.reshape(-1, 1), (1, months))
        product_type_matrix = np.tile(bond_data.get('product_type', '').values.reshape(-1, 1), (1, months))
        
        # 向量化计算月份差异
        month_diff_matrix = np.zeros((num_bonds, months), dtype=int)
        
        # 将NumPy datetime64转换为年、月数组
        maturity_years = maturity_date_matrix.astype('datetime64[Y]').astype(int) + 1970
        maturity_months = maturity_date_matrix.astype('datetime64[M]').astype(int) % 12 + 1
        
        for i in range(months):
            current_date = date_list[i]
            current_year = current_date.year
            current_month = current_date.month
            
            # 计算月份差异
            month_diff = (current_year - maturity_years[:, i]) * 12 + (current_month - maturity_months[:, i])
            month_diff_matrix[:, i] = month_diff
        
        month_diff_abs_matrix = np.abs(month_diff_matrix)
        
        # 向量化计算每个月的现金流
        for j in tqdm(range(months), desc="计算现金流", unit="月", ncols=80):
            current_date = date_list[j]

            # 创建掩码矩阵
            is_maturity_month = np.vectorize(lambda x: is_same_month(current_date, pd.Timestamp(x)))(maturity_date_matrix[:, j])
            is_before_maturity = date_matrix[:, j] <= maturity_date_end_matrix[:, j]
            is_not_perpetual = product_type_matrix[:, j] != "优先股"
            is_freq_98 = payment_freq_matrix[:, j] == 98
            is_freq_99 = payment_freq_matrix[:, j] == 99
            freq_not_zero = payment_freq_matrix[:, j] != 0

            # 计算第一部分：基础本金支付（确保freq=98只在到期月支付本金）
            part1_mask = is_not_perpetual & ((is_freq_98 & is_maturity_month) | (~is_freq_98 & is_maturity_month))
            part1 = np.zeros(num_bonds)
            part1[part1_mask] = principal_matrix[part1_mask, j]

            # 计算第二部分：到期利息支付（排除freq=98的债券）
            part2_mask = is_not_perpetual & is_maturity_month & is_freq_99
            part2 = np.zeros(num_bonds)
            part2[part2_mask] = principal_matrix[part2_mask, j] * \
                              coupon_rate_matrix[part2_mask, j] * \
                              years_held_matrix[part2_mask, j]

            # 计算第三部分：常规利息支付（排除freq=98的债券）
            valid_regular_interest = is_not_perpetual & is_before_maturity & ~(is_freq_98 | is_freq_99)
            part3_condition = (month_diff_abs_matrix[:, j] % (12 / payment_freq_matrix[:, j])) == 0
            part3_mask = valid_regular_interest & freq_not_zero & part3_condition

            part3 = np.zeros(num_bonds)
            part3[part3_mask] = principal_matrix[part3_mask, j] * \
                              coupon_rate_matrix[part3_mask, j] / \
                              payment_freq_matrix[part3_mask, j]

            # 累加三部分现金流
            result[:, j] =part1 + part2 + part3
            principal_result[:, j] = part1
            coupon_result[:, j] = part2 + part3
        
        # 转换为DataFrame
        columns = [date.strftime('%Y%m') for date in date_list]
        result_df = pd.DataFrame(result, columns=columns)
        principal_result_df = pd.DataFrame(principal_result, columns=columns)
        coupon_result_df = pd.DataFrame(coupon_result, columns=columns)
        
        # 添加债券ID、名称、账户
        metadata_fields = [
            ('account_1', lambda: bond_data['account_1'] if 'account_1' in bond_data.columns else ''),
            ('account_2', lambda: bond_data['account_2'] if 'account_2' in bond_data.columns else ''),
            ('product_type', lambda: bond_data['product_type'] if 'product_type' in bond_data.columns else ''),
            ('bond_code', lambda: bond_data['bond_code'] if 'bond_code' in bond_data.columns else ''),
            ('bond_name', lambda: bond_data['bond_name'] if 'bond_name' in bond_data.columns else '')
        ]
        
        for df in [result_df, principal_result_df, coupon_result_df]:
            for idx, (field_name, value_getter) in enumerate(metadata_fields):
                df.insert(idx, field_name, value_getter())
        
        return result_df, principal_result_df, coupon_result_df
    
    except Exception as e:
        print(f"处理Excel文件时出错: {str(e)}")
        return None, None, None

def main():
    print("===== 债券现金流计算器（向量化计算） =====")
    
    # 获取输入文件路径
    file_path = input("请输入债券数据Excel文件路径 (例如: bonds.xlsx): ").strip()
    
    # 检查文件路径是否提供
    if not file_path:
        print("错误: 未提供文件路径")
        sys.exit(1)
    
    # 获取开始日期
    start_date = input("请输入开始日期 (格式: YYYYMMDD, 例如: 20250101): ").strip()
    
    if not start_date:
        print("错误: 未提供开始日期")
        sys.exit(1)
    
    # 生成现金流表
    print("正在计算现金流...")
    result_df, principal_result_df, coupon_result_df = generate_cashflows(file_path, start_date, months=601)
    
    if principal_result_df is not None and coupon_result_df is not None:
        # 获取输出Excel文件名
        output_file = input("请输入输出Excel文件名 (例如: cashflow_analysis.xlsx): ").strip() or "cashflow_analysis.xlsx"
        
        # 确保文件扩展名为.xlsx
        if not output_file.lower().endswith('.xlsx'):
            output_file += '.xlsx'
        
        # 创建ExcelWriter对象
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 将三个表分别写入不同的sheet
            principal_result_df.to_excel(writer, sheet_name='本金', index=False)
            coupon_result_df.to_excel(writer, sheet_name='利息', index=False)
            result_df.to_excel(writer, sheet_name='总现金流', index=False)
        
        print(f"\n现金流表已生成并保存到: {output_file}")
        print(f"包含三个sheet:")
        print(f"1. 本金")
        print(f"2. 利息")
        print(f"3. 总现金流")
        print(f"\n共处理 {len(result_df)} 只债券，覆盖 {result_df.shape[1]-5} 个月")
    else:
        print("生成现金流表失败")

if __name__ == "__main__":
    main()