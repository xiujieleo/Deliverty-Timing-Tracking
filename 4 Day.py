#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd


file_path = r"F:\imile\9月15号\Dispatch waybill query.xlsx"
df = pd.read_excel(file_path, header=0)

# 2. 指定列名
signed_col = '签收时间'
arrival_col = '到件时间'
waybill_col = '运单号码'  # 运单号列名（如果你需要）

# 3. 转换为 datetime
df[signed_col] = pd.to_datetime(df[signed_col], errors='coerce')
df[arrival_col] = pd.to_datetime(df[arrival_col], errors='coerce')

# 4. 计算时间差（单位：天）
df['签收-到件(天)'] = (df[signed_col] - df[arrival_col]).dt.total_seconds() / (24 * 3600)

# 5. 筛选时间差 > 4 且为正数的记录
df_valid = df[df['签收-到件(天)'] > 4].copy()

# 6. 导出结果
output_cols = [waybill_col, arrival_col, signed_col, '签收-到件(天)']
output_path = r"F:\imile\8yue7\签收超4天.xlsx"
df_valid[output_cols].to_excel(output_path, index=False)

# 7. 输出信息
print(f"共找到 {len(df_valid)} 条签收-到件时间 > 4 天的记录")
print("样例预览：")
print(df_valid[output_cols].head())

# 5. 筛选时间差 > 4 且为正数的记录
df_valid = df[df['签收-到件(天)'] > 4].copy()

# 5.1 签收 <= 4 天的记录数量
within_4_days_df = df[(df['签收-到件(天)'] >= 0) & (df['签收-到件(天)'] <= 4)].copy()
within_4_days_count = within_4_days_df.shape[0]
print(f"共有 {within_4_days_count} 条在4天内签收的运单")

# 5.2 导出 4天内签收 的运单数据
output_path_within_4 = r"F:\imile\7月16\签收4天以内.xlsx"
within_4_days_df[output_cols].to_excel(output_path_within_4, index=False)
print("已导出签收在4天以内的运单记录到：", output_path_within_4)

# 6. 导出 >4 天的结果
output_cols = [waybill_col, arrival_col, signed_col, '签收-到件(天)']
output_path = r"F:\imile\8yue28\23\7\Dispatch waybill query20签收超4天.xlsx"
df_valid[output_cols].to_excel(output_path, index=False)

# 7. 输出信息
print(f"共找到 {len(df_valid)} 条签收-到件时间 > 4 天的记录")
print("样例预览：")
print(df_valid[output_cols].head())


# In[ ]:





# In[ ]:




