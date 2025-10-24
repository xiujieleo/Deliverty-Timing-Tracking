#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[1]:


import pandas as pd
from datetime import datetime, timedelta

# 1. 读取Excel文件（请替换为实际路径）
file_path = r"F:\imile\6月17\Dispatch waybill query催拍.xlsx"
try:
    # 跳过前3行，第4行作为列头
    df = pd.read_excel(file_path, header=3, skiprows=[4,5])
except Exception as e:
    print(f"文件读取失败: {str(e)}")
    exit()

# 2. 显示所有列名供确认
print("所有列名：")
print(df.columns.tolist())

# 3. 手动指定关键列（根据您看到的列名）
time_col = '2025-06-16 18:24:56'  # 这是实际的时间数据列名
waybill_col = '6061125935561'      # 这是运单号列名

# 4. 验证列是否存在
if time_col not in df.columns or waybill_col not in df.columns:
    print("错误：指定的列不存在")
    print("请从以下列中选择时间列和运单号列：")
    print(df.columns.tolist())
    exit()

# 5. 转换时间列为datetime格式
df['派件时间'] = pd.to_datetime(df[time_col], errors='coerce')

# 6. 计算时间差（当前时间 - 派件时间）
current_time = datetime.now()
df['时间差(小时)'] = (current_time - df['派件时间']).dt.total_seconds() / 3600

# 7. 筛选超过24小时的记录
overdue_shipments = df[df['时间差(小时)'] > 24].copy()

# 8. 保存结果
output_cols = [waybill_col, '派件时间', '时间差(小时)']
output_path = r"F:\imile\6月17\超24小时运单.xlsx"
overdue_shipments[output_cols].to_excel(output_path, index=False)

print(f"已筛选出 {len(overdue_shipments)} 条超24小时运单")
print("结果已保存到:", output_path)
print("\n样例数据：")
print(overdue_shipments[output_cols].head())


# In[ ]:





# In[ ]:




