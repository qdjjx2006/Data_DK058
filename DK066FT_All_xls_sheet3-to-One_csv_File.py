# -*- coding: utf-8 -*-
"""
Created on Mon Aug 18 17:07:46 2025

@author: pc
"""

import pandas as pd
import glob
import os

# 1. 获取所有xls文件
folder_path = r'c:\MiddleTest\DK066\FT\CN'  # 替换为你的文件夹路径
all_files = glob.glob(os.path.join(folder_path, "*.xls"))

# 2. 创建空列表存储数据
all_data = []

df = pd.read_excel(
    all_files[0],
    sheet_name='DUT_DATA',
    #skiprows=range(4),  # 跳过前5行表头
    nrows=3
    )
all_data.append(df)

# 3. 遍历处理每个文件
for file in all_files:
    try:
        # 读取DUT_DATA工作表，跳过前5行（0-4）
        df = pd.read_excel(
            file,
            sheet_name='DUT_DATA',
            skiprows=[1,2,3,4],  # 跳过前5行表头
            )
        
        # 添加文件名列（可选）
        #df['source_file'] = os.path.basename(file)
        
        # 添加到数据列表
        all_data.append(df)
        
        print(f"成功处理: {os.path.basename(file)}")
    
    except Exception as e:
        print(f"处理失败 [{os.path.basename(file)}]: {str(e)}")

# 4. 合并所有数据
if all_data:
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # 5. 保存为CSV
    output_path = os.path.join(folder_path, "combined_data.csv")
    combined_df.to_csv(output_path, index=False)
    
    print(f"\n合并完成! 共处理 {len(all_data)} 个文件")
    print(f"输出文件: {output_path}")
    print(f"总记录数: {len(combined_df)}")
else:
    print("未找到有效数据")
