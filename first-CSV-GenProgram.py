# -*- coding: utf-8 -*-
"""
Created on Fri May  9 16:25:19 2025

@author: pc
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. 读取CSV文件
file_path = r"C:\Users\pc\python_project\deepseek_project\DK051C_5-6.csv"  # 修改为你的文件路径
df = pd.read_csv(file_path, index_col="Chip_ID")  # 假设第一列是芯片编号

# 2. 查看数据概览
print("数据前5行：")
print(df.head())
print("\n数据统计描述：")
print(df.describe())
print("\n缺失值统计：")
print(df.isnull().sum())

# 3. 数据清洗（示例）
# 删除包含缺失值的行
df_clean = df.dropna()
# 重置索引
df_clean.reset_index(inplace=True)

# 4. 数据可视化
plt.figure(figsize=(15, 10))

# 4.1 折线图 - 各芯片参数趋势
plt.subplot(2, 2, 1)
for column in df_clean.columns[1:3]:  # 跳过芯片ID列
    plt.plot(df_clean["Chip_ID"], df_clean[column], marker='o', label=column)
plt.title("芯片参数趋势图")
plt.xlabel("芯片编号")
plt.ylabel("参数值")
plt.xticks(rotation=45)
plt.legend()


plt.subplot(2, 2, 4)
for column in df_clean.columns[-2:]:  # 跳过芯片ID列
    plt.plot(df_clean["Chip_ID"], df_clean[column], marker='o', label=column)
plt.title("芯片参数趋势图")
plt.xlabel("芯片编号")
plt.ylabel("参数值")
plt.xticks(rotation=45)
plt.legend()

# 4.2 箱线图 - 参数分布
plt.subplot(2, 2, 2)
sns.boxplot(data=df_clean.drop(["Chip_ID","IVCC12V","IVCC29V5"], axis=1))
plt.title("参数分布箱线图")
plt.xticks(rotation=45)

# 4.3 热力图 - 参数相关性
plt.subplot(2, 2, 3)
corr_matrix = df_clean.drop("Chip_ID", axis=1).corr()
sns.heatmap(corr_matrix, annot=True, cmap="coolwarm")
plt.title("参数相关性热力图")

# 4.4 散点图矩阵
#sns.pairplot(df_clean.drop("Chip_ID", axis=1))  # 可能需要调整
sns.pairplot(df_clean[["IVCC12V","IVCC29V5"]])
plt.suptitle("参数散点图矩阵")

plt.tight_layout()
plt.show()

# 5. 保存处理后的数据
df_clean.to_csv("processed_chip_data.csv", index=False)