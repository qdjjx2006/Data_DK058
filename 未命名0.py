# -*- coding: utf-8 -*-
"""
Created on Tue May 27 11:19:25 2025

@author: pc
"""

import pandas as pd
import glob

def StrToValue(x):
    a = x.replace('=', '', -1)
    b = a.replace('"', '', -1)
    
    """try:
        c = float(b)
    except ValueError:
        c = float('nan')
        
    return c"""

    d = b.replace('.', '', -1)
    if d.isdigit():
        c = float(b)
    else:
        c = float('nan')
        
    return c
    
# 获取当前文件夹下所有CSV文件
file_list = glob.glob(r'../data_065/*.csv')

"""
with open(r'file_list_name_backup.txt', 'r+', encoding="utf-8") as f:
    try:
        read_data = f.read()
    except:
        print("File is empty")
    for str_file in file_list:
        f.write( str_file )
        f.write('\n')
print(f.closed)
        
for str_file in file_list:
    os.rename(str(str_file),str(str_file[:20]+'.csv'))
"""





# 初始化一个空的DataFrame来存储结果
combined_df = pd.DataFrame()    

for file in file_list:
    # 读取CSV文件
    df = pd.read_csv(file)
    
    # 提取Q列数据（从第8行开始）
    #q_data = df['7:VFBmax'].iloc[7:].reset_index(drop=True).to_frame(name='Q')
    q_data = df.iloc[7:,16].reset_index(drop=True)#.to_frame(name='Vfb')
    #q_data = q_data.astype(float)
    #q_data = q_data.apply(lambda x: x.replace('=','',-1))
    #q_data = q_data.apply(lambda x: x.replace('"','',-1))
    #q_data = q_data.apply(lambda x: float(x) if x.replace('=', '', 1).isdigit() else float('nan'))
    q_data = q_data.apply(StrToValue).to_frame(name='Vfb')
    
    # 将数据添加到总DataFrame    
    combined_df = pd.concat([combined_df, q_data], ignore_index=True)

labels = pd.cut(combined_df.Vfb,[.6,1.4,2.2,3,3.8,4.6,4.7,4.8,4.9,5,5.1,5.2,5.3,5.4,5.7,6.2,7])
grouped = combined_df.groupby([labels],observed=False)
TJ = grouped.size().to_frame(name='Vfb统计')

#combined_df.to_excel('Vfb_data.xlsx', index=True, sheet_name="Sheet1")

def style_negative(v, props=''):
    return props if v < 0 else None
with pd.ExcelWriter(
    "Vfb_data.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace") as writer:
    combined_df.style.\
        map(style_negative, props='color:red;').\
        highlight_max(axis=0).\
        to_excel(writer, index=True, sheet_name="Sheet1", engine='openpyxl')
    TJ.style.\
        map(lambda v: 'color:blue;' if (v < 1e6) and (v > 1e3) else None).\
        highlight_max(axis=0).\
        to_excel(writer, index=True, sheet_name="Sheet2", engine='openpyxl')

print("数据处理完成，结果已保存到 Vfb_data.xlsx")
#df['A'] = df['A'].apply(lambda x: float(x) if x.replace('.', '', 1).isdigit() else float('nan'))
"""
def make_pretty(styler):
    styler.set_caption("Weather Conditions")
    styler.format(rain_condition)
    styler.format_index(lambda v: v.strftime("%A"))
    styler.background_gradient(axis=None, vmin=1, vmax=5, cmap="YlGnBu")
    return styler
"""