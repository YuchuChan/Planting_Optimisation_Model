import numpy as np
import pandas as pd

# 读取数据
file_path1 = '附件2.xlsx'
y23 = pd.read_excel(file_path1, sheet_name='2023年的农作物种植情况')

for i in range(0,y23.shape[0]):
    if(y23.loc[i,'种植地块'] is np.nan):
        y23.loc[i,'种植地块'] = y23.loc[i-1,'种植地块']

output_file_path = '2023年种植数据.xlsx'
y23.to_excel(output_file_path, sheet_name = 'Sheet1', index=False, engine='openpyxl')

y23['地块编号'] = [0]*len(y23)
for i in range(1,y23.shape[0]+1):
    if(i >=2):
        if(y23.loc[i-1,'种植地块'] == y23.loc[i-2,'种植地块']):
            y23.loc[i-1,'地块编号'] = y23.loc[i-2,'地块编号']
        else:
            y23.loc[i-1,'地块编号'] = y23.loc[i-2,'地块编号']+1
    else:
        y23.loc[i-1,'地块编号'] = i

file_path2 = 'res.xlsx'
result1 = pd.read_excel(file_path2,sheet_name='1')
result2 = pd.read_excel(file_path2,sheet_name='2')

for i in range(0,y23.shape[0]):
    if(y23.loc[i,'种植季次'] == '单季' or y23.loc[i,'种植季次'] == '第一季'):
        result1.loc[y23.loc[i,'地块编号']-1, y23.loc[i,'作物名称']] = y23.loc[i,'种植面积/亩']
    else:
        result2.loc[y23.loc[i,'地块编号']-27, y23.loc[i,'作物名称']] = y23.loc[i,'种植面积/亩']

result = pd.concat([result1,result2])
output_file_path = '2023年种植方案.xlsx'
result.to_excel(output_file_path, sheet_name = 'Sheet1', index=False, engine='openpyxl')