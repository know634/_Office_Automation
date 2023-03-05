# -- coding: utf-8 --
import time
start_time = time.time()
import pandas as pd
from openpyxl import load_workbook

# 读取 Excel 表格
wb = load_workbook(filename='1随机生成.xlsx')
ws = wb['Sheet1']

# 使用 pandas 将 Excel 表格转换成 DataFrame
df = pd.DataFrame(ws.values)

# 将第二列“学号”转换成数字类型
df[1][1:] = df[1][1:].astype(int)

# 将第三列“成绩”转换成文本类型
df[2] = df[2].astype(str)

# 按照第三列“成绩”进行降序排序
df_sorted = df.sort_values(2, ascending=False)

# 将排序后的数据写入 Excel 表格
df_sorted.to_excel('2排序表.xlsx', index=True, startrow=-1, startcol=-1)

end_time = time.time()
total_time = end_time - start_time
print("\033[31m程序运行时间为：{:.2f}秒\033[0m".format(total_time))
