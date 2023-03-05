# -- coding: utf-8 --
import time
start_time = time.time()
import pandas as pd

# 读取Excel文件
df = pd.read_excel("2排序表.xlsx", sheet_name="Sheet1")

# 添加“奖项”列
df.insert(3, "奖项", "")  # 在第4列插入空白列“奖项”

# 标注奖项
df.loc[df["成绩"] >= 80, "奖项"] = "一等奖"
df.loc[(df["成绩"] >= 70) & (df["成绩"] < 80), "奖项"] = "二等奖"
df.loc[df["成绩"] < 70, "奖项"] = "三等奖"

# 写入Excel文件
df.to_excel("3完整表.xlsx", index=False)

end_time = time.time()
total_time = end_time - start_time
print("\033[31m程序运行时间为：{:.2f}秒\033[0m".format(total_time))

