import pandas as pd
import numpy as np

# 生成随机矩阵 - 将类型修改为浮点数而不是整数
np.random.seed(42)
data = np.random.randint(10000, 500001, size=(165, 165)).astype(float)  # 将整数数组转换为浮点数

# 设置特定区域为空白（转换为NaN）
# 1. A1-A5（列索引0，行索引0-4）
data[0:5, 0] = np.nan

# 2. B6-B9（列索引1，行索引5-8）
data[5:9, 1] = np.nan

# 3. 对角线右下方的连续空白区域
for col_idx in range(1, 41):  # 处理B列到AN列（列索引1-40）
    start_row = 5 + 4 * (col_idx - 1)  # 计算起始行索引
    end_row = start_row + 4
    data[start_row:end_row, col_idx] = np.nan  # 设置当前列连续4行空白

# 转换为DataFrame并保存
df = pd.DataFrame(data)
df.to_excel("special_blank_matrix.xlsx", index=False, header=False)

# 验证空白区域位置
print("A1-A5空白验证:", df.iloc[0:5, 0].isna().all())      # 应返回True
print("B6-B9空白验证:", df.iloc[5:9, 1].isna().all())      # 应返回True
print("C10-C13空白验证:", df.iloc[9:13, 2].isna().all())   # 应返回True
print("最后一组空白验证:", df.iloc[161:165, 40].isna().all()) # 应返回True