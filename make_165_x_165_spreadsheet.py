import pandas as pd
import numpy as np

# 创建165x165矩阵
data = np.arange(10000, 10000+165*165).reshape(165,165, order='F')

# 转换为DataFrame
df = pd.DataFrame(data)

# 保存为Excel
df.to_excel("165x165_Sequence.xlsx", index=False, header=False)