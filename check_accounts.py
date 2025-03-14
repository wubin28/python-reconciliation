import pandas as pd
import sys

# 检查命令行参数
if len(sys.argv) != 4:
    print("请提供Excel文件名和两个部门名称作为参数，例如：python3 check_accounts.py file_name.xlsx 十院（含装配式建筑研究院） 建筑与工程咨询院")
    sys.exit(1)

# 获取命令行参数
file_name = sys.argv[1]  # Excel文件名
dept1 = sys.argv[2]      # 第一个部门
dept2 = sys.argv[3]      # 第二个部门

# 读取Excel文件
df = pd.read_excel(file_name, sheet_name="Sheet1")

# 提取部门1（dept1）的数据
dept1_receivable = df.loc[(df["对账部门"] == dept1) & (df["应收应付"] == "应收款"), dept2].values[0]
dept1_payable = df.loc[(df["对账部门"] == dept1) & (df["应收应付"] == "应付款"), dept2].values[0]

# 提取部门2（dept2）的数据
dept2_receivable = df.loc[(df["对账部门"] == dept2) & (df["应收应付"] == "应收款"), dept1].values[0]
dept2_payable = df.loc[(df["对账部门"] == dept2) & (df["应收应付"] == "应付款"), dept1].values[0]

# 打印结果
print(f"{dept1} 声称应收 {dept2}: {dept1_receivable}")
print(f"而{dept2} 声称应付 {dept1}: {dept2_payable}")

print(f"{dept2} 声称应收 {dept1}: {dept2_receivable}")
print(f"而{dept1} 声称应付 {dept2}: {dept1_payable}")
