import pandas as pd
import sys
import numpy as np

# 检查命令行参数
if len(sys.argv) != 2:
    print("请提供Excel文件名作为参数，例如：python3 check_accounts_all_auto.py 一览表.xlsx")
    sys.exit(1)

# 获取命令行参数
file_name = sys.argv[1]  # Excel文件名

# 读取Excel文件
try:
    df = pd.read_excel(file_name, sheet_name="Sheet1")
except Exception as e:
    print(f"读取Excel文件失败: {str(e)}")
    sys.exit(1)

# 从"对账部门"列提取所有部门名称并去重
# 处理可能存在的NaN值和混合类型问题
department_series = df["对账部门"].dropna()  # 删除NaN值
# 将所有值转换为字符串并过滤掉'nan'字符串
all_departments = [str(dept) for dept in department_series.unique() if str(dept) != 'nan']
# 排序
all_departments = sorted(all_departments)

print(f"共发现 {len(all_departments)} 个不同的部门")

# 生成部门对
department_pairs = []
for i in range(len(all_departments)):
    for j in range(i + 1, len(all_departments)):  # 从i+1开始，确保不会出现自己和自己的组合，也不会重复
        department_pairs.append((all_departments[i], all_departments[j]))

print(f"形成 {len(department_pairs)} 个部门对，开始进行对账检查...")

# 遍历部门对列表
for idx, dept_pair in enumerate(department_pairs, 1):
    dept1 = dept_pair[0]  # 第一个部门
    dept2 = dept_pair[1]  # 第二个部门
    
    print(f"\n[{idx}/{len(department_pairs)}] 检查 {dept1} 和 {dept2} 之间的对账情况:")
    print("-" * 60)
    
    try:
        # 提取部门1（dept1）的数据
        dept1_rows = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应收款")]
        if dept1_rows[dept2].empty or pd.isna(dept1_rows[dept2].values[0]):
            dept1_receivable = "数据不存在"
        else:
            dept1_receivable = dept1_rows[dept2].values[0]
            
        dept1_rows = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应付款")]
        if dept1_rows[dept2].empty or pd.isna(dept1_rows[dept2].values[0]):
            dept1_payable = "数据不存在"
        else:
            dept1_payable = dept1_rows[dept2].values[0]

        # 提取部门2（dept2）的数据
        dept2_rows = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应收款")]
        if dept2_rows[dept1].empty or pd.isna(dept2_rows[dept1].values[0]):
            dept2_receivable = "数据不存在"
        else:
            dept2_receivable = dept2_rows[dept1].values[0]
            
        dept2_rows = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应付款")]
        if dept2_rows[dept1].empty or pd.isna(dept2_rows[dept1].values[0]):
            dept2_payable = "数据不存在"
        else:
            dept2_payable = dept2_rows[dept1].values[0]

        # 打印结果
        print(f"{dept1} 声称应收 {dept2}: {dept1_receivable}")
        print(f"而{dept2} 声称应付 {dept1}: {dept2_payable}")
        print()
        print(f"{dept2} 声称应收 {dept1}: {dept2_receivable}")
        print(f"而{dept1} 声称应付 {dept2}: {dept1_payable}")
    except IndexError:
        print(f"无法找到 {dept1} 和 {dept2} 之间的对账数据")
    except Exception as e:
        print(f"处理 {dept1} 和 {dept2} 时发生错误: {str(e)}")

print("\n对账检查完成！")