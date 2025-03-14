import pandas as pd
import sys

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

# 保持Excel表中原始顺序获取部门名称
all_departments_with_order = []
seen = set()  # 用于跟踪已经见过的部门，避免重复

# 遍历DataFrame获取部门，保持原始顺序
for dept in df["对账部门"]:
    if pd.notna(dept) and str(dept) != 'nan':  # 确保部门名称有效
        dept_str = str(dept)  # 转换为字符串
        if dept_str not in seen:  # 只添加未见过的部门
            seen.add(dept_str)
            all_departments_with_order.append(dept_str)

print(f"按Excel表顺序，共发现 {len(all_departments_with_order)} 个不同的部门")

# 生成部门对，保持原始顺序
department_pairs = []
for i in range(len(all_departments_with_order)):
    for j in range(i + 1, len(all_departments_with_order)):
        department_pairs.append((all_departments_with_order[i], all_departments_with_order[j]))

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
        if len(dept1_rows) == 0 or dept2 not in dept1_rows.columns or pd.isna(dept1_rows[dept2].values[0]):
            dept1_receivable = "数据不存在"
        else:
            dept1_receivable = dept1_rows[dept2].values[0]
            
        dept1_rows = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应付款")]
        if len(dept1_rows) == 0 or dept2 not in dept1_rows.columns or pd.isna(dept1_rows[dept2].values[0]):
            dept1_payable = "数据不存在"
        else:
            dept1_payable = dept1_rows[dept2].values[0]

        # 提取部门2（dept2）的数据
        dept2_rows = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应收款")]
        if len(dept2_rows) == 0 or dept1 not in dept2_rows.columns or pd.isna(dept2_rows[dept1].values[0]):
            dept2_receivable = "数据不存在"
        else:
            dept2_receivable = dept2_rows[dept1].values[0]
            
        dept2_rows = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应付款")]
        if len(dept2_rows) == 0 or dept1 not in dept2_rows.columns or pd.isna(dept2_rows[dept1].values[0]):
            dept2_payable = "数据不存在"
        else:
            dept2_payable = dept2_rows[dept1].values[0]

        # 打印结果
        print(f"{dept1} 声称应收 {dept2}: {dept1_receivable}")
        print(f"而{dept2} 声称应付 {dept1}: {dept2_payable}")
        print()
        print(f"{dept2} 声称应收 {dept1}: {dept2_receivable}")
        print(f"而{dept1} 声称应付 {dept2}: {dept1_payable}")
    except Exception as e:
        print(f"处理 {dept1} 和 {dept2} 时发生错误: {str(e)}")

print("\n对账检查完成！")