import pandas as pd
import sys

if len(sys.argv) != 2:
    print("请提供Excel文件名作为参数，例如：python3 check_accounts_all_auto.py 一览表.xlsx")
    sys.exit(1)

file_name = sys.argv[1]

try:
    df = pd.read_excel(file_name, sheet_name="Sheet1")
except Exception as e:
    print(f"读取Excel文件失败: {str(e)}")
    sys.exit(1)

all_departments_with_order = []
seen = set()

for dept in df["对账部门"]:
    if pd.notna(dept) and str(dept) != 'nan':
        dept_str = str(dept)
        if dept_str not in seen:
            seen.add(dept_str)
            all_departments_with_order.append(dept_str)

print(f"按Excel表顺序，共发现 {len(all_departments_with_order)} 个不同的部门")

department_pairs = []
for i in range(len(all_departments_with_order)):
    for j in range(i + 1, len(all_departments_with_order)):
        department_pairs.append((all_departments_with_order[i], all_departments_with_order[j]))

total_groups = len(department_pairs) * 2  # 新增总组数统计
print(f"形成 {len(department_pairs)} 个部门对（共{total_groups}组数据），开始进行对账检查...")

group_idx = 1  # 新增全局计数器

for idx, dept_pair in enumerate(department_pairs, 1):
    dept1, dept2 = dept_pair
    
    print(f"\n[{idx}/{len(department_pairs)}] 检查 {dept1} 和 {dept2} 之间的对账情况:")
    print("-" * 60)
    
    try:
        # 部门1应收数据
        dept1_receivable_row = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应收款")]
        dept1_receivable = dept1_receivable_row[dept2].values[0] if (
            not dept1_receivable_row.empty and 
            dept2 in dept1_receivable_row.columns and 
            pd.notna(dept1_receivable_row[dept2].values[0])
        ) else "数据不存在"
        
        # 部门2应付数据
        dept2_payable_row = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应付款")]
        dept2_payable = dept2_payable_row[dept1].values[0] if (
            not dept2_payable_row.empty and 
            dept1 in dept2_payable_row.columns and 
            pd.notna(dept2_payable_row[dept1].values[0])
        ) else "数据不存在"
        
        # 第一组检查
        print(f"[{group_idx}/{total_groups}] 检查 {dept1} 和 {dept2} 之间的对账情况:")
        print(f"{dept1} 声称应收 {dept2}: {dept1_receivable}")
        print(f"而{dept2} 声称应付 {dept1}: {dept2_payable}\n")
        group_idx += 1
        
        # 部门2应收数据
        dept2_receivable_row = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应收款")]
        dept2_receivable = dept2_receivable_row[dept1].values[0] if (
            not dept2_receivable_row.empty and 
            dept1 in dept2_receivable_row.columns and 
            pd.notna(dept2_receivable_row[dept1].values[0])
        ) else "数据不存在"
        
        # 部门1应付数据
        dept1_payable_row = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应付款")]
        dept1_payable = dept1_payable_row[dept2].values[0] if (
            not dept1_payable_row.empty and 
            dept2 in dept1_payable_row.columns and 
            pd.notna(dept1_payable_row[dept2].values[0])
        ) else "数据不存在"
        
        # 第二组检查
        print(f"[{group_idx}/{total_groups}] 检查 {dept1} 和 {dept2} 之间的对账情况:")
        print(f"{dept2} 声称应收 {dept1}: {dept2_receivable}")
        print(f"而{dept1} 声称应付 {dept2}: {dept1_payable}\n")
        group_idx += 1
        
    except Exception as e:
        print(f"处理 {dept1} 和 {dept2} 时发生错误: {str(e)}")
        group_idx += 2  # 保持计数器同步

print("\n对账检查完成！")