import pandas as pd
import sys

def truncate_to_two_decimals(num_str):
    """截取数字字符串到小数点后两位"""
    if '.' in num_str:
        integer_part, decimal_part = num_str.split('.', 1)
        decimal_truncated = decimal_part[:2].ljust(2, '0')
        return f"{integer_part}.{decimal_truncated}"
    else:
        return f"{num_str}.00"

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

total_groups = len(department_pairs) * 2
print(f"形成 {len(department_pairs)} 个部门对（共{total_groups}组数据），开始进行对账检查...")

group_idx = 1
matched_groups = []  # 存储一致的数据组

for idx, dept_pair in enumerate(department_pairs, 1):
    dept1, dept2 = dept_pair
    
    print(f"\n[{idx}/{len(department_pairs)}] 检查 {dept1} 和 {dept2} 之间的对账情况:")
    print("-" * 60)
    
    try:
        # 第一组数据检查（部门A应收 vs 部门B应付）
        dept1_receivable_row = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应收款")]
        dept1_receivable = dept1_receivable_row[dept2].values[0] if (
            not dept1_receivable_row.empty and 
            dept2 in dept1_receivable_row.columns and 
            pd.notna(dept1_receivable_row[dept2].values[0])
        ) else "数据不存在"
        
        dept2_payable_row = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应付款")]
        dept2_payable = dept2_payable_row[dept1].values[0] if (
            not dept2_payable_row.empty and 
            dept1 in dept2_payable_row.columns and 
            pd.notna(dept2_payable_row[dept1].values[0])
        ) else "数据不存在"
        
        # 打印并检查第一组
        group_output = [
            f"[{group_idx}/{total_groups}] 检查 {dept1} 和 {dept2} 之间的对账情况:",
            f"{dept1} 声称应收 {dept2}: {dept1_receivable}",
            f"而{dept2} 声称应付 {dept1}: {dept2_payable}"
        ]
        print('\n'.join(group_output) + '\n')
        
        # 检查数值一致性
        if "数据不存在" not in [str(dept1_receivable), str(dept2_payable)]:
            try:
                v1 = f"{float(dept1_receivable):.10f}"
                v2 = f"{float(dept2_payable):.10f}"
                if truncate_to_two_decimals(v1) == truncate_to_two_decimals(v2):
                    matched_groups.append('\n'.join(group_output))
            except:
                pass
        
        group_idx += 1

        # 第二组数据检查（部门B应收 vs 部门A应付）
        dept2_receivable_row = df[(df["对账部门"] == dept2) & (df["应收应付"] == "应收款")]
        dept2_receivable = dept2_receivable_row[dept1].values[0] if (
            not dept2_receivable_row.empty and 
            dept1 in dept2_receivable_row.columns and 
            pd.notna(dept2_receivable_row[dept1].values[0])
        ) else "数据不存在"
        
        dept1_payable_row = df[(df["对账部门"] == dept1) & (df["应收应付"] == "应付款")]
        dept1_payable = dept1_payable_row[dept2].values[0] if (
            not dept1_payable_row.empty and 
            dept2 in dept1_payable_row.columns and 
            pd.notna(dept1_payable_row[dept2].values[0])
        ) else "数据不存在"
        
        # 打印并检查第二组
        group_output = [
            f"[{group_idx}/{total_groups}] 检查 {dept1} 和 {dept2} 之间的对账情况:",
            f"{dept2} 声称应收 {dept1}: {dept2_receivable}",
            f"而{dept1} 声称应付 {dept2}: {dept1_payable}"
        ]
        print('\n'.join(group_output) + '\n')
        
        # 检查数值一致性
        if "数据不存在" not in [str(dept2_receivable), str(dept1_payable)]:
            try:
                v1 = f"{float(dept2_receivable):.10f}"
                v2 = f"{float(dept1_payable):.10f}"
                if truncate_to_two_decimals(v1) == truncate_to_two_decimals(v2):
                    matched_groups.append('\n'.join(group_output))
            except:
                pass
        
        group_idx += 1
        
    except Exception as e:
        print(f"处理 {dept1} 和 {dept2} 时发生错误: {str(e)}")
        group_idx += 2

# 打印最终结果
print("\n对账检查完成！")
if matched_groups:
    print("\n经过对比，发现下面的对账数据是一致的：")
    for group in matched_groups:
        print(f"\n{group}")
else:
    print("\n没有发现一致的对账数据。")