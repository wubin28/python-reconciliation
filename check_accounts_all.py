import pandas as pd
import sys

# 检查命令行参数
if len(sys.argv) != 2:
    print("请提供Excel文件名作为参数，例如：python3 check_accounts_all.py 一览表.xlsx")
    sys.exit(1)

# 获取命令行参数
file_name = sys.argv[1]  # Excel文件名

# 定义部门对列表
# 每个元素是一个包含两个部门名称的元组 (dept1, dept2)
department_pairs = [
    ("一院", "二院"),
    ("一院", "三院"),
    ("一院", "四院"),
    ("一院", "五院"),
    ("一院", "七院（含未来设计院）"),
    ("一院", "八院"),
    ("一院", "十院（含装配式建筑研究院）"),
    ("一院", "十一院"),
    ("一院", "建筑与城市院"),
    ("一院", "建筑与环艺院"),
    ("一院", "建筑与工程咨询院"),

    ("建筑与工程咨询院", "十院（含装配式建筑研究院）"),
    # 这里可以添加更多的部门对
]

# 读取Excel文件
df = pd.read_excel(file_name, sheet_name="Sheet1")

# 遍历部门对列表
for dept_pair in department_pairs:
    dept1 = dept_pair[0]  # 第一个部门
    dept2 = dept_pair[1]  # 第二个部门
    
    print(f"\n检查 {dept1} 和 {dept2} 之间的对账情况:")
    print("-" * 50)
    
    try:
        # 提取部门1（dept1）的数据
        dept1_receivable = df.loc[(df["对账部门"] == dept1) & (df["应收应付"] == "应收款"), dept2].values[0]
        dept1_payable = df.loc[(df["对账部门"] == dept1) & (df["应收应付"] == "应付款"), dept2].values[0]

        # 提取部门2（dept2）的数据
        dept2_receivable = df.loc[(df["对账部门"] == dept2) & (df["应收应付"] == "应收款"), dept1].values[0]
        dept2_payable = df.loc[(df["对账部门"] == dept2) & (df["应收应付"] == "应付款"), dept1].values[0]

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