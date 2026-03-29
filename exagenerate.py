import pandas as pd
import os

def generate_dictionary_files():
    print("正在准备中英混合商业词典数据...")
    
    # 准备字典数据 (贴合 DataCopilot 的测试场景)
    data = {
        "English_Term": ["Artificial Intelligence", "Revenue", "Net Profit", "Gross Margin", 
                         "Department", "Employee ID", "Year over Year (YoY)", "Month over Month (MoM)", 
                         "Dashboard", "Database", "Machine Learning", "Supply Chain", 
                         "Customer Acquisition Cost", "Return on Investment (ROI)"],
        "Chinese_Term": ["人工智能", "营业收入", "净利润", "毛利率", 
                         "部门", "员工工号", "同比", "环比", 
                         "数据看板", "数据库", "机器学习", "供应链", 
                         "获客成本", "投资回报率"],
        "Category": ["Technology", "Finance", "Finance", "Finance", 
                     "HR", "HR", "Analysis", "Analysis", 
                     "Data", "Technology", "Technology", "Operations", 
                     "Marketing", "Finance"],
        "Priority": ["High", "High", "High", "Medium", 
                     "High", "High", "Medium", "Medium", 
                     "Low", "Medium", "High", "Medium", 
                     "High", "High"]
    }

    # 转化为 Pandas DataFrame
    df = pd.DataFrame(data)

    print("\n开始生成文件...")

    # 1. 生成 CSV 文件 (带有 utf-8-sig BOM，防止 Excel 打开中文乱码)
    csv_name = "IT_Business_Dict.csv"
    df.to_csv(csv_name, index=False, encoding="utf-8-sig")
    print(f"✅ 成功生成 CSV 格式: {csv_name}")

    # 2. 生成标准的 XLSX 文件
    xlsx_name = "IT_Business_Dict.xlsx"
    df.to_excel(xlsx_name, index=False, engine="openpyxl")
    print(f"✅ 成功生成 XLSX 格式: {xlsx_name}")

    # 3. 生成 XLSM (启用宏的 Excel 工作簿) 格式
    # 注：Pandas 无法直接编写 VBA 代码，但可以利用 openpyxl 保存为合法的 .xlsm 容器后缀。
    # 这样用户打开后可以直接在里面录制宏并正常保存。
    xlsm_name = "IT_Business_Dict.xlsm"
    df.to_excel(xlsm_name, index=False, engine="openpyxl")
    print(f"✅ 成功生成 XLSM 格式: {xlsm_name}")

    print("\n🎉 所有格式的词典文件已生成完毕，请在当前文件夹查看！")

if __name__ == "__main__":
    generate_dictionary_files()