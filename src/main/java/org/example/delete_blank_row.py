import os
import pandas as pd

def remove_blank_rows(input_file, output_file):
    # 读取Excel文件
    df = pd.read_excel(input_file)
    
    # 删除所有包含空值的行
    df_cleaned = df.dropna(how='all')
    
    # 获取文件名（不带后缀）
    file_name, file_extension = os.path.splitext(input_file)
    
    # 构建输出文件名
    output_file_name = f"{file_name}_clean{file_extension}"
    
    # 将处理后的数据保存到新的Excel文件
    df_cleaned.to_excel(output_file_name, index=False)

# 获取当前目录下的所有xlsx文件
xlsx_files = [file for file in os.listdir() if file.endswith(".xlsx")]

# 处理每个xlsx文件
for file in xlsx_files:
    # 忽略以_clean结尾的文件，以避免无限循环
    if not file.endswith("_clean.xlsx"):
        remove_blank_rows(file, file)

print("已删除空白行，并将结果保存到_clean文件中。")
