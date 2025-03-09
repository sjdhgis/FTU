import os
import pandas as pd

# 指定Excel文件路径
excel_file = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\模板.xlsx"  # 替换为你的Excel文件名
sheet_name = "Sheet3"  # 替换为你的工作表名称（如果需要）

# 指定目标文件夹路径
target_folder = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\项目工程"  # 替换为你想要的路径

# 确保目标文件夹存在
if not os.path.exists(target_folder):
    os.makedirs(target_folder)
    print(f"目标文件夹 '{target_folder}' 已创建。")

# 读取Excel文件的A列内容
df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols="A")  # 只读取A列

# 获取A列的所有值
folder_names = df.iloc[:, 0].dropna().tolist()  # 去除空值

# 在目标文件夹下创建子文件夹
for folder_name in folder_names:
    folder_name = str(folder_name).strip()  # 去除可能的空格
    full_path = os.path.join(target_folder, folder_name)  # 拼接完整路径

    if not os.path.exists(full_path):  # 检查文件夹是否存在
        os.makedirs(full_path)
        print(f"文件夹 '{full_path}' 已创建。")
    else:
        print(f"文件夹 '{full_path}' 已存在。")





#template_path = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\施工单位重要事项承诺书（紫电）.doc"  # Word模板文件路径
#folder_path = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\项目工程"  # 项目工程的根文件夹路径
#excel_file = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\模板.xlsx"  # Excel文件路径