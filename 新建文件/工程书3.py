import os
from openpyxl import load_workbook

def rename_folders(root_folder, excel_path):
    """
    将根目录下的文件夹名称修改为 Excel 表 Sheet3 中 A 列的名称（跳过第一行）。

    参数:
        root_folder (str): 根目录路径。
        excel_path (str): Excel 文件路径。
    """
    # 加载 Excel 文件并读取 A 列数据（跳过第一行）
    workbook = load_workbook(excel_path)
    sheet = workbook['Sheet3']  # 指定工作表
    folder_names = [cell.value for cell in sheet['A'][1:] if cell.value is not None]  # 从第二行开始读取 A 列的非空值

    print("从 Excel 中读取的文件夹名称:", folder_names)

    # 获取根目录下的所有文件夹
    existing_folders = [name for name in os.listdir(root_folder) if os.path.isdir(os.path.join(root_folder, name))]
    print("根目录下的现有文件夹:", existing_folders)

    # 检查文件夹数量是否匹配
    if len(existing_folders) != len(folder_names):
        print(f"警告：文件夹数量不匹配！现有文件夹数量：{len(existing_folders)}，Excel 中的文件夹名称数量：{len(folder_names)}")
        return

    # 重命名文件夹
    for old_name, new_name in zip(existing_folders, folder_names):
        old_path = os.path.join(root_folder, old_name)
        new_path = os.path.join(root_folder, new_name)

        # 检查新名称是否已存在
        if os.path.exists(new_path):
            print(f"文件夹 {new_name} 已存在，跳过重命名。")
            continue

        # 重命名文件夹
        try:
            os.rename(old_path, new_path)
            print(f"文件夹 {old_name} 已重命名为 {new_name}。")
        except Exception as e:
            print(f"重命名文件夹 {old_name} 时发生错误：{e}")

# 指定根目录和 Excel 文件路径
root_folder = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\项目工程"
excel_path = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\模板.xlsx"

rename_folders(root_folder, excel_path)