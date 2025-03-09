import os
import shutil


def copy_template_to_subfolders(template_path, folder_path):
    """
    将Word模板文件复制到每个子文件夹中。

    参数:
        template_path (str): Word模板文件的路径。
        folder_path (str): 包含多个子文件夹的父文件夹路径。
    """
    # 获取所有子文件夹
    subfolders = [f.path for f in os.scandir(folder_path) if f.is_dir()]

    # 遍历每个子文件夹
    for subfolder in subfolders:
        # 创建Word文件的新路径
        new_doc_path = os.path.join(subfolder, os.path.basename(template_path))

        # 复制模板文件到子文件夹
        shutil.copy(template_path, new_doc_path)
        print(f"文件已复制到 {new_doc_path}")


# 示例用法
template_path = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\施工单位重要事项承诺书(紫电) .docx"  # Word模板文件路径
folder_path = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\项目工程"  # 项目工程的根文件夹路径

copy_template_to_subfolders(template_path, folder_path)