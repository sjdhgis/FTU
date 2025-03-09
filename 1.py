import shutil
import os

def delete_all_files_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            shutil.rmtree(file_path)  # 删除目录及其内容
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

folder_path = r'C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\项目工程'
delete_all_files_in_folder(folder_path)