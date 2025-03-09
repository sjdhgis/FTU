import os
import shutil

# 源文件路径
source_file = 'path/to/your/source/file.txt'  # 替换为你的源文件路径

# 目标文件夹路径
destination_folder = 'path/to/your/destination/folder'  # 替换为目标文件夹路径

# 确保目标文件夹存在
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# 复制文件51次
for i in range(1, 52):  # 从1到51，总共51份
    # 构造目标文件路径，添加编号
    destination_file = os.path.join(destination_folder, f'file_copy_{i}.docx')  # 生成副本文件名
    # 复制文件
    shutil.copy2(source_file, destination_file)  # 使用copy2保留元数据

print(f"文件已成功复制到 {destination_folder}，共51份。")