import os

def get_docx_filenames(folder_path):
    """
    提取指定文件夹中所有Word文档的文件名。

    参数:
        folder_path (str): 文件夹路径。
    """
    # 检查文件夹路径是否存在
    if not os.path.exists(folder_path):
        print("指定的文件夹路径不存在。")
        return []

    # 初始化一个列表来存储文件名
    docx_filenames = []

    # 遍历文件夹中的所有文件
    for filename in os.listdir(folder_path):
        # 检查文件扩展名是否为.docx
        if filename.endswith('.docx'):
            # 将文件名添加到列表中
            docx_filenames.append(filename)

    return docx_filenames

# 示例用法
folder_path = r"C:\Users\33\Desktop\2024年网改项目结算\网改项目\项目\文件2"  # 替换为你的文件夹路径
docx_files = get_docx_filenames(folder_path)

if docx_files:
    print("找到的Word文档文件名：")
    for file in docx_files:
        print(file)
else:
    print("没有找到Word文档。")