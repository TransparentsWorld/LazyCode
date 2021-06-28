import os

dir_path: str = input("请输入文件夹所在路径：")
print(os.path.exists(dir_path))
