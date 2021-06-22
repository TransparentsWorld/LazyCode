from OfficeJob.src.OfficeUtil import *
from shutil import copyfile
import re

# 此脚本用于从指定文件夹，按文件名提取文件到指定文件夹，可用来整理凌乱的文件。

# dir = r'C:\Users\m\Desktop\表格'
# target_dir = r'C:\Users\m\Desktop\test'
# pattern:str = r'.*模板.*'

dir = input('输入需要提取的文件所在文件夹的路径：')
target_dir = input('请输入提取文件存放的文件夹路径：')
pattern: str = input('请输入用于提取的正则表达式：')

if os.path.exists(target_dir) == False:
    os.makedirs(target_dir)

files_path: list = OfficeUtil.get_file_absolute_path_in_dir(dir)

for file_path in files_path:
    absolute_path = os.path.join(target_dir, os.path.basename(file_path))  # 拼接绝对路径
    # 通过正则表达式选择文件，然后复制
    if re.search(pattern, absolute_path) != None:
        copyfile(file_path, absolute_path)
