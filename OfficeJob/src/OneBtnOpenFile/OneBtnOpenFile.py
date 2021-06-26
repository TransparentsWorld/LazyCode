import os
import configparser

# 从配置文件读取批处理命令，并执行实现一键打开多个常见文件、文件夹、程序和网页的功能

config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")

print('配置文件中一键启动组如下所示：')
start_set = config.sections()
print(start_set)
file_collection = input('请输入需要一键启动的启动组:')
print('\r')

# 检测输入是否正确，错误的话要求用户从选项中选择。
while file_collection not in start_set:
    print('启动组名称输入错误，请从下列选项中选择：')
    print(start_set)
    file_collection = input('请输入需要一键启动的启动组:')
    print('\r')
else:
    config_items: list = config.items(file_collection)
    for config_item in config_items:
        os.system(config_item[1])