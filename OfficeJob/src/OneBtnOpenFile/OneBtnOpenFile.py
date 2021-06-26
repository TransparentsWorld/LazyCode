import os
import configparser

# 从配置文件读取批处理命令
config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")

print('配置文件中一键启动组如下所示：')
print(config.sections())
file_collection = input('请输入需要一键启动的启动组:')
print('\r')

config_items: list = config.items(file_collection)
for config_item in config_items:
    for config in config_item:
        os.system(config)