import os
import configparser
from OfficeJob.src.OfficeUtil import *


def compare_list(dir_path: str, stardrd: list):
    comparsion: list = OfficeUtil.get_filename_in_dir(dir_path)
    resultsets: list = list(set(stardrd) - set(comparsion))
    return resultsets


if __name__ == '__main__':
    print('本程序通过文件名比对来判断材料是否交齐全。')
    output_file_path: str = os.path.join(OfficeUtil.get_desktop_path(), "比较结果.txt")
    dir_path: str = input("请输入文件夹所在路径：")

    # 判断文件夹是否存在，不存在则重新输入
    while not os.path.exists(dir_path):
        print('文件夹不存在，请重新输入正确文件夹路径：')
        dir_path: str = input("请输入文件夹所在路径：")
    else:
        config = configparser.ConfigParser()
        config.read("config.ini", encoding="utf-8")
        config_items: list = config.items("comparsion")

        # 将列表转换成字典，将字符串转换成列表
        config_dicts: dict = {}
        for config_item in config_items:
            config_dicts[config_item[0]] = config_item[1].split(',')  # 用split函数将字符串转换成列表

        print('\r')
        print("筛选条件：")
        for key, value in config_dicts.items():
            print("-->筛选条件名：%s" % (key))
            print("-->筛选条件值：%s" % (value))
            print('\r')

        # 输入筛选条件
        index: str = input('请输入筛选条件名：')
        resultsets = compare_list(dir_path, config_dicts[index])

        # 将结果写入txt文件
        with open(output_file_path, "w") as f:
            for resultset in resultsets:
                f.write(resultset + "\n")
