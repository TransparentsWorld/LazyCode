from OfficeJob.src.OfficeUtil import *
import codecs
import configparser


def compare_list(dir_path: str, stardrd: list):
    comparsion: list = OfficeUtil.get_filename_in_dir(dir_path)
    resultsets: list = list(set(stardrd) - set(comparsion))
    return resultsets


if __name__ == '__main__':
    output_file_path: str = os.path.join(OfficeUtil.get_desktop_path(), "比较结果.txt")
    dir_path: str = input("请输入文件夹所在路径：")
    select_no: int = input('请输入序号用来选择比较条件（1为各县区，2为局属各部门，3为督察支队）：')
    resultsets: list = []

    # 从配置文件读取文件路径
    # cf = configparser.SafeConfigParser()
    # with codecs.open('./config.ini', 'r', encoding='utf-8') as f:
    #     cf.readfp(f)
    # secs = cf.sections()
    # all_bj_bureau: list = cf.get('comparsion', 'all_bj_bureau')
    # departments: list = cf.get('comparsion', 'departments')
    # supervision_division_staff: list = cf.get('comparsion', 'supervision_division_staff')
    # print(all_bj_bureau)

    all_bj_bureau: list = ['七星关', '大方', '黔西', '金沙', '织金', '纳雍', '威宁', '赫章', '百里杜鹃', '金海湖', '洪家渡']
    departments: list = ["办公室", "政治部", "机关党委", "督察", "审计", "警保", "警卫", "刑侦", "扫黑", "法制", "治安", "经侦", "监管",
                         "禁毒", "技侦", "国保", "维稳", "反恐", "网安", "情报", "科通", "培训学校", "出入境", "特警", "机场", "交警"]
    supervision_division_staff: list = ['黄晓咏', '何万学', '肖奎', '沈晓冬', '高守红', '杨渊', '臧红艳', '温卫忠', '安忠涛', '舒畅', '陈寿远', '李仁银',
                                        '徐昌辉', '许猛', '支军焱']
    if select_no == '1':
        resultsets = compare_list(dir_path, all_bj_bureau)
    elif select_no == '2':
        resultsets = compare_list(dir_path, departments)
    elif select_no == '3':
        resultsets = compare_list(dir_path, supervision_division_staff)

    with open(output_file_path, "w") as f:
        for resultset in resultsets:
            f.write(resultset + "\n")
