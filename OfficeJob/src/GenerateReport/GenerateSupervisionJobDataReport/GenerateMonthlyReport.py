import logging
import pandas as pd
from pandas import DataFrame, Series
import codecs
import configparser

from OfficeJob.src.GenerateSupervisionJobDataReport.Util import *
from OfficeJob.src.OfficeUtil import *


class GenerateMonthlyReport:
    '''
    生成督察支队每年每季度每月周报。
    '''

    def statistics_data(self, df: DataFrame, begin_date: str, end_date: str, last_column: int, unit: str) -> Series:

        if unit == '全市':
            df_all: DataFrame = df[(df['begin_date'] >= begin_date) & (df['end_date'] <= end_date)]

            # 清洗数据，把nan值替换成0，把str替换成int
            df_all = df_all.fillna(0)
            df_int_all = df_all.iloc[:, 3:last_column].astype('int')

            sum: Series = df_int_all.apply(lambda x: x.sum())
            return sum
        elif unit == '市局':
            df_bj: DataFrame = df[
                (df['unit'] == unit) & (df['begin_date'] >= begin_date) & (df['end_date'] <= end_date)]

            # 清洗数据，把nan值替换成0，把str替换成int
            df_bj = df_bj.fillna(0)
            df_int_bj = df_bj.iloc[:, 3:last_column].astype('int')

            sum_bj: Series = df_int_bj.apply(lambda x: x.sum())
            return sum_bj

    def generate_render_dict(self, series: Series, unit: str) -> dict:
        '''
        将series转成字典，用于生成word报告。
        :param series: series
        :param unit: 单位，用来区分生成全市数据或者市局数据
        :return: 返回字典
        '''
        dictionary: dict = series.astype('int').to_dict()
        if unit == '全市':
            return dictionary
        elif unit == '市局':
            dictionary = Util.modify_dict_key(dictionary)  # 修改字典名
            return dictionary


if __name__ == '__main__':

    print('本程序用来汇总统计督察每月每季度每年的主要业务数据，并生成报告。请按下面提示操作。')
    begin_date = input('请输入起始时间(如2021-4-13):')
    end_date = input('请输入结束时间(如2021-5-13):')

    date: str = Util.generate_date_in_monthly_report(begin_date, end_date)

    # 从配置文件读取文件路径
    cf = configparser.ConfigParser()
    with codecs.open('./config.ini', 'r', encoding='utf-8') as f:
        cf.read_file(f)
    secs = cf.sections()
    all_summary_path: str = cf.get('Util', 'all_summary_path')
    word_template_path: str = cf.get('Util', 'word_template_path')
    monthly_report: str = cf.get('Util', 'report')
    word_target_path: str = os.path.join(monthly_report, date + '督察主要业务数据') + '.docx'

    sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]
    try:
        gmr: GenerateMonthlyReport = GenerateMonthlyReport()

        # 处理举报投诉表
        df_complaint: DataFrame = pd.read_excel(io=all_summary_path, engine='openpyxl', sheet_name=sheets[0], header=0)
        complaint_last_column: int = df_complaint.shape[1]  # 获取表最后一列
        # 处理毕节数据
        bj_complaint_series: Series = gmr.statistics_data(df_complaint, begin_date, end_date, complaint_last_column,
                                                          '市局')
        bj_complaint_dictionary: dict = gmr.generate_render_dict(bj_complaint_series, '市局')

        # 处理全市数据
        complaint_series: Series = gmr.statistics_data(df_complaint, begin_date, end_date, complaint_last_column, '全市')
        complaint_dictionary: dict = gmr.generate_render_dict(complaint_series, '全市')
        # 汇总举报投诉数据到同一个字典中
        complaint_dictionary.update(bj_complaint_dictionary)

        # 处理维权表
        df_rights_protection: DataFrame = pd.read_excel(io=all_summary_path, engine='openpyxl', sheet_name=sheets[1],
                                                        header=0)
        rights_protection_last_column: int = df_rights_protection.shape[1]  # 获取表最后一列
        bj_rights_protection_series: Series = gmr.statistics_data(df_rights_protection, begin_date, end_date,
                                                                  rights_protection_last_column, '市局')
        bj_rights_protection_dictionary: dict = gmr.generate_render_dict(bj_rights_protection_series, '市局')
        rights_protection_series: Series = gmr.statistics_data(df_rights_protection, begin_date, end_date,
                                                               rights_protection_last_column,
                                                               '全市')
        rights_protection_dictionary: dict = gmr.generate_render_dict(rights_protection_series, '全市')
        rights_protection_dictionary.update(bj_rights_protection_dictionary)

        # 督察工作表
        df_supervision: DataFrame = pd.read_excel(io=all_summary_path, engine='openpyxl', sheet_name=sheets[2],
                                                  header=0)
        supervision_last_column: int = df_supervision.shape[1]  # 获取表最后一列
        bj_supervision_series: Series = gmr.statistics_data(df_supervision, begin_date, end_date,
                                                            supervision_last_column, '市局')
        bj_supervision_dictionary: dict = gmr.generate_render_dict(bj_supervision_series, '市局')
        supervision_series: Series = gmr.statistics_data(df_supervision, begin_date, end_date,
                                                         supervision_last_column, '全市')
        supervision_dictionary: dict = gmr.generate_render_dict(supervision_series, '全市')
        supervision_dictionary.update(bj_supervision_dictionary)

        # 汇总字典数据
        render: dict = {}
        render.update(complaint_dictionary)
        render.update(rights_protection_dictionary)
        render.update(supervision_dictionary)
        render['date'] = date

        # 生成word报告
        Util.render_docx(render, word_template_path, word_target_path)
        print('报告生成完毕，位置为' + word_target_path)
    except BaseException as e:
        logging.exception(e)
