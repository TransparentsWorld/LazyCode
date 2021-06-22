from shutil import copyfile
from docxtpl import DocxTemplate
import copy
import logging
from datetime import datetime
import configparser
import codecs
import pandas as pd
import xlwings as xw
from openpyxl import load_workbook
from pandas import DataFrame, ExcelWriter
from xlwings import Book, App, Sheet
from OfficeJob.src.OfficeUtil import *


class Util:

    @staticmethod
    def modify_dict_key(modified_dict: dict) -> dict:
        """
        为字典的key添加"_bj"后缀，方便后期渲染毕节的汇总数据
        :param modified_dict: 被修改的字典
        :return: 修改后的字典
        """
        temp_dict: dict = {}
        for key in modified_dict.keys():
            temp_key: str = key + "_bj"
            temp_dict[temp_key] = modified_dict[key]
        modified_dict.clear()  # 删除字典中所有数据
        modified_dict = temp_dict  # 将临时字典的值赋值给modified_dict
        return modified_dict

    @staticmethod
    def render_docx(report_data: dict, template_path: str, target_path) -> None:
        """
        用数据生成成报告
        :param report_data: 生成报告所使用的数据
        :param template_path: word模板所在位置
        :param target_path: word报告生成后存放位置
        :return:
        """
        copyfile(template_path, target_path)
        doc: DocxTemplate = DocxTemplate(target_path)
        doc.render(report_data)
        doc.save(target_path)

    @staticmethod
    def generate_report_filename(begin_date: str, end_date: str) -> str:
        '''
        将日期改写成2.5-3.15的格式
        :param begin_date: 起始日期，如：2021-2-5
        :param end_date: 结束日期，如：2021-3-15
        :return: 返回2.5-3.15作为报告的文件名
        '''
        begin_result = begin_date.split('-')
        end_result = end_date.split('-')
        start_date = begin_result[1] + '.' + begin_result[2]
        finish_date = end_result[1] + '.' + end_result[2]
        return start_date + '-' + finish_date

    @staticmethod
    def generate_date_in_monthly_report(begin_date: str, end_date: str) -> str:
        '''
        将日期改写成2021年2月——3月的格式
        :param begin_date: 起始日期，如：2021-2-5
        :param end_date: 结束日期，如：2021-3-15
        :return: 返回2021年2月——3月
        '''
        begin_result = begin_date.split('-')
        end_result = end_date.split('-')
        year = begin_result[0]
        begin_month = begin_result[1]
        end_month = end_result[1]
        return '{}年{}月-{}月'.format(year, begin_month, end_month)

    @staticmethod
    def generate_date_in_weekly_report(begin_date: str, end_date: str) -> str:
        '''
        将日期改写成2021年2月——3月的格式
        :param begin_date: 起始日期，如：2021-2-5
        :param end_date: 结束日期，如：2021-3-15
        :return: 返回2021年2月——3月
        '''
        begin_result = begin_date.split('-')
        end_result = end_date.split('-')
        year = begin_result[0]
        begin_month = begin_result[1]
        begin_day = begin_result[2]
        end_month = end_result[1]
        end_day = end_result[2]
        return '{}年{}月{}日-{}月{}日'.format(year, begin_month, begin_day, end_month, end_day)

    @staticmethod
    def get_sheet_last_row(wb: Book, sheet: Sheet) -> int:
        '''
        获取工作表最后一行
        :param wb: 工作簿
        :param sheet: 工作表
        :return: 返回最后一行
        '''
        last_cell = wb.sheets[sheet].used_range.last_cell
        last_row = last_cell.row + 1
        return last_row

    @staticmethod
    def get_sheet_last_column(wb: Book, sheet: Sheet) -> int:
        '''
        获取工作表最后一列
        :param wb: 工作簿
        :param sheet: 工作表
        :return: 返回最后一列
        '''
        last_cell = wb.sheets[sheet].used_range.last_cell
        last_column = last_cell.column
        return last_column

    @staticmethod
    def merge_workbook(weekly_summary_wb: Book, files_dir_path: str) -> None:
        '''
        合并各县区上报的数据
        :param weekly_summary_wb: 每周汇总数据表
        :param files_dir_path: 各县区上报表格所在的文件夹路径
        :return: 无返回值
        '''
        sub_wb: Book = None  # 用于汇总的分表
        data_sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]  # 工作簿中的三张工作表

        # excel_paths = self.get_dir_all_file(files_dir_path)
        excel_paths = OfficeUtil.get_dir_all_file_path(files_dir_path)
        # 循环访问文件夹中的工作簿
        print("正在合并")
        for excel_path in excel_paths:
            try:
                sub_wb: Book = xw.Book(excel_path)
                print("-->" + excel_path)
                # 循环访问工作簿中的三张工作表
                for data_sheet in data_sheets:
                    last_row: int = Util.get_sheet_last_row(weekly_summary_wb, data_sheet)
                    last_column: int = Util.get_sheet_last_column(weekly_summary_wb, data_sheet)
                    # 通过复制粘贴来合并工作表
                    weekly_summary_wb.sheets[data_sheet].range((last_row, 1), (last_row, last_column)).value = \
                        sub_wb.sheets[data_sheet].range((3, 1), (3, last_column)).value
                    last_row = last_row + 1
            except BaseException as e:
                logging.exception(e)
            finally:
                weekly_summary_wb.save()
                sub_wb.close()
        print('合并完毕')

