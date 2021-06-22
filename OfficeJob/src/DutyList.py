from shutil import copyfile
import pandas as pd
from pandas import DataFrame, ExcelWriter
from openpyxl import load_workbook
from OfficeJob.src.OfficeUtil import *


class DutyList:
    '''
        从每月值班表中自动生成每周值班表。
    '''
    def dataframe_append_excel(self, df: DataFrame, path: str, sheetname: str, start_row: int, start_col: int) -> None:
        """
        将pandas中的dataframe数据类型粘贴到excel工作表中。
        :param df: 需要追加的dataframe
        :param path: 全部数据汇总工作簿路径
        :param sheetname: 写入的工作表名
        :param start_row: 粘贴的开始行
        :param start_col: 粘贴的开始列
        :return: None
        """
        book = load_workbook(path)
        writer: ExcelWriter = ExcelWriter(path, engine='openpyxl', date_format="yyyy-mm-dd")
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}

        df.to_excel(writer, sheet_name=sheetname, startrow=start_row, startcol=start_col, index=False,
                    header=False)
        writer.save()

    def filter_duty_list_by_date(self, begin_date: str, end_date: str, file_path: str) -> DataFrame:
        df: DataFrame = pd.read_excel(io=file_path, sheet_name="值班备勤", skiprows=4, skipfooter=1,
                                      header=None, usecols=[0, 1, 2, 3, 4, 5, 6])

        # 为df赋新的列名
        field_name = ['date', 'department', 'person', 'time', 'phone_no', 'supervisor', 'situation']
        df.columns = field_name
        df = df.fillna('')
        df_date = df[(df.date >= begin_date) & (df.date <= end_date)]
        return df_date

    def get_dir_all_file(self, dir_path: str) -> list:
        """
        获取某个文件夹下的所有文件路径，不会获取子文件夹下的文件路径
        :param dir_path: 需要遍历的文件夹的路径
        :return: 返回文件夹的路径下的所有文件的路径
        """
        file_paths: list = os.listdir(dir_path)
        temp_paths: list = []
        for file_path in file_paths:
            temp_paths.append(os.path.join(dir_path, file_path))
        return temp_paths


if __name__ == '__main__':
    begin_date: str = r'2021-05-06'
    end_date: str = r'2021-05-09'

    template_duty_list_path: str = r"D:\zhi\督察支队\工作台账\值班表\平时值班模板.xlsx"
    monthly_list_dir_path = r'D:\zhi\督察支队\工作台账\值班表\2021年\5月值班表'
    weekly_list_dir_path = r"D:\zhi\督察支队\工作台账\值班表\2021年\5.06-5.09"

    dl: DutyList = DutyList()
    #创建每周工作表存放的文件夹
    if os.path.exists(weekly_list_dir_path) == False:
        os.makedirs(weekly_list_dir_path)

    monthly_list_files_path: list = dl.get_dir_all_file(monthly_list_dir_path)

    for monthly_list_file_path in monthly_list_files_path:
        excel_name = OfficeUtil.path_convert_filename(monthly_list_file_path) + '.xlsx'
        weekly_list_file_path = os.path.join(weekly_list_dir_path, excel_name)
        print(excel_name)

        copyfile(template_duty_list_path, weekly_list_file_path)
        df: DataFrame = dl.filter_duty_list_by_date(begin_date, end_date, monthly_list_file_path)
        df['date'] = df['date'].dt.date  # 将datetime转为date
        dl.dataframe_append_excel(df, weekly_list_file_path, '值班备勤', 4, 0)
