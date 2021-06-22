from docxtpl import DocxTemplate
from shutil import copyfile
import os
import xlwings as xw
from xlwings import Book, App, Sheet
import logging
import pandas as pd
from pandas import DataFrame, ExcelWriter
from openpyxl import load_workbook
import copy
from datetime import datetime
from src.DutyList import *


class TestDutyList():
    def test_filter_duty_list_by_date(self):
        begin_date: str = r'2021-05-06'
        end_date: str = r'2021-05-21'
        monthly_list_file_path: str = r'C:\Users\m\Desktop\5月值班表\政治部.xls'
        template_duty_list_path: str = r'C:\Users\m\Desktop\值班表模板.xlsx'
        temp_path = r'C:\Users\m\Desktop\test.xlsx'

        copyfile(template_duty_list_path, temp_path)

        dl: DutyList = DutyList()
        df: DataFrame = dl.filter_duty_list_by_date(begin_date, end_date, monthly_list_file_path)
