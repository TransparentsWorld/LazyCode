from unittest import TestCase

from OfficeJob.src.OfficeUtil import *


class TestOfficeUtil():
    def test_get_filename_in_dir(self):
        path = r'C:\Users\m\Desktop\表格'
        file_name = OfficeUtil.get_filename_in_dir(path)
        print(file_name)

    def test_get_filename_with_extension_in_dir(self):
        path = r'C:\Users\m\Desktop\表格'
        file_name = OfficeUtil.get_filename_with_extension_in_dir(path)
        print(file_name)

    def test_get_dir_file_absolute_path(self):
        path = r'C:\Users\m\Desktop\表格'
        absolute_paths: list = []
        files_path = OfficeUtil.get_dir_file_absolute_path(path)
        print(files_path)

    def test_get_file_absolute_path_in_dir(self):
        path = r'C:\Users\m\Desktop\表格'
        files_path = OfficeUtil.get_file_absolute_path_in_dir(path)
        for file_path in files_path:
            print(file_path)

    def test_get_dir_all_file_path(self):
        path = r'C:\Users\Administrator\Desktop\2021'
        files_path = OfficeUtil.get_file_absolute_path_in_dir(path)
        for file_path in files_path:
            print(file_path)
