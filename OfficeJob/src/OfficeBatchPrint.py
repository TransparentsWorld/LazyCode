import os
import win32com
import xlwings as xw
import logging
from OfficeJob.src.OfficeUtil import *


class OfficeBatchPrint:
    '''
        批量打印
    '''

    def excel_batch_print(file_path: str) -> None:
        """
        打印一个excel工作簿
        :param file_path: 要打印的excel工作簿路径
        :return:
        """

        file_name = os.path.basename(file_path)  # 通过文件路径提取文件名
        print("--> " + file_name) # 在控制台显示当前正在打印的文件的文件名

        wb = xw.Book(file_path)
        wb.api.PrintOut()
        wb.close()

    def word_batch_print(file_path: str, word) -> None:
        """
        打印一个word文档
        :param file_path: 要打印的word文档路径
        :param word: word程序的实例
        :return:None
        """

        # 在控制台显示当前正在打印的文件的文件名
        file_name = os.path.basename(file_path)  # 通过文件路径提取文件名
        print("--> " + file_name)

        doc = word.Documents.Open(file_path)
        doc.PrintOut()
        doc.Close()

    if __name__ == '__main__':
        print("此程序仅能批量打印处于同一文件夹下的word和excel文档，子文件夹下的文件不能打印")
        print()
        dir_paths: str = input("请输入文件夹所在路径：")
        file_paths: str = OfficeUtil.get_dir_all_file_path(dir_paths)

        try:
            word = win32com.client.Dispatch('Word.Application')
            # 初始化word实例
            word.Visible = 0  # 后台运行
            word.DisplayAlerts = 0  # 不显示，不警告

            # 初始化excel实例
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            # 循环打印word或excel文件
            print()
            print("===================  正在打印  ===================")
            for file_path in file_paths:
                if file_path.endswith("xls") or file_path.endswith("xlsx"):
                    excel_batch_print(file_path)
                elif file_path.endswith("doc") or file_path.endswith("docx"):
                    word_batch_print(file_path, word)
            print('打印完成')
        except BaseException as e:
            logging.exception(e)
        finally:
            word.Quit()
            app.quit()
