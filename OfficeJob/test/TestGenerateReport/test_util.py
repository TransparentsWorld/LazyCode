from OfficeJob.src.GenerateReport.Util import *
from shutil import copyfile
import xlwings as xw
from xlwings import Book, App
import logging



class TestUtil():
    def test_merge_workbook(self):
        template_path: str = r"./模板.xlsx"
        summary_path: str = r"C:\Users\Administrator\Desktop"
        files_dir_path: str = r"C:\Users\Administrator\Desktop\新建文件夹"

        copyfile(template_path, summary_path)
        app: App = None
        summary_wb: Book = None
        try:
            app: App = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb: Book = xw.Book(summary_path)
            Util.merge_workbook(summary_wb, files_dir_path)

        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()
