from OfficeJob.src.GenerateSupervisionJobDataReport.GenerateWeeklyReport import *
from shutil import copyfile
import xlwings as xw
from xlwings import Book, App
import logging


class TestGenerateWord():

    def test_merge_workbook(self):
        template_path: str = r"./模板.xlsx"
        summary_path: str = r"D:\zhi\督察支队\全市督察部门主要业务数据汇总表\数据\2021年\每周汇总数据"
        files_dir_path: str = r"D:\zhi\督察支队\全市督察部门主要业务数据汇总表\数据\2021年\各县区上报数据\5.14-5.20"

        copyfile(template_path, summary_path)
        app: App = None
        summary_wb: Book = None
        try:
            app: App = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb: Book = xw.Book(summary_path)

            gw: GenerateWeeklyReport = GenerateWeeklyReport()
            gw.merge_workbook(summary_wb, files_dir_path)

        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()

    def test_deal_data_complaint(self):
        try:
            begin_date: str = r'2021-4-13'
            end_date: str = r'2021-5-13'
            weekly_summary_path: str = r"C:\Users\m\Desktop\表格\汇总表.xlsx"
            data_sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]
            sheet = data_sheets[0]
            app: App = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb: Book = xw.Book(weekly_summary_path)
            gw: GenerateWeeklyReport = GenerateWeeklyReport()
            last_column = gw.get_sheet_last_column(summary_wb, sheet)
            df = gw.deal_data_complaint(weekly_summary_path, sheet, last_column, begin_date, end_date)
            print(df)
        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()

    def test_deal_data_supervision(self):
        weekly_summary_path: str = r"C:\Users\m\Desktop\表格\汇总表.xlsx"
        data_sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]
        sheet = data_sheets[2]
        try:

            app: App = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb: Book = xw.Book(weekly_summary_path)
            gw: GenerateWeeklyReport = GenerateWeeklyReport()
            last_column = gw.get_sheet_last_column(summary_wb, sheet)
            last_row = gw.get_sheet_last_row(summary_wb, sheet)
            df = gw.deal_data_supervision(weekly_summary_path, sheet, last_column)
            render_dict: dict = gw.generate_render_dict(df, last_row, last_column, '市局')
            print(render_dict)
        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()

    def test_deal_data_rights_protection(self):
        weekly_summary_path: str = r"C:\Users\m\Desktop\表格\汇总表.xlsx"
        data_sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]
        sheet = data_sheets[1]
        try:
            app: App = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb: Book = xw.Book(weekly_summary_path)
            gw: GenerateWeeklyReport = GenerateWeeklyReport()
            last_column = gw.get_sheet_last_column(summary_wb, sheet)
            last_row = gw.get_sheet_last_row(summary_wb, sheet)
            df = gw.deal_data_rights_protection(weekly_summary_path, sheet, last_column)
            render_dict: dict = gw.generate_render_dict(df, last_row, last_column, '市局')
            print(render_dict)
        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()

    def test_generate_render_dict(self):
        weekly_summary_path: str = r"C:\Users\m\Desktop\表格\汇总表.xlsx"
        all_summary_path: str = r"C:\Users\m\Desktop\表格\所有数据汇总表.xlsx"
        data_sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]
        sheet = data_sheets[0]
        try:

            app: App = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb: Book = xw.Book(weekly_summary_path)
            gw: GenerateWeeklyReport = GenerateWeeklyReport()
            last_column = gw.get_sheet_last_column(summary_wb, sheet)
            last_row = gw.get_sheet_last_row(summary_wb, sheet)
            df = gw.deal_data_complaint(weekly_summary_path, sheet, last_column)
            render_dict: dict = gw.generate_render_dict(df, last_row, last_column, '市局')
            print(render_dict)
        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()

    def test_render_docx(self):
        files_dir_path: str = r"C:\Users\m\Desktop\表格\分表"
        weekly_summary_path: str = r"C:\Users\m\Desktop\表格\汇总表.xlsx"
        all_summary_path: str = r"C:\Users\m\Desktop\表格\所有数据汇总表.xlsx"
        template_path: str = r"C:\Users\m\Desktop\表格\模板.docx"
        target_path: str = r"C:\Users\m\Desktop\表格\报告.docx"

        data: dict = {'from_superior': 32, 'from_manager': 24, 'phone': 0, 'record': 0, 'letter': 13,
                      'cop_jurisdiction': 7, 'self_deal': 46, 'supervise_deal': 0, 'archive': 11, 'total_deal_case': 8,
                      'real': 1, 'fake': 7, 'total_deal_person': 0, 'notice': 0, 'confine': 0, 'suspension': 0,
                      'discipline': 0, 'fired_auxiliary_cop': 0, 'from_superior_bj': 12, 'from_manager_bj': 8,
                      'phone_bj': 0, 'record_bj': 0,
                      'letter_bj': 1, 'cop_jurisdiction_bj': 0, 'self_deal_bj': 12, 'supervise_deal_bj': 0,
                      'archive_bj': 0, 'total_deal_case_bj': 0, 'real_bj': 0, 'fake_bj': 0, 'total_deal_person_bj': 0,
                      'notice_bj': 0, 'confine_bj': 0, 'suspension_bj': 0, 'discipline_bj': 0,
                      'fired_auxiliary_cop_bj': 0}

        gw: GenerateWeeklyReport = GenerateWeeklyReport()
        gw.render_docx(data, template_path, target_path)

    def test_modify_dict_key(self):
        modified_dict = {'from_manager': 8, 'phone': 0, 'record': 0, 'letter': 1, 'cop_jurisdiction': 0,
                         'self_deal': 12, 'supervise_deal': 0, 'archive': 0, 'total_deal_case': 0, 'real': 0, 'fake': 0,
                         'total_deal_person': 0, 'notice': 0, 'confine': 0, 'suspension': 0, 'discipline': 0,
                         'fired_auxiliary_cop': 0}
        print(modified_dict.keys())
        print('=========================================')
        # 通过遍历keys()来获取所有的键
        gw: GenerateWeeklyReport = GenerateWeeklyReport()
        dict1 = gw.modify_dict_key(modified_dict)
        print(dict1)

    def test_dataframe_append_excel(self):
        word_template_path: str = r"C:\Users\m\Desktop\表格\模板.docx"
        word_target_path: str = r"C:\Users\m\Desktop\表格\报告.docx"
        weekly_summary_path: str = r"C:\Users\m\Desktop\表格\汇总表.xlsx"
        all_summary_path: str = r"C:\Users\m\Desktop\表格\所有数据汇总表.xlsx"
        files_dir_path: str = r"C:\Users\m\Desktop\表格\分表"

        sheets: list = ["举报投诉核查工作", "维权工作", "督察工作情况"]
        app: App = None
        summary_wb: Book = None

        begin_date: str = r'2021-4-13'
        end_date: str = r'2021-5-13'
        gw: GenerateWeeklyReport = GenerateWeeklyReport()

        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            summary_wb = xw.Book(weekly_summary_path)

            # 举报投诉工作表
            last_column = gw.get_sheet_last_column(summary_wb, sheets[0])
            last_row = gw.get_sheet_last_row(summary_wb, sheets[0])
            df_complaint = gw.deal_data_complaint(weekly_summary_path, sheets[0], last_column, begin_date, end_date)
            gw.dataframe_append_excel(df_complaint, all_summary_path, sheets[0])

        except BaseException as e:
            logging.exception(e)
        finally:
            summary_wb.save()
            summary_wb.close()
            app.quit()

