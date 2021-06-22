import os
import os.path
import win32com.client as win32


class OfficeUtil:

    @staticmethod
    def excel_format_convertion(self, file_dir: str) -> None:
        '''
        将同一个文件夹下的excel文件从xls格式转换为xlsx格式
        :param file_dir: 文件夹路径
        :return: None
        '''

        # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
        for parent, dirnames, filenames in os.walk(file_dir):
            for fn in filenames:
                filedir = os.path.join(parent, fn)
                print(filedir)

                excel = win32.gencache.EnsureDispatch('Excel.Application')
                wb = excel.Workbooks.Open(filedir)
                # xlsx: FileFormat=51
                # xls:  FileFormat=56,
                # 后缀名的大小写不通配，需按实际修改：xls，或XLS
                wb.SaveAs(filedir.replace('xls', 'xlsx'), FileFormat=51)  # 我这里原文件是大写
                wb.Close()
                excel.Application.Quit()

    @staticmethod
    def path_convert_filename(filePath: str) -> str:
        '''
        从文件路径中提取不带后缀的文件名
        :return: 不带后缀的文件名
        '''
        return os.path.splitext(os.path.basename(filePath))[0]

    @staticmethod
    def get_dir_all_file_path(dir_path: str) -> list:
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

    @staticmethod
    def get_desktop_path() -> str:
        '''
        获取windows系统桌面文件夹所在路径
        :return: 桌面文件夹所在路径
        '''
        return os.path.join(os.path.expanduser("~"), "desktop")

    @staticmethod
    def get_filename_in_dir(dir_path: str) -> list:
        '''
        提取某文件夹下所有文件的不带扩展名的文件名
        :param dir_path: 文件夹路径
        :return: 文件名组成的列表
        '''
        files_name: list = []
        for root, dir, paths in os.walk(dir_path):
            for path in paths:
                file_name: str = OfficeUtil.path_convert_filename(path)
                files_name.append(file_name)
        return files_name

    @staticmethod
    def get_filename_with_extension_in_dir(dir_path: str) -> list:
        '''
        返回文件夹及其子文件夹下所有文件的带扩展名的文件名
        :param dir_path: 根目录
        :return: 所有文件名构成的列表
        '''
        files_path: list = []
        for root, dir, paths in os.walk(dir_path):
            for path in paths:
                files_path.append(path)
        return files_path

    # @staticmethod
    # def get_dir_file_absolute_path(dir_path: str) -> list:
    #     files_path = os.listdir(dir_path)
    #     absolute_paths:list = []
    #     for file_path in files_path:
    #         absolute_path = os.path.join(dir_path, file_path)
    #         # isdir和isfile参数必须跟绝对路径
    #         if os.path.isdir(absolute_path):
    #             OfficeUtil.get_dir_file_absolute_path(absolute_path)
    #         else:
    #             absolute_paths = os.path.join(dir_path, absolute_path)
    #     return absolute_paths

    @staticmethod
    def get_file_absolute_path_in_dir(dir_path: str) -> list:
        '''
        返回文件夹及其子文件夹下所有文件的绝对路径
        :param dir_path: 文件夹路径
        :return: 所有文件的绝对路径
        '''
        paths = os.listdir(dir_path)
        files_path = list()

        for path in paths:
            absolute_path = os.path.join(dir_path, path)  # 拼接得到绝对路径
            if os.path.isdir(absolute_path):  # 判断当前路径是不是一个文件夹
                files_path = files_path + OfficeUtil.get_file_absolute_path_in_dir(absolute_path)
            else:
                files_path.append(absolute_path)
        return files_path
