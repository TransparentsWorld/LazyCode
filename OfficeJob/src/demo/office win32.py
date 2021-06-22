import win32com
from win32com.client import Dispatch

#打开word文档
word_file_name:str = r"C:\Users\m\Desktop\test.doc"
word = win32com.client.Dispatch('Word.Application')
# 或者使用下面的方法，使用启动独立的进程：
# w = win32com.client.DispatchEx('Word.Application')

# 后台运行，显示程序界面，不警告
word.Visible = 1 #这个至少在调试阶段建议打开，否则如果等待时间长的话，它至少给你耐心。。。
word.DisplayAlerts = 0

#打开excel文档
excel_file_name:str = r"C:\Users\m\Desktop\导入信息.xls"
excel = win32com.client.Dispatch('Excel.Application')

# 后台excel运行，显示程序界面，不警告
excel.Visible = 1 #这个至少在调试阶段建议打开，否则如果等待时间长的话，它至少给你耐心。。。
excel.DisplayAlerts = 0

# 打开新的文件
worddoc = word.Documents.Open(word_file_name) #打开word文件
excelxls = excel.Workbooks.Open(excel_file_name)#打开excel文件
sheet1 = excelxls.worksheets("Sheet1")

#读取word文档中的信息
paragraph_count:int = worddoc.paragraphs.count
for i in range(1,paragraph_count+1):
    personnel_file_information=worddoc.paragraphs(i).range.text
    #姓名
    name:str = personnel_file_information[0:personnel_file_information.find("同志")]
    #部门和职位
    department_and_position:str  = personnel_file_information[personnel_file_information.find("任")+1:personnel_file_information.rfind("（")]
    position:str  = department_and_position[-4:]
    department:str  = department_and_position[:len(department_and_position) - 4]

    #将word中提取的信息写入到excel中
    sheet1.cells(i + 1, 1).value = name #写入姓名
    sheet1.cells(i + 1, 3).value = department
    sheet1.cells(i + 1, 4).value = position
print("程序执行完毕")
