import requests
from bs4 import BeautifulSoup
import xlwings as xw

same_city_used_apartment_url = r"https://bijie.58.com/bjqxg/ershoufang/h1/pn2/?&area=100_120&bunengdaikuan=0&PGTID=0d30000c-03a7-2e7e-98d8-443242930681&ClickID=1"
# r = requests.get(url)
# print(r.text)
page_count = 71
# 获取页面url
# def get_result_url(base_url,page_count):
#     result_url_list = []
#     for page_number in range(1,page_count):
#         pn_start_number = base_url.find("pn")
#         pn_end_number = base_url.find("/",pn_start_number)
#         result_url = base_url[:pn_start_number]+"pn"+ str(page_number) + base_url[pn_end_number:]
#         result_url_list.append(result_url)
#     return result_url_list

#beautifulsoup4解析html页面
html_doc = requests.get(same_city_used_apartment_url).text
soup = BeautifulSoup(html_doc,"lxml")
apartment_name_list = soup.select("ul.house-list-wrap h2.title a")
apartment_type_list = soup.select("ul.house-list-wrap div.list-info > p span")
apartment_sum_list = soup.select("ul.house-list-wrap div.price p.sum b")
apartment_unit_list = soup.select("ul.house-list-wrap div.price p.unit")


# 将数据写入excel
wb = xw.Book(r"C:\Users\m\Desktop\毕节房价表.xlsx")
sheet = wb.sheets[0]

# 提取列表中的tag对象的文本
apartment_name_text_list = [] #房源
apartment_type_text_list = [] #户型、面积、朝向、层高、位置
apartment_sum_text_list = [] #总价
apartment_unit_text_list = [] #单价
apartment_href_text_list = [] #链接

for apartment_name in apartment_name_list:
    apartment_name_text_list.append(apartment_name.get_text())

for apartment_href in apartment_name_list:
    apartment_href_text_list.append(apartment_href["href"])

for apartment_type in apartment_type_list:
    apartment_type_text_list.append(apartment_type.get_text())

for apartment_sum in apartment_sum_list:
    apartment_sum_text_list.append(apartment_sum.get_text())

for apartment_unit in apartment_unit_list:
    apartment_unit_text_list.append(apartment_unit.get_text())

# 数据写入单元格
item_per_page = 69
for i in range(0,page_count):
    sheet.range("B"+str(2+i*item_per_page)).options(transpose=True).value = apartment_name_text_list

    sheet.range("C"+str(2+i*item_per_page)).options(transpose=True).value = apartment_type_text_list[::5] #户型
    sheet.range("D"+str(2+i*item_per_page)).options(transpose=True).value = apartment_type_text_list[1::5] #面积
    sheet.range("E"+str(2+i*item_per_page)).options(transpose=True).value = apartment_type_text_list[2::5] #朝向
    sheet.range("F"+str(2+i*item_per_page)).options(transpose=True).value = apartment_type_text_list[3::5] #层高
    sheet.range("G"+str(2+i*item_per_page)).options(transpose=True).value = apartment_type_text_list[4::5] #位置

    sheet.range("H"+str(2+i*item_per_page)).options(transpose=True).value = apartment_sum_text_list
    sheet.range("I"+str(2+i*item_per_page)).options(transpose=True).value = apartment_unit_text_list
    sheet.range("J"+str(2+i*item_per_page)).options(transpose=True).value = apartment_href_text_list


