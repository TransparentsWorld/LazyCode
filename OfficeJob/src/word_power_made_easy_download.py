import requests
import re
from bs4 import BeautifulSoup

url = "https://hanxiaomax.github.io/WordPowerMadeEasy/wpme/How_to_talk_about_personality.html#_1-egoist-%E5%88%A9%E5%B7%B1%E4%B8%BB%E4%B9%89%E8%80%85"
html_doc = requests.get(url).text
soup = BeautifulSoup(html_doc,"lxml")
h2_word_list = soup.select("div.content h2")
li_word_list = soup.select("div.content ul li")

for li_word in li_word_list:
    with open(r'C:\Users\m\Desktop\word.txt','a') as file_object:
        middle_str = re.sub('\s', '', li_word.get_text())
        result_str = middle_str[0:middle_str.find('：')]
        file_object.write(result_str + '\n')