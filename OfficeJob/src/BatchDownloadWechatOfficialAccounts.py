import os
import re

import requests
from markdownify import markdownify
from pyquery import PyQuery as pq


# 此脚本用来批量下载渤海小吏微信公众号封建脉络百战的文章

def get_pics_list(wechat_article_url, pq_html):
    """
    获取markdown文件里的所有图片url
    :param markdown:
    :return: 返回包含所有图片url列表
    """
    pics_list = []
    for img in pq_html('img').items():
        # 真是晕死，微信公众号居然用data-src来代替src
        pics_list.append(img.attr('data-src'))
    return pics_list


def get_wechat_article_link(pq_html):
    """
    下载文章链接
    :param pq_html:传入下载好的html页面的pyquery对象
    :return articles_url:
    """
    wechat_articles_url = []
    for item in pq_html("div.rich_media_area_primary_inner p a").items():
        link: str = item.attr("href")
        # 将文章标题名中包含“全”字的链接删除
        if "全" not in item.text():
            wechat_articles_url.append(link)
    return wechat_articles_url


def download_wechat_article_convert_markdown(url, headers):
    """
    下载某篇微信公众号文章
    :param urls: 下载地址
    :param headers: requests模块get函数使用的请求头参数
    :return:
    """

    r_article = requests.get(url, headers=headers)
    article_pq_html = pq(r_article.text)

    # 获取文章标题
    article_title = article_pq_html("h2.rich_media_title").text()

    article_pq_html("p:first-child").prepend(f'<h2>{article_title}</h2>')
    # 为文章插入标题
    article_html = article_pq_html("div.rich_media_content").html()
    article_html = re.sub("<p><br/></p>", "", article_html)
    article_html = re.sub("<br/>", "", article_html)
    # 获取文章内容并删除多余的空行

    # 下载文章保存路径
    article_save_path = os.path.normpath(
        wechat_save_directory + '\\' + article_title + '.md')
    # 下载文章中图片保存路径
    article_img_save_path = os.path.normpath(
        wechat_save_directory + '\\images\\' + article_title)

    print(article_title)

    # 保存html文件中的图片
    if not os.path.exists(article_img_save_path):
        os.mkdir(article_img_save_path)

    img_generator = article_pq_html('div.rich_media_content img').items()
    img_no = len(list(img_generator))
    # 将更改保存在html中。将作出的更改保存，否则不会在结果中显示。例如：多次创建同一个对象，每次作出不同的操作，结果错误。
    # 虽然保存了更改，但输出的值是未作出更改之前的。
    pq_article_html = pq(article_html)
    for count, img in enumerate(article_pq_html('div.rich_media_content img').items()):
        print(f'>>>正在下载第{count + 1}张,共{img_no}张')
        # 真是晕死，微信公众号居然用data-src来代替src
        img_url = img.attr('data-src')
        r_img = requests.get(img_url, headers=headers).content
        with open(os.path.join(article_img_save_path, f'{count + 1}.jpg'), 'w+') as picture:
            picture.buffer.write(r_img)

        pic_relative_path = f'images/{article_title}/{count + 1}.jpg'
        # 本地图片的相对路径
        pq(pq_article_html('img')[count]).attr('data-src', pic_relative_path)
        # 将网络图片更改为本地图片

    # 将微信文章中img标签中的data-src属性替换成src
    article_html = re.sub('data-src', 'src', pq_article_html.html())
    # 保存网页为markdown格式
    article_md = markdownify(article_html)
    with open(article_save_path, "w", encoding="utf-8") as wechat_article:
        wechat_article.write(article_md)


if __name__ == '__main__':
    wechat_html: str = input("请输入文章网址：")
    # wechat_html = r'https://mp.weixin.qq.com/s?__biz=MzUyMzUyNzM4Ng==&mid=100001858&idx=1&sn=bd7345366a61a46d30d1aa85d63e2ab0&chksm=7a3a7df74d4df4e12a162f7a04ee5182154d009ff3fa4bb0e81155e1b3d1720851f8d07635e6&scene=18#wechat_redirect'
    wechat_save_directory = input("请输入文章保存文件夹：")
    # wechat_save_directory = r"C:\Users\Administrator\Desktop\渤海小吏"
    send_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        "Connection": "keep-alive",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.8"
    }

    r = requests.get(wechat_html, headers=send_headers)
    pq_sum_html = pq(r.text)
    wechat_articles_url = get_wechat_article_link(pq_sum_html)
    wechat_articles_no = len(wechat_articles_url)

    for count, wechat_article_url in enumerate(wechat_articles_url):
        print(f'正在下载:第{count + 1}篇，共{wechat_articles_no}篇')
        download_wechat_article_convert_markdown(
            wechat_article_url, headers=send_headers)

    print("下载完毕")
