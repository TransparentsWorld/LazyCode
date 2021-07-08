import requests
from pyquery import PyQuery as pq
from OfficeJob.src.BatchDownloadWechatOfficialAccounts import *

class TestBatchDownloadWechatOfficialAccount():
    def test_get_wechat_article_link(self):
        # wechat_html: str = input("请输入文章网址：")
        wechat_html = r'https://mp.weixin.qq.com/s?__biz=MzUyMzUyNzM4Ng==&mid=100012844&idx=1&sn=209bb13d6f9a8c545bcee61ec392f431&chksm=7a3986994d4e0f8f0f6fab78d5697c21a4fc10026a14bfcf08534039827139a3dfe04c04afbc&scene=18#wechat_redirect'

        send_headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            "Connection": "keep-alive",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.8"
        }

        r = requests.get(wechat_html, headers=send_headers)
        pq_sum_html = pq(r.text)
        wechat_articles_url = get_wechat_article_link(pq_sum_html)
        for url in wechat_articles_url:
            print(url)
