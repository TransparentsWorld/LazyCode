from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random, time
import pyttsx3

'''
    学法用法自动刷分程序，但需要每次都手动输入验证码。
'''

users = {
    # "吴健":{"5207022030022":"tiny954"},
    "支军焱":{"5207010630014":"tiny954"},
}


for owner,account in users.items():
    for username,password in account.items():
        browser = webdriver.Chrome()
        browser.maximize_window()
        browser.get("http://xf.faxuan.net/bps/index.html")
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "form#loginForm input.btn_denglu")))

        # def discern_verification_code(location):
        #     screenshot_name = 'screenshot_windows.png'
        #     code_name = 'code.png'
        #     browser.save_screenshot(screenshot_name)  # 获取屏幕截图
        #     img = Image.open(screenshot_name)
        #     img = img.crop(location)  # 获取验证码截图
        #     img = img.convert('L')  # 图片灰度处理
        #     img.show()
        #     img.save(code_name)
        #     img = Image.open(code_name)
        #     codes = pytesseract.image_to_string(img)  # 识别图文验证码
        #     code = ''
        #     for i in codes.strip():  # 正则表达式去除特殊字符
        #         pattern = re.compile(r'[a-zA-Z0-9]')
        #         m = pattern.search(i)
        #         if m != None:
        #             code += i
        #     return code

        username_input = browser.find_elements_by_id("userAccount")[0] #输入用户名
        username_input.send_keys(username)
        password_input = browser.find_elements_by_id("userPassword")[0] #输入密码
        password_input.send_keys(password)
        # verification_code: str = discern_verification_code(screen_shot_location)  # 识别后的验证码

        engine = pyttsx3.init()
        engine.say('请输入验证码')
        engine.runAndWait()
        verification_code = input("请输入验证码：")#输入验证码
        browser.find_elements_by_id("usercheckcode")[0].send_keys(verification_code)
        browser.find_element_by_css_selector("form#loginForm input.btn_denglu").click()#点击登录
        WebDriverWait(browser,10).until(EC.presence_of_element_located((By.ID,"page")))

        for i in range(1,5):
            browser.find_element_by_css_selector("ul#page li.clear div a h3:nth-child(1)").click()#选择课程
            #等待3秒候切换到新标签页
            browser.implicitly_wait(3)
            #切换标签页
            windows = browser.current_window_handle
            all_handles = browser.window_handles
            for handle in all_handles:
                if handle != windows:
                    browser.switch_to.window(handle)

            #等待10分钟后退出
            random_learn_time = random.randrange(620,650)
            print("用户%s开始第%d次学习，线程停止%d秒" % (owner,i,random_learn_time))
            time.sleep(random_learn_time)
            browser.find_element_by_css_selector("div#mainyemian2 div.lessoncontainer div.timebtn a").click()#点击退出学习按钮
            time.sleep(3)
            browser.find_element_by_css_selector("a#popwinConfirm").click()#点击确定退出学习按钮

            #切换标签页
            all_handles = browser.window_handles
            for handle in all_handles:
                browser.switch_to.window(handle)
        browser.quit()
print("刷分结束")