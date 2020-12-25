# @Time :2020/12/24 10:27
# @Author: Lydia

import datetime, os, sys, time,json,requests,re
import pandas as pd
import numpy as np
import pyautogui
import pyperclip
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import eventlet
import pytesseract
from PIL import Image


# driver = webdriver.Chrome() # 谷歌浏览器
# driver = webdriver.Ie() # IE浏览器

# 从excel中读取待验证的发票信息
def read_xlsx(file_path, time_str):
    data_list = []
    # 筛选出待验证的数据列表
    df = pd.read_excel(file_path, converters={'发票代码': str, '发票号码': str, '不含税金额': str,}, dtype={'开票日期':datetime.datetime})  #
    if not df.empty:
        data_frame = df[df.验真结果.isnull()]
        # 遍历
        for row in data_frame.itertuples():
            invoice_id = getattr(row, '申请流水号')
            invoice_code = getattr(row, '发票代码')
            invoice_num = getattr(row, '发票号码')
            invoice_date = getattr(row, '开票日期')
            invoice_account = getattr(row, '不含税金额')
            item_list = [invoice_id, invoice_code, invoice_num, invoice_date, invoice_account]
            if pd.isnull(invoice_id) or pd.isnull(invoice_code) or pd.isnull(invoice_num) or pd.isnull(
                    invoice_date) or pd.isnull(invoice_account):
                row_index = df[df.申请流水号.isin([invoice_id])].index.tolist()[0]
                df.loc[row_index, '验真结果'] = "参数不全"
                df.loc[row_index, '验真时间'] = time_str
                df.to_excel(file_path, index=False)

            else:
                data_list.append(item_list)

    return data_list, "待验证的数据还有%s条" % len(data_list)

# 写入验证结果
def write_result(check_condition, result, file_path):
    # 在invoice_info.xlsx中查询到该行，将结果写入
    df = pd.read_excel(file_path, converters={'发票代码': str, '发票号码': str, '不含税金额': str}, dtype={'开票日期':datetime.datetime})

    row_index = df[df.申请流水号.isin([invoice_name])].index.tolist()[0]
    msg2 = ''
    if df.empty:
        msg2 = "未筛选到对应信息，无法插入查询结果"
    else:
        df.loc[row_index, '验真结果'] = result
        df.loc[row_index, '验真时间'] = time_str
        df.to_excel(file_path, index=False)
        msg2 = "发票“%s，验真结果：%s，信息已更新" % (invoice_name, result)
    return msg2

# 验证码识别
def verfied_code(code_requirements,VerificationCodeSavePath):
    eventlet.monkey_patch()
    sumcount = 0
    if "红色" in code_requirements:
        color = "01"
    elif "黄色" in code_requirements:
        color = "02"
    elif "蓝色" in code_requirements:
        color = "03"
    else:
        color = "00"
    print("color: ",color)
    URL = "http://rpa-captcha.datagrand.com:8889/rpaservice/captcha"
    picpath = VerificationCodeSavePath
    with eventlet.Timeout(10, False):
        with open(picpath, 'rb') as fr:
            response = requests.post(URL,
                                     files={"file": fr},
                                     data={"way": 3,
                                           "type": 3060,
                                           "casensensitive": 1,
                                           "color": color,
                                           "key": "8E0518293D8ED2867A235FE00C8D1A7C"})
        code_recognition_result = json.loads(response.content).get("data", "")
        # print("原始图片结果：",code_recognition_result)
        if color != "00":
            while len(code_recognition_result) == 6 and sumcount < 5:
                sumcount = sumcount + 1;
                with open(picpath, 'rb') as fr:
                    response = requests.post(URL,
                                             files={"file": fr},
                                             data={"way": 3,
                                                   "type": 3060,
                                                   "casensensitive": 1,
                                                   "color": color,
                                                   "key": "8E0518293D8ED2867A235FE00C8D1A7C"})
                code_recognition_result = json.loads(response.content).get("data", "")

    if not code_recognition_result:
        code_recognition_result = 0

    return code_recognition_result


# 解析查验结果文本
def get_result_text(result_pngpath):
    # 图片识别出文字

    tessdata_dir_config = '--tessdata-dir "C://Program Files (x86)/Tesseract-OCR/tessdata"'

    pytesseract.pytesseract.tesseract_cmd = 'C://Program Files (x86)/Tesseract-OCR/tesseract.exe'

    result_text = pytesseract.image_to_string(Image.open(result_pngpath), lang="eng", config=tessdata_dir_config)
    print("打印结果文本",result_text)


if __name__ == '__main__':

    file_path = "D:\invoice\invoice_info.xls"
    ver_img_rootpath = "D:\invoice\imgs"
    result_img_rootpath = "\invoice\result_imgs"
    ie_driver = "C:\Program Files\Internet Explorer\IEDriverServer.exe"
    os.environ["webdriver.ie.driver"] = ie_driver
    driver = webdriver.Ie(ie_driver)

    driver.get("https://inv-veri.chinatax.gov.cn/index.html")  # 打开国家税务局
    driver.maximize_window()
    now = datetime.datetime.now()
    time_str = now.strftime("%Y-%m-%d")
    data_list, msg = read_xlsx(file_path, time_str)
    print(msg)
    if not data_list:
        sys.exit(0)

    # 获取网页元素
    element_code = driver.find_element_by_id("fpdm")
    element_num = driver.find_element_by_id("fphm")
    element_date = driver.find_element_by_id("kprq")
    element_account = driver.find_element_by_id("kjje")
    element_codejy = driver.find_element_by_id("fpdmjy").text
    element_numjy = driver.find_element_by_id("fphmjy").text
    element_datejy = driver.find_element_by_id("kprqjy").text
    element_accountjy = driver.find_element_by_id("kjjejy").text

    element_vercode = driver.find_element_by_id("yzm")
    element_vercodeinfo = driver.find_element_by_id("yzminfo")
    element_img = driver.find_element_by_id("yzm_img")  # 验证码图片
    element_cybutton = driver.find_element_by_id("checkfp")  # 查验
    above = ActionChains(driver)

    # 遍历待验证的信息
    for item in data_list:
        invoice_name, invoice_code, invoice_num, invoice_date, invoice_account = item[0].replace("\t", ''), item[
            1].replace("\t", ''), item[2].replace("\t", ''), item[3], item[4].replace("\n\t", '')
        # 输入发票信息 以下被注掉的代码太慢
        # element_code.send_keys(invoice_code)
        # element_num.send_keys(invoice_num)
        # element_date.send_keys(invoice_date)
        # element_account.send_keys(invoice_account)
        element_code_js = 'document.getElementById("fpdm").value="%s"' % invoice_code
        driver.execute_script(element_code_js)
        time.sleep(1)
        element_code.click()
        element_num_js = 'document.getElementById("fphm").value="%s"' % invoice_num
        driver.execute_script(element_num_js)
        time.sleep(1)
        element_num.click()
        element_date_js = 'document.getElementById("kprq").value="%s"' % invoice_date
        driver.execute_script(element_date_js)
        time.sleep(1)
        # element_date.click()
        element_account_js = 'document.getElementById("kjje").value="%s"' % invoice_account
        driver.execute_script(element_account_js)
        time.sleep(1)
        element_account.click()
        # 点击验证码输入框
        element_vercode.click()

        # 获取yz信息,假如四要素输入有误，则继续
        error_msg = "有误"
        if error_msg in element_codejy or error_msg in element_numjy or error_msg in element_datejy or error_msg in element_accountjy:
            print(element_codejy, element_numjy, element_datejy, element_accountjy)
            print("参数有误")
            result = "参数有误"
            msg2 = write_result(invoice_name, result, file_path)
            print(msg2)
            continue

        # 等待15s验证码出现
        time.sleep(10)
        # 判断是否出验证码要求
        yzm_require = driver.find_element_by_id("yzminfo").text
        print(yzm_require)
        # 如果没有出现验证码要求，则跳过此条数据，继续循环
        if not yzm_require:
            print("%s：未出验证码要求，跳过"%invoice_name)
            continue
        # 找到图片后右键单击图片
        above.move_to_element(element_img)  # 定位到元素
        above.context_click(element_img)  # 点击右键
        above.perform()  # 执行
        time.sleep(1)
        pyautogui.typewrite(['S'])  # v 是保存的快捷键
        time.sleep(1)  # 等待一秒
        img_path = os.path.join(ver_img_rootpath,invoice_name+".jpg")
        pyperclip.copy(img_path)  # 把指定的路径拷贝过来
        time.sleep(1)  # 等待一秒
        pyautogui.hotkey('ctrlleft', 'v')  # 粘贴
        time.sleep(1)  # 等待一秒
        pyautogui.press('enter')
        if os.path.exists(img_path):
            pyautogui.press('y')
        time.sleep(1)  # 等待一秒
        print("验证码图片保存完成:%s" % img_path)
        verfied_result = verfied_code(yzm_require,img_path)
        element_yzm_js = 'document.getElementById("yzm").value="%s"' % verfied_result
        driver.execute_script(element_yzm_js)
        element_vercode.click()
        time.sleep(0.5)
        # 单击
        element_account.click()
        # 点击查验
        driver.find_element_by_id("checkfp").click()

        time.sleep(2)
        try:
            # 点击按钮后显示弹窗
            driver.find_element(By.ID,"checkfp").click()
            # driver.switch_to.default_content()
            time.sleep(3)
            # 创建弹窗对象
            alert = driver.switch_to.alert
            alert_text = alert.text
            print("弹窗内容为：",alert_text)
            alert.accept() # 点击弹窗中的确定
        except:
            result_pngpath = os.path.join(result_img_rootpath,invoice_name+'jpg')
            # 截取当前窗口，并指定截图图片的保存位置
            driver.get_screenshot_as_file(result_pngpath)
            get_result_text(result_pngpath)
        break
        # driver.refresh()  # 刷新当前页面



