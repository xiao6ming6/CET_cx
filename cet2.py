#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @author  : xiumi
# @time    : 2024/12/13 16:12
# @function: the script is used to do ...
# @version : V1
import time
import pandas as pd
from selenium.common import TimeoutException
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def get_info(user_name, id_code):
    # 设置 Chrome 浏览器选项
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")  # 禁用 GPU 加速
    chrome_options.add_argument("--no-sandbox")  # 避免一些浏览器启动错误

    # 启动浏览器
    driver = uc.Chrome(options=chrome_options, driver_executable_path=ChromeDriverManager().install())

    # 访问页面
    driver.get("https://cet-bm.neea.edu.cn/Home/QuickPrintTestTicket")

    # 显式等待，直到 id 为 selProvince 的元素加载完成
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "selProvince")))

    # 选择 selProvince 下拉框中的 value=41
    province_select = Select(driver.find_element(By.ID, "selProvince"))
    province_select.select_by_value("41")

    # 选择 selIDType 下拉框中的 value=1
    id_type_select = Select(driver.find_element(By.ID, "selIDType"))
    id_type_select.select_by_value("1")

    # 在 txtIDNumber 输入框中输入 idCode
    driver.find_element(By.ID, "txtIDNumber").send_keys(id_code)

    # 在 txtName 输入框中输入 user_name
    driver.find_element(By.ID, "txtName").send_keys(user_name)

    # 提示用户手动输入验证码，等待验证码输入框
    print("请输入验证码，点击 '登录' 按钮后，继续执行程序")
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "txtVerificationCode")))

    # 用户手动输入验证码，暂停执行直到输入完成
    while not driver.find_element(By.ID, "txtVerificationCode").get_attribute("value"):
        time.sleep(1)  # 每秒检查一次是否有输入

    # 获取用户输入的验证码
    verification_code = driver.find_element(By.ID, "txtVerificationCode").get_attribute("value")

    # 在 txtVerificationCode 输入框中输入用户输入的验证码
    driver.find_element(By.ID, "txtVerificationCode").send_keys(verification_code)

    # 检测直到登录按钮消失或不可点击
    try:
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "ibtnLogin")))
        print("登录按钮不可用或已消失，继续执行后续操作")
    except TimeoutException:
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "messager-error")))
            # 如果找到了 'messager-error' 元素，表示未报名
            print(f"{user_name} 未报名，写入文件并关闭浏览器")
            with open("data/cet_data.txt", "a", encoding="utf-8") as file:
                file.write(f"{user_name}\t未报名\n")
            driver.quit()  # 关闭浏览器
            return  # 退出函数，不继续执行后续操作
        except TimeoutException:
            print("未找到 'messager-error' 元素，继续执行后续操作")
            return


    # 等待页面加载
    time.sleep(1)

    # 读取 tbody 下的 tr 中第一个 td 的内容
    tbody = driver.find_element(By.ID, "tbody")
    first_td_content = tbody.find_element(By.XPATH, ".//tr[1]/td[1]").text

    # 将数据写入 cet_data.txt 文件
    with open("data/cet_data.txt", "a", encoding="utf-8") as file:
        file.write(f"{user_name}\t{first_td_content}\n")

    # 关闭浏览器
    driver.quit()




name_dict = {}
# 读取 Excel 文件
file_path = 'data/data.xlsx'
try:
    df = pd.read_excel(file_path)
    index = 0
    # 遍历每一行数据
    for _, row in df.iterrows():
        user_name = row['姓名']
        id_card = row['身份证号']
        index +=1

        name_dict[user_name] = id_card
    print(f'总数：{index}')
except Exception as e:
    print(f"Error reading the Excel file: {e}")


for key,value in name_dict.items():
    print(f"{key}\t{value}")
    get_info(key, value)
