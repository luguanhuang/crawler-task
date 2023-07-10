                                              # To extract personal data

from selenium.webdriver import Chrome                               # Chromedriver
from selenium.webdriver.common.by import By                         # To find the elements
from selenium.webdriver.support.ui import WebDriverWait             # To wait
from selenium.webdriver.support import expected_conditions as EC    # To wait
from selenium.webdriver.chrome.options import Options               # To resize the chrome window
from selenium.common.exceptions import TimeoutException             # To apply try/except
from selenium import webdriver
from threading import Thread
import time
import threading
import undetected_chromedriver as uc
# from utils import *
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from configparser import ConfigParser
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

import win32gui
import win32con
import win32com.client
import win32clipboard as w
import pymouse,pykeyboard
from pymouse import *
import os
import time
import win32api
import win32con
import pymouse,pykeyboard
from pymouse import *
from pykeyboard import PyKeyboard
from ctypes import *
from pykeyboard import PyKeyboard

import xlrd
import logging
from logging.config import fileConfig
from os import path
import logging
from configparser import ConfigParser
from datetime import datetime
from xlrd import xldate_as_datetime, xldate_as_tuple

from time import sleep              # To wait
from sys import exit                # To stop the code

class _ConfigReader:
    
    def __init__(self):
        self.config_object = ConfigParser()
        self.config_object.read("config.ini")

    def read_prop(self, section_header, prop_name):
        value = self.config_object.get(section_header,prop_name)
        logging.debug("{}:{}".format(prop_name, value))
        return value

    def read_bool_prop(self, section_header, prop_name):
        value = self.config_object.getboolean(section_header,prop_name)
        logging.debug("{}:{}".format(prop_name, value))
        return value

class AccountInfo:
    def __init__(self, begintme, endtme, account, passwd, center, VisaType, AppointmentType):
        self.begintme = begintme;
        self.endtme = endtme;
        self.account = account;
        self.passwd = passwd;
        self.center = center;
        self.VisaType = VisaType;
        self.AppointmentType = AppointmentType;

sheetname = 'Sheet1'
file_path = path.join(path.dirname(path.abspath(__file__)), '监控提醒.xls')
excel = xlrd.open_workbook(file_path)  # 打开excel文件
sheet = excel.sheet_by_index(0)  # 获取工作薄
name=sheet.name  #获取表的姓名
# print(name) 
 
nrows=sheet.nrows  #获取该表总行数
# print(nrows)  
 
ncols=sheet.ncols  #获取该表总列数
# print(ncols) 

arrdata = []
for i in range(1, nrows):
    begintme = xldate_as_datetime(sheet.cell_value(i,0), 0) #读取每一行的第二列的日期中年份给real_date_1_y 
    endtme = xldate_as_datetime(sheet.cell_value(i,1), 0) #读取每一行的第二列的日期中年份给real_date_1_y 
    account = AccountInfo(begintme, endtme, sheet.cell_value(i,2), sheet.cell_value(i,3),
                        sheet.cell_value(i,4), sheet.cell_value(i,5), sheet.cell_value(i,6))
    arrdata.append(account)
  
for i in range(len(arrdata)):
    print("begintime=%s endtime=%s account=%s passwd=%s center=%s VisaType=%s AppointmentType=%s\n"%
          (arrdata[i].begintme, arrdata[i].endtme, arrdata[i].account, arrdata[i].passwd, arrdata[i].center,
          arrdata[i].VisaType, arrdata[i].AppointmentType))
    
log_file_path = path.join(path.dirname(path.abspath(__file__)), 'logging.ini')
print("log_file_path=%s" % log_file_path)
# logging.config.fileConfig(log_file_path)
fileConfig(log_file_path)
logging = logging.getLogger("vfsbot");
_config_reader = _ConfigReader()
_interval = _config_reader.read_prop("DEFAULT", "interval")
xx = _config_reader.read_prop("DEFAULT", "x")
yy = _config_reader.read_prop("DEFAULT", "y")
qqnickname = _config_reader.read_prop("DEFAULT", "qqnickname")

logging.debug("Interval: {}".format(_interval))
logging.debug("x: {}".format(xx))
logging.debug("y: {}".format(yy))
logging.debug("qqnickname: {}".format(qqnickname))
vfs_login_url = _config_reader.read_prop("VFS", "vfs_login_url")
driver = path.join(path.dirname(path.abspath(__file__)), 'chromedriver.exe')
# exit(1)
# driver = "F:/code/assign/chinese/监控提示/chromedriver_win32/chromedriver.exe"
logging.debug(vfs_login_url)
logging.debug(driver)

def func():
    global xx,yy, qqnickname
    i=1
    msg="testttt222"
    # name='QQ'
    name=qqnickname
    # 获取窗口句柄
    handle = win32gui.FindWindow(None, name)
    print("handle=%d" % handle);
    pythoncom.CoInitialize()
    shell = win32com.client.Dispatch("WScript.Shell")

    shell.SendKeys('%')
    win32gui.SetForegroundWindow(handle)
    xy_pos = win32gui.GetWindowPlacement(handle) # 返回的是(0, 1, (-1, -1), (-1, -1), (712, 305, 1207, 775))
    #我们要取得是 (712, 305, 1207, 775) 窗口的四点坐标 所以
    x_y = win32gui.GetWindowPlacement(handle)[4] #(712, 305, 1207, 775

    # loginid = win32gui.GetWindowPlacement(handle)
    # x_y = win32gui.GetWindowPlacement(loginid)[4] #(712, 305, 1207, 775)

    # 获取窗口的顶点坐标 x， y
    x, y = x_y[0], x_y[1]
    # 获取窗口的高和宽 hight width
    hight = x_y[3] - x_y[1]
    width = x_y[2] - x_y[0]

    #604*92 客户的qq窗口

    # pos_x = x + 321
    # # pos_y = y + hight * 0.54
    # pos_y = y + 72

    pos_x = x + int(xx)
    # pos_y = y + hight * 0.54
    pos_y = y + int(yy)

    print("hight=%d width=%d pos_x=%d pos_y=%d x=%d y=%d" % (hight, width, pos_x, pos_y, x, y))

    time.sleep(1) # 本程序添加了大量的 time sleep延时函数 是为了程序可以正常运行
    #设置鼠标位置 位于账号的输入栏
    windll.user32.SetCursorPos(int(pos_x), int(pos_y))

    #输入完成后便可以登录了 回车键登录
    # time.sleep(0.1)
    # win32api.keybd_event(13, 0, 0, 0)
    # win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    time.sleep(0.2)
    cnt=1
    while 1:
        if cnt>=10:
            break
        cnt+=1
        time.sleep(1)
    print("will exit");

while 1:
    for i in range(len(arrdata)):
        
        # exit(1)
        service = Service(driver)
        options = webdriver.ChromeOptions()

        options.add_argument("disable-infobars")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-application-cache")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-dev-shm-usage")

        browser = webdriver.Chrome(service=service, options=options)
        browser.maximize_window()
        url1=vfs_login_url
        browser.get(url1)
        time.sleep(4)
        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-input-0"]'))
                )
        except:
            pass
        time.sleep(4)
        browser.find_element(By.XPATH, r'//*[@id="mat-input-0"]').click()
        # browser.find_element(By.XPATH, r"//input[@id='mat-input-0']").send_keys("dokcy17681@gmailqq.com")
        # browser.find_element(By.XPATH, r"//input[@id='mat-input-1']").send_keys("BCNOF56789z!")
        browser.find_element(By.XPATH, r"//input[@id='mat-input-0']").send_keys(arrdata[i].account)
        browser.find_element(By.XPATH, r"//input[@id='mat-input-1']").send_keys(arrdata[i].passwd)

        # while 1:#lgh0701
        #     try:
        #             element = WebDriverWait(browser, 0.1).until(EC.element_to_be_clickable((By.CLASS_NAME, 'ngx-overlay')))
        #     except:
        #         break

        # time.sleep(4)
        # while 1:
        #     try:
        #             element = WebDriverWait(browser, 0.1).until(
        #                 EC.element_to_be_clickable((By.CLASS_NAME, 'ngx-overlay'))
        #             )
        #     except:
        #         break
        # # browser.find_element(By.XPATH,r'//*[@id="onetrust-accept-btn-handler"]').click() lgh
        # time.sleep(2)
        # while 1:
        #     try:
        #             element = WebDriverWait(browser, 0.1).until(
        #                 EC.element_to_be_clickable((By.CLASS_NAME, 'ngx-overlay'))
        #             )
        #     except:
        #         break
        time.sleep(4)
        browser.find_element(By.XPATH, r"//button/span").click()
        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//section/div/div[2]/button/span"))
                )
        except:
            pass
        time.sleep(5)
        _new_booking_button = browser.find_element(By.XPATH, "//section/div/div[2]/button/span")
        _new_booking_button.click()
        # 

        # browser.find_element(By.XPATH, r"//button/span").click()
        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, r"//mat-form-field/div/div/div[3]"))
                )
        except:
            pass
        time.sleep(4)
        _visa_centre_dropdown = browser.find_element(By.XPATH, r"//mat-form-field/div/div/div[3]")
        _visa_centre_dropdown.click()
        # exit(1)
        # time.sleep(2)

        # try:
        # visa_centre = 'Italy Visa Application Center, Jinan'
        
        visa_centre = arrdata[i].center

        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//mat-option[starts-with(@id,'mat-option-')]/span[contains(text(), '{}')]".format(visa_centre)))
                )
        except:
            pass
        time.sleep(4)
        _visa_centre = browser.find_element(By.XPATH, "//mat-option[starts-with(@id,'mat-option-')]/span[contains(text(), '{}')]".format(visa_centre))
        print("VFS Centre: " + _visa_centre.text)
        browser.execute_script("arguments[0].click();", _visa_centre)
        # exit(1)
        # time.sleep(5)
        time.sleep(4)
        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, r"//div[@id='mat-select-value-3']"))
                )
        except:
            pass

        time.sleep(4)

        _category_dropdown = browser.find_element(By.XPATH, r"//div[@id='mat-select-value-3']")
        _category_dropdown.click()
        # time.sleep(3)
        time.sleep(4)
        # category = '意大利签证申请'
        category = arrdata[i].VisaType

        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, r"//mat-option[starts-with(@id,'mat-option-')]/span[contains(text(), '{}')]".format(category)))
                )
        except:
            pass

        _category = browser.find_element(By.XPATH, r"//mat-option[starts-with(@id,'mat-option-')]/span[contains(text(), '{}')]".format(category))
        print("Category: " + _category.text)
        browser.execute_script("arguments[0].click();", _category)
        # exit(1)
        # time.sleep(5)
        
        try:
            element = WebDriverWait(browser, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@id='mat-select-value-5']"))
            )
        except:
            pass
        time.sleep(4)
        _subcategory_dropdown = browser.find_element(By.XPATH, "//div[@id='mat-select-value-5']")
        browser.execute_script("arguments[0].click();", _subcategory_dropdown)
        
        time.sleep(4)
        # sub_category='SUNDAY VIP JINAN'
        sub_category=arrdata[i].AppointmentType

        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//mat-option[starts-with(@id,'mat-option-')]/span[contains(text(), '{}')]".format(sub_category)))
                )
        except:
            pass
        time.sleep(4)
        _subcategory = browser.find_element(By.XPATH, "//mat-option[starts-with(@id,'mat-option-')]/span[contains(text(), '{}')]".format(sub_category))

        browser.execute_script("arguments[0].click();", _subcategory)
        print("Sub-Cat: " + _subcategory.text)
        # time.sleep(5)

        try:
                element = WebDriverWait(browser, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[4]/div"))
                )
        except:
            pass

        time.sleep(4)
        _message = browser.find_element(By.XPATH, "//div[4]/div")
        print("Message: " + _message.text)

        if "目前没有可预约时段" in _message.text:
            print("not have");
        else:
            print("have");
            exit(0)
        browser.close()
        browser.quit()
    time.sleep(int(_interval))
        # exit(1)
    # t=Thread(target=func)
    # t.start()
    # while 1:
    #      time.sleep(1);
    # msg="testttt222"
    # # name='QQ'
    # name='test'
    # # 获取窗口句柄
    # handle = win32gui.FindWindow(None, name)
    # print("handle=%d" % handle);
    # win32gui.SetForegroundWindow(handle)
    # xy_pos = win32gui.GetWindowPlacement(handle) # 返回的是(0, 1, (-1, -1), (-1, -1), (712, 305, 1207, 775))
    # #我们要取得是 (712, 305, 1207, 775) 窗口的四点坐标 所以
    # x_y = win32gui.GetWindowPlacement(handle)[4] #(712, 305, 1207, 775

    # # loginid = win32gui.GetWindowPlacement(handle)
    # # x_y = win32gui.GetWindowPlacement(loginid)[4] #(712, 305, 1207, 775)

    # # 获取窗口的顶点坐标 x， y
    # x, y = x_y[0], x_y[1]
    # # 获取窗口的高和宽 hight width
    # hight = x_y[3] - x_y[1]
    # width = x_y[2] - x_y[0]

    # #604*92 客户的qq窗口

    # # pos_x = x + 321
    # # # pos_y = y + hight * 0.54
    # # pos_y = y + 72

    # pos_x = x + 424
    # # pos_y = y + hight * 0.54
    # pos_y = y + 73

    # print("hight=%d width=%d pos_x=%d pos_y=%d x=%d y=%d" % (hight, width, pos_x, pos_y, x, y))

    # time.sleep(1) # 本程序添加了大量的 time sleep延时函数 是为了程序可以正常运行
    # #设置鼠标位置 位于账号的输入栏
    # windll.user32.SetCursorPos(int(pos_x), int(pos_y))

    # #输入完成后便可以登录了 回车键登录
    # # time.sleep(0.1)
    # # win32api.keybd_event(13, 0, 0, 0)
    # # win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    # time.sleep(0.01)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    # time.sleep(0.2)
    # while 1:
    #     time.sleep(1)