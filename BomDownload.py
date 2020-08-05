#coding=utf-8
import os
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
import time
import json
import logging
import csv
from configparser import ConfigParser

logging.basicConfig(level = logging.INFO,format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
username = ''
password = ''
url = ''
driverPath = ''

def login(username, password, driver=None):
    driver.get(url)
    time.sleep(3)
    try:
        driver.find_element_by_xpath('//*[@id="inputform"]').send_keys(username)
        logger.info ('已输入email')
        driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
        logger.info ('已输入password')
        driver.find_element_by_xpath('//*[@id="rememail-checkbox"]/input').click()
        logger.info('已经记住email')
        driver.find_element_by_xpath('//*[@id="app-login-button"]').click()
        logger.info ('已点击Login')
        time.sleep(2)
        code = input('Please input Code:')
        inputCode = driver.find_element_by_name('form_tfa_code')
        inputCode.send_keys(code)
        logger.info('已经输入Code')
        checkBox = driver.find_element_by_name('form_tfa_trust_p')
        checkBox.click()
        logger.info('CheckBox Trust Click')
        Confirm = driver.find_element_by_name('submit')
        Confirm.click()
        logger.info ('已点击Submit')
    except:
        logger.info('Login Save Cookie Error!')
        return
    return driver

def login_test(username=None, password=None, driver=None):
    try:
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="inputform"]').send_keys(username)
        logger.info ('已输入email')
        driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
        logger.info ('已输入password')
        driver.find_element_by_xpath('//*[@id="rememail-checkbox"]/input').click()
        logger.info('已经记住email')
    except:
        logger.info('can not find element')
        return driver
    driver.find_element_by_xpath('//*[@id="app-login-button"]').click()
    logger.info('已点击Login')
    return driver

def get_cookies_from_network(driver=None):
    driver = login(username,password, driver=driver)
    # 获取cookies保存到本地并且返回此cookies
    cookies = driver.get_cookies()
    cookie_dict = {}
    logger.info(cookies)
    for cookie in cookies:
        if 'name' in cookie.keys() and 'value' in cookie.keys():
            cookie_dict[cookie['name']] = cookie['value']
    logger.info(cookie_dict)
    with open('cookies.txt', 'w') as f:
        # 保存cookies到本地
        f.write(json.dumps(cookies))
        logger.info("保存cookis成功.")
    return cookie_dict

def get_cookies_from_cache():
    cookies_dict = {}
    if os.path.exists('cookies.txt'):
        # 如果本地有cookies文件，则读取本地cookies，否则返回空
        with open('cookies.txt', 'r') as f:
            for i in json.loads(f.read()):
                if 'name' in i.keys() and 'value' in i.keys():
                        cookies_dict[i['name']] = i['value']
    else:
        return cookies_dict
    logger.info(cookies_dict)
    return cookies_dict

def get_cookies():
    cookie_dict = get_cookies_from_cache()
    if not cookie_dict:
        cookie_dict = get_cookies_from_network(driver=driver)
    return cookie_dict

def get_content(driver=None):
    # 获取cookies
    cookies = get_cookies()
    # 先访问一遍目标网站
    driver.get(url)
    # 删除当前的cookies
    driver.delete_all_cookies()
    # 添加cookies
    for k,v in cookies.items():
        driver.add_cookie({'name': k, 'value': v})
    time.sleep(1)
    logger.info('get url again')
    # 再次访问目标网站，模拟登录成功
    driver.get(url)
    driver=login_test(username=username,password=password,driver=driver)
    return driver

def SearchGo(driver=None,SearchNumber=None,Version=None,):
    # time.sleep(2)
    #输入料号，点击Go
    try:
        time.sleep(2)
        search_textfield = driver.find_element_by_xpath('//*[@id="WorldBar_Search"]/form/ul/li[4]/input[1]')
        search_textfield.send_keys(SearchNumber)
        time.sleep(0.5)
        SearchGo = driver.find_element_by_xpath('//*[@id="WorldBar_Search"]/form/ul/li[4]/input[2]')
        SearchGo.click()
        logger.info('Search Go Button Click OK')
    except:
        logger.info('input PN number fail')
        return None
    time.sleep(5)

    # 检查是否跳转到新页面
    # 页面一:带有title的选择页面 =====>
    # has data  ====>  select data
    # no data ===>
    # 页面二:不带有title的选择页面  选择料号版本  《====》 select version

    try:
        driver.find_element_by_xpath('//*[@id="fasttable-main"]/table/thead/tr')
        bodyvalue = driver.find_element_by_xpath('//*[@id="noData"]')
        if bodyvalue.text == 'No Items were found based on the search criteria.':
            logger.info('please check Bom PN number' + SearchNumber)
            return None
        else:
            hreflink = driver.find_element_by_xpath('//*[@id="fasttable-main"]/table/tbody/tr[1]/td[7]/a')
            hreflink.click()
            time.sleep(5)

        VersionSelect = driver.find_element_by_name("display_revision")
        Alloption = VersionSelect.find_elements_by_tag_name('option')
        version = ''
        for option in Alloption:
            version = option.text
            if version.find(Version) != -1:
                break
            if version == '':
                logger.info('please check PN number version  ' + SearchNumber)
                return None
        RevisionSelect = Select(driver.find_element_by_name('display_revision'))
        RevisionSelect.select_by_visible_text(version)
        logger.info('select revision success ')
        time.sleep(2)
        return driver

    except:
        VersionSelect = driver.find_element_by_name("display_revision")
        Alloption = VersionSelect.find_elements_by_tag_name('option')
        version = ''
        for option in Alloption:
            version = option.text
            if version.find(Version) != -1:
                break
            if version == '':
                logger.info('please check PN number version  '+SearchNumber)
                return None
        RevisionSelect = Select(driver.find_element_by_name('display_revision'))
        RevisionSelect.select_by_visible_text(version)
        logger.info('select revision success ')
        time.sleep(2)
        return driver

def ExportAction(driver=None,Type=None):
    time.sleep(3)
    # click one
    try:
        ActionElement = driver.find_element_by_xpath('//*[@id="actionPopupTrigger"]')
        ActionChains(driver).move_to_element(ActionElement).perform()
        time.sleep(2)
        Export = driver.find_element_by_xpath('//*[@id="action-export"]')
        Export.click()
        logger.info('Click Export OK')
    except:
        logger.info('Click Export Error')
        return None

    time.sleep(5)
    # click two
    try:
        select_mode = Select(driver.find_element_by_xpath('//*[@id="select_mode"]'))
        select_mode.select_by_visible_text(Type)
        ExportNow = driver.find_element_by_xpath('//*[@id="MultiPartAction_UseExistingSettings"]/tbody/tr[3]/td[2]/input[1]')
        ExportNow.click()
        logger.info('Click ExportNow OK')
    except:
        logger.info('Click ExportNow Error')
        return None
    time.sleep(2)
    return driver

def FlushWeb(driver=None):

    selectDay = Select(driver.find_element_by_xpath('//*[@id="InboxMessages"]/tbody/tr/td/div[1]/div/select[2]'))
    selectDay.select_by_value('3')
    time.sleep(3)
    selectDay = Select(driver.find_element_by_xpath('//*[@id="InboxMessages"]/tbody/tr/td/div[1]/div/select[2]'))
    selectDay.select_by_value('1')

def DownLoadBom(driver=None,Type=None):

    time.sleep(2)
    Dashboard = driver.find_element_by_xpath('//*[@id="wb-dashboard"]/a')
    Dashboard.click()
    time.sleep(3)

    FlushWeb(driver=driver)

    if Type.find('EE') != -1:
        time.sleep(30)
    time.sleep(15)

    status = driver.find_element_by_xpath('//*[@id="InboxMessagesList"]/tbody/tr[2]/td[6]').text
    while status != 'Completed':
        time.sleep(2)
        FlushWeb(driver=driver)
        time.sleep(3)
        status = driver.find_element_by_xpath('//*[@id="InboxMessagesList"]/tbody/tr[2]/td[6]').text

    download = driver.find_element_by_xpath('//*[@id="InboxMessagesList"]/tbody/tr[2]/td[8]/a')
    download.click()

def ParaseCsvFile(driver=None):
    with open(BomListPath,'r') as fp:
        reader = csv.DictReader(fp)
        for csvData in reader:
            test_val = SearchGo(driver=driver,SearchNumber=csvData['BOM PN'],Version=csvData['Version'])
            if test_val == None:
                with open('ErrorList.txt','a+') as f_err:
                    f_err.write(csvData['BOM PN']+','+csvData['Version']+','+csvData['Type'] + '\n')
            else:
                ExportAction(driver=driver,Type=csvData['Type'])
                DownLoadBom(driver=driver, Type=csvData['Type'])

if __name__ == '__main__':
    config = ConfigParser()
    config.read('config.ini')
    username = config.get("Setup", "username")
    password = config.get('Setup', 'password')
    url = config.get('Setup', 'url')
    driverPath = config.get('Setup', 'Path')
    BomListPath=config.get('Setup','BomListPath')
    driver = webdriver.Chrome(executable_path=driverPath)
    get_content(driver=driver)
    ParaseCsvFile(driver=driver)