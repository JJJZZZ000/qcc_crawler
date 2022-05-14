import os
from datetime import datetime
from selenium import webdriver
from lxml import etree
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains
import re
import xlwt
import pandas as pd
from ChromeUpdate import check_update_chromedriver
import warnings
warnings.filterwarnings("ignore")

def init():
    dir_res = './Result/'
    dir_chrome = '/chrome/'
    if not os.path.exists(dir_res): 
        os.makedirs(dir_res)
    if not os.path.exists(dir_chrome): 
        os.makedirs(dir_chrome)
    
    check_update_chromedriver('.\chrome\\')
    
    option = webdriver.ChromeOptions()
    # 此步骤很重要，设置为开发者模式，防止被各大网站识别出来使用了Selenium
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    option.add_argument("--disable-blink-features")
    option.add_argument("--disable-blink-features=AutomationControlled")
    option.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})

    ## 实例化浏览器对象 传入浏览器的驱动程序
    bro = webdriver.Chrome(executable_path='./chrome/chromedriver', options=option)
    return bro

def login(bro: webdriver):
    bro.get('https://www.qcc.com/?msclkid=ad9ac512d10211ec92d04dea2031dd27')
    while(True): 
        try:
            log_in = find_ele(bro.find_elements_by_xpath("//span"), "登录 | 注册")
            if(log_in == None): break 
            log_in.click()
            while(True):
                try:
                    bro.find_element_by_xpath("//*[@id='loginModal']")
                except:
                    break
        except:
            pass

def find_ele(ele, str):
    for e in ele:
        if(e.text==str): return e
    return None

def search(bro, company_name, xpaths) -> dict:
    input = bro.find_element_by_xpath('//*[@id="searchKey"]')
    input.clear()
    input.send_keys(company_name)
    search_btns = bro.find_elements_by_tag_name("button")
    search_btn = find_ele(search_btns, "查一下")
    bro.execute_script("arguments[0].click();", search_btn)

    tables = bro.find_elements_by_tag_name("table")
    if(len(tables) == 0): pass
    result_ls = bro.find_elements_by_tag_name("table")[-1]
    link = result_ls.find_element_by_xpath("./tr[1]//a")
    link.click()

    new_window = bro.window_handles[-1]
    bro.close()
    bro.switch_to.window(new_window)

    parseResult = {}
    for item in xpaths:
        try:
            name = str(item[0])
            attr = str(item[1]).lower()
            xpath = str(item[2])
            res = bro.find_element_by_xpath(xpath)
            if(attr.lower() == 'text' or attr.lower() == 'txt'):
                parseResult[name] = res.text
            else:
                parseResult[name] = res.get_attribute(attr)
        except:
            bro.get("https://www.qcc.com/")
            pass
    parseResult["Search input"] = company_name
    return parseResult



companies = []
xpaths = []

## read the company & xpath in the txt
with open('./Company.txt','r',encoding = "utf-8")as f:
    c = f.read()
    companies = re.split("\n|\t|,|，", c)

with open('./Xpath.txt','r',encoding = "utf-8")as f:
    lines = f.readlines()
    for line in lines:
        line = line.replace('\n','')
        entry = re.split("\t|:|：", line)
        if(len(entry) < 3): continue
        xpaths.append([entry[0].replace(" ",""), entry[1].replace(" ",""), entry[2].replace(" ","")])




bro = init()
login(bro)


# check if the verification happens
try:
    input = bro.find_element_by_xpath('//*[@id="searchKey"]')
except:
    os.system("pause")

parseResults = []
for company in companies:
    try:
        parseResult = search(bro, company, xpaths)
        parseResults.append(parseResult)
    except:
        pass
print("Crawl complete!")


col = []
for item in xpaths:
    col.append(item[0])
dataf = pd.DataFrame(columns=col)


for parseResult in parseResults:
    dataf = dataf.append(parseResult, ignore_index=True)

now = datetime.now()
filename = now.strftime("%Y%m%d%H%M%S")+".xlsx"

dataf.to_excel(dir+filename)
bro.quit() 
print("Write complete!")