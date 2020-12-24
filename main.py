from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

from selenium.webdriver.support.select import Select

if __name__ == '__main__':
    options = Options()
    options.add_argument("--disable-notifications")

    chrome = webdriver.Chrome('./chromedriver', chrome_options=options)

    chrome.get("https://www.lensmode.com/auth/login/redirectUrl/%252Fmypage%252Findex%252F/")

    chrome.find_element_by_xpath("/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[2]/td[2]/input[1]").send_keys('lensmamajp@gmail.com')
    chrome.find_element_by_xpath("/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[3]/td[2]/input").send_keys('kk20201201')
    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[4]/td/input[2]').click()

    chrome.get("https://www.lensmode.com/goods/index/gc/C1T/")

    dropdown = Select(chrome.find_element_by_id('PWR'))
    dropdown.select_by_index(2)

    dropdown1 = Select(chrome.find_element_by_id('NUM'))
    dropdown1.select_by_index('1')

    dropdown2 = Select(chrome.find_element_by_id('PWR2'))
    dropdown2.select_by_index(2)

    dropdown3 = Select(chrome.find_element_by_id('NUM2'))
    dropdown3.select_by_index('1')

    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article[1]/form/div/ul[2]/li/input[3]').click()
    options = Options()
    options.add_argument("--disable-notifications")
    time.sleep(3)
    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/div/article/ul/li/ul/li/a').click()
    #time.sleep(3)
    chrome.find_element_by_xpath(".//input[@type='radio' and @value='PMT031C']").click()
    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/form/div[2]/ul/li[2]/input[3]').click()
    time.sleep(3)
    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/article/form/span/input[1]').click()

    rwdata = chrome.find_elements_by_xpath("/html/body/div[3]/div[2]/article[1]/div[1]/table/tbody/tr[7]/td[3]")
    print(len(rwdata))
    for r in rwdata:
        print(r.text)

    chrome.find_element_by_xpath('/html/body/header/div[2]/ul[2]/li[2]/a').click()
    time.sleep(1)
    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/div/table/tbody/tr[2]/td[5]/a').click()
    time.sleep(1)
    chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/div/table/tbody/tr[2]/td[5]/a').click()

    #chrome.close()