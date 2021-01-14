from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'

def PowerOnlyContactlens(counter, Url, CellList):

    for x in counter:
        chrome.get(Url)
        dropdown = Select(chrome.find_element_by_id('PWR'))
        dropdown.select_by_index(5)

        dropdown1 = Select(chrome.find_element_by_id('NUM'))
        dropdown1.select_by_index(x)
        # /html/body/div[3]/div[2]/article[1]/form/div/ul[2]/li/input[3]
        # chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/article[1]/form/div/ul[2]/li/input[3]').click()
        chrome.find_element_by_class_name('btn_cart241').click()
        time.sleep(3)
        chrome.find_element_by_xpath('/html/body/div[3]/div[2]/div/article/ul/li/ul/li/a').click()
        time.sleep(3)
        try:
            chrome.find_element_by_id('PMT031C').click()
        except:

            try:
                time.sleep(1)
                chrome.find_element_by_id('PMT031C').click()
            except:
                time.sleep(2)

                chrome.find_element_by_xpath(
                    '/html/body/div[11]/div/div/div/div/div/div[1]/div/p').click()
                chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/form/table[3]/tbody/tr[2]/td[5]/input').click()
            else:
                pass
        else:
            pass
        # chrome.find_element_by_xpath("/html/body/div[3]/div[2]/article/form/table[3]/tbody/tr[2]/td[5]/input").click()
        chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/form/div[2]/ul/li[2]/input[3]').click()

        chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/article/form/span/input[1]').click()
        #  pop-up close element
        #  /html/body/div[11]/div/div/div/div/div/div[1]/div/p

        # accumlate 8 boxes path

        # /html/body/div[3]/div[2]/article[1]/div[1]/table/tbody/tr[10]/td[3]
        # One time purchase 8 boxes path
        # /html/body/div[3]/div[2]/article[1]/div[1]/table/tbody/tr[6]/td[3]
        time.sleep(3)
        GetPrice = chrome.find_elements_by_xpath("/html/body/div[3]/div[2]/article[1]/div[1]/table/tbody/tr[6]/td[3]")
        #print(GetPrice)

        for r in GetPrice:


            UpdatePrice = r.text.strip(" \¥")
            print(UpdatePrice)

            creds = None
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

            # The ID of a spreadsheet.
            SAMPLE_SPREADSHEET_ID = '1taXktSkHSN3THYRP6m4wNeaP7vLrazemh132P_xLhDo'
            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()

            # result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            #                             range="10款1day格價!A1:BC39").execute()
            # get googlesheet cell value
            # values = result.get('values', [])
            print(CellList[x])
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="10款1day格價!" + CellList[x],
                                  valueInputOption="USER_ENTERED", body={"values": [[UpdatePrice]]}).execute()
            time.sleep(3)

        chrome.find_element_by_xpath('/html/body/header/div[2]/ul[2]/li[2]/a').click()
        time.sleep(1)
        try:
            chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/div/table/tbody/tr[2]/td[5]/a').click()
        except:
            try:
                time.sleep(1)
                chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/div/table/tbody/tr[2]/td[5]/a').click()
            except:
                time.sleep(1)

                chrome.find_element_by_xpath('/html/body/div[11]/div/div/div/div/div/div[1]/div/p').click()
                chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/div/table/tbody/tr[2]/td[5]/a').click()
            else:
                pass
        else:
            pass


def PowerBCContactlens(counter, Url, CellList):

    for x in counter:

        chrome.get(Url)
        dropdown = Select(chrome.find_element_by_id('BCDIA'))
        dropdown.select_by_index(1)

        dropdown = Select(chrome.find_element_by_id('PWR'))
        dropdown.select_by_index(5)

        dropdown1 = Select(chrome.find_element_by_id('NUM'))
        dropdown1.select_by_index(x)

        #chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/article[1]/form/div/ul[2]/li/input[3]').click()
        chrome.find_element_by_class_name('btn_cart241').click()

        try:
            chrome.find_element_by_class_name('btn_210').click()
        except:
            try:
                time.sleep(1)
                chrome.find_element_by_class_name('btn_210').click()
            except:
                # chrome.find_element_by_xpath('/html/body/div[11]/div/div/div/div/div/div[1]/div/p').click()
                chrome.find_element_by_xpath('/html/body/div[3]/div[2]/div/article/ul/li/ul/li/a').click()
            else:
                pass
        else:
            pass
        time.sleep(3)
        button = chrome.find_element_by_xpath("/html/body/div[3]/div[2]/article/form/table[3]/tbody/tr[2]/td[5]/input")
        chrome.execute_script("arguments[0].click();", button)

        chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/form/div[2]/ul/li[2]/input[3]').click()

        chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/article/form/span/input[1]').click()

        time.sleep(3)
        GetPrice = chrome.find_elements_by_xpath("/html/body/div[3]/div[2]/article[1]/div[1]/table/tbody/tr[6]/td[3]")
        #print(GetPrice)

        for r in GetPrice:
            UpdatePrice = r.text.strip(" \¥")
            print(UpdatePrice)

            # CellList = ["", "R3", "S3", "", "T3", "", "U3", "", "V3"]

            creds = None
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

            # The ID of a spreadsheet.
            SAMPLE_SPREADSHEET_ID = '1taXktSkHSN3THYRP6m4wNeaP7vLrazemh132P_xLhDo'
            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()

            # result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            #                             range="10款1day格價!A1:BC39").execute()
            # get googlesheet cell value
            # values = result.get('values', [])
            print(CellList[x])
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="10款1day格價!" + CellList[x],
                                  valueInputOption="USER_ENTERED", body={"values": [[UpdatePrice]]}).execute()
            time.sleep(3)

        chrome.find_element_by_xpath('/html/body/header/div[2]/ul[2]/li[2]/a').click()
        time.sleep(1)
        chrome.find_element_by_xpath('/html/body/div[3]/div[2]/article/div/table/tbody/tr[2]/td[5]/a').click()

def PowerBCColorContactlens(counter, Url, CellList):

    for x in counter:

        chrome.get(Url[x])
        if x == 1:
            dropdown = Select(chrome.find_element_by_name('selRightEyeBC'))
            dropdown.select_by_index(1)

            dropdown = Select(chrome.find_element_by_name('selRightEyeColor'))
            dropdown.select_by_index(4)

            dropdown = Select(chrome.find_element_by_id('selRightEyePWR'))
            dropdown.select_by_index(2)
            chrome.find_element_by_name('txtAmount').clear()
            chrome.find_element_by_name('txtAmount').send_keys(x)

            chrome.find_element_by_id('cartbtn').click()
        elif x > 1:
            dropdown = Select(chrome.find_element_by_name('selRightEyeBC'))
            dropdown.select_by_index(1)

            dropdown = Select(chrome.find_element_by_name('selRightEyeColor'))
            dropdown.select_by_index(4)

            dropdown = Select(chrome.find_element_by_id('selRightEyePWR'))
            dropdown.select_by_index(2)

            dropdown = Select(chrome.find_element_by_name('selLeftEyeBC'))
            dropdown.select_by_index(1)

            dropdown = Select(chrome.find_element_by_name('selLeftEyeColor'))
            dropdown.select_by_index(4)

            dropdown = Select(chrome.find_element_by_id('selLeftEyePWR'))
            dropdown.select_by_index(2)

            chrome.find_element_by_id('cartbtn').click()

        time.sleep(1)
        chrome.find_element_by_xpath('/html/body/div[1]/div[1]/article/section[2]/section/form[2]/p/input').click()
        time.sleep(1)
        radiobtn = chrome.find_element_by_xpath('/html/body/div[1]/div/article/section[2]/section/form[2]/div/table[3]/tbody/tr[3]/td/div/ins')
        radiobtn.click()

        chrome.find_element_by_xpath("/html/body/div[1]/div/article/section[2]/section/form[2]/div/p[3]/input").click()


        GetPrice = chrome.find_elements_by_xpath("/html/body/div[1]/div[1]/article/section[2]/section/table/tfoot/tr/td/table/tbody/tr[4]/td/strong")
        #print(GetPrice)

        for r in GetPrice:
            UpdatePrice = r.text.strip(" \円")
            print(UpdatePrice)

            # CellList = ["", "R3", "S3", "", "T3", "", "U3", "", "V3"]

            creds = None
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

            # The ID of a spreadsheet.
            SAMPLE_SPREADSHEET_ID = '1taXktSkHSN3THYRP6m4wNeaP7vLrazemh132P_xLhDo'
            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()

            # result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            #                             range="10款1day格價!A1:BC39").execute()
            # get googlesheet cell value
            # values = result.get('values', [])
            print(CellList[x])
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="10款1day格價!" + CellList[x],
                                  valueInputOption="USER_ENTERED", body={"values": [[UpdatePrice]]}).execute()
            #time.sleep(3)

        chrome.find_element_by_xpath('/html/body/div[1]/header/div[2]/nav/ul/li[4]/a').click()
        time.sleep(1)
        chrome.find_element_by_xpath('/html/body/div[1]/div[1]/article/section[2]/section/form[1]/table/tbody/tr/td[5]/div/p/input').click()

def LaboPowerBCContactlens(counter, Url, CellList):

    for x in counter:

        chrome.get(Url[x])
        if x == 1:
            dropdown = Select(chrome.find_element_by_xpath('/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div/table/tbody/tr[1]/td/select'))
            dropdown.select_by_index(1)

            dropdown = Select(chrome.find_element_by_id('selRightEyePWR'))
            dropdown.select_by_index(2)

            chrome.find_element_by_id('cartbtn').click()
        elif x > 1:
            dropdown = Select(chrome.find_element_by_xpath('/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div[1]/table/tbody/tr[2]/td/select'))
            dropdown.select_by_index(1)

            dropdown = Select(chrome.find_element_by_id('selRightEyePWR'))
            dropdown.select_by_index(2)

            dropdown = Select(chrome.find_element_by_xpath(
                '/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[2]/td/select'))
            dropdown.select_by_index(1)

            dropdown = Select(chrome.find_element_by_id('selLeftEyePWR'))
            dropdown.select_by_index(2)


            chrome.find_element_by_id('cartbtn').click()

        time.sleep(3)
        chrome.find_element_by_xpath('/html/body/div[1]/div[1]/article/section[2]/section/form[2]/p/input').click()
        time.sleep(3)
        radiobtn = chrome.find_element_by_xpath(
            '/html/body/div[1]/div/article/section[2]/section/form[2]/div/table[3]/tbody/tr[3]/td/div/ins')
        radiobtn.click()
        chrome.find_element_by_xpath("/html/body/div[1]/div/article/section[2]/section/form[2]/div/p[3]/input").click()


        GetPrice = chrome.find_elements_by_xpath("/html/body/div[1]/div[1]/article/section[2]/section/table/tfoot/tr/td/table/tbody/tr[4]/td/strong")
        #print(GetPrice)

        for r in GetPrice:
            UpdatePrice = r.text.strip(" \円")
            print(UpdatePrice)

            # CellList = ["", "R3", "S3", "", "T3", "", "U3", "", "V3"]

            creds = None
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

            # The ID of a spreadsheet.
            SAMPLE_SPREADSHEET_ID = '1taXktSkHSN3THYRP6m4wNeaP7vLrazemh132P_xLhDo'
            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()

            # result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            #                             range="10款1day格價!A1:BC39").execute()
            # get googlesheet cell value
            # values = result.get('values', [])
            print(CellList[x])
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="10款1day格價!" + CellList[x],
                                  valueInputOption="USER_ENTERED", body={"values": [[UpdatePrice]]}).execute()
            #time.sleep(3)

        chrome.find_element_by_xpath('/html/body/div[1]/header/div[2]/nav/ul/li[4]/a').click()
        time.sleep(1)
        chrome.find_element_by_xpath('/html/body/div[1]/div[1]/article/section[2]/section/form[1]/table/tbody/tr/td[5]/div/p/input').click()


def LaboPowerBCQtyContactlens(counter, Url, CellList):

    for x in counter:

        chrome.get(Url[x])
        time.sleep(1)
        if x == 1:
            dropdown = Select(chrome.find_element_by_name('selRightEyeBC'))
            dropdown.select_by_index(1)

            try:
                dropdown = Select(WebDriverWait(chrome, 20).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div/table/tbody/tr[2]/td/select']"))))
            except:
                dropdown = Select(chrome.find_element_by_name('selRightEyePWR'))
                dropdown.select_by_index(2)
            else:
                # dropdown = Select(WebDriverWait(chrome, 20).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div/table/tbody/tr[2]/td/select']"))))
                dropdown.select_by_index(2)

            try:
                dropdown = Select(chrome.find_element_by_name('txtAmount'))
            except:
                chrome.find_element_by_name('txtAmount').clear()
                chrome.find_element_by_name('txtAmount').send_keys("1")

            else:
                dropdown.select_by_index(1)


            chrome.find_element_by_id('cartbtn').click()
        elif x > 1:
            dropdown = Select(chrome.find_element_by_name('selRightEyeBC'))
            dropdown.select_by_index(1)

            try:
                dropdown = Select(WebDriverWait(chrome, 20).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div/table/tbody/tr[2]/td/select']"))))
            except:
                dropdown = Select(chrome.find_element_by_name('selRightEyePWR'))
                dropdown.select_by_index(2)
            else:
                dropdown.select_by_index(2)

            dropdown = Select(chrome.find_element_by_name('selLeftEyeBC'))
            dropdown.select_by_index(1)

            try:
                dropdown = Select(chrome.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/form/article/section/section[1]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[3]/td/select'))
            except:
                dropdown = Select(chrome.find_element_by_name('selLeftEyePWR'))
                dropdown.select_by_index(2)
            else:
                dropdown.select_by_index(2)

            chrome.find_element_by_id('cartbtn').click()

        time.sleep(2)
        chrome.find_element_by_xpath('/html/body/div[1]/div[1]/article/section[2]/section/form[2]/p/input').click()
        time.sleep(1)

        radiobtn = chrome.find_element_by_xpath(
            '/html/body/div[1]/div/article/section[2]/section/form[2]/div/table[3]/tbody/tr[3]/td/div/ins')
        radiobtn.click()
        chrome.find_element_by_xpath("/html/body/div[1]/div/article/section[2]/section/form[2]/div/p[3]/input").click()

        GetPrice = chrome.find_elements_by_xpath(
            "/html/body/div[1]/div[1]/article/section[2]/section/table/tfoot/tr/td/table/tbody/tr[4]/td/strong")
        # print(GetPrice)

        for r in GetPrice:
            UpdatePrice = r.text.strip(" \円")
            print(UpdatePrice)

            creds = None
            creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

            # The ID of a spreadsheet.
            SAMPLE_SPREADSHEET_ID = '1taXktSkHSN3THYRP6m4wNeaP7vLrazemh132P_xLhDo'
            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()

            # result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            #                             range="10款1day格價!A1:BC39").execute()
            # get googlesheet cell value
            # values = result.get('values', [])
            print(CellList[x])
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="10款1day格價!" + CellList[x],
                                  valueInputOption="USER_ENTERED", body={"values": [[UpdatePrice]]}).execute()
            # time.sleep(3)
        time.sleep(1)
        chrome.find_element_by_xpath('/html/body/div[1]/header/div[2]/nav/ul/li[4]/a').click()
        time.sleep(1)
        chrome.find_element_by_xpath(
            '/html/body/div[1]/div[1]/article/section[2]/section/form[1]/table/tbody/tr/td[5]/div/p/input').click()


if __name__ == '__main__':

    # WEBSITE: LENSMODE
    options = Options()
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")

    chrome: WebDriver = webdriver.Chrome('./chromedriver', options=options)
    #
    # chrome.get("https://www.lensmode.com/auth/login/redirectUrl/%252Fmypage%252Findex%252F/")
    #
    # chrome.find_element_by_xpath("/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[2]/td[2]/input[1]").send_keys('lensmamajp@gmail.com')
    # chrome.find_element_by_xpath("/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[3]/td[2]/input").send_keys('kk20201201')
    # chrome.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[4]/td/input[2]').click()
    #
    # # Dailies Total 1
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/C1T/"
    # CellList = ["","C3","D3","","E3","","F3","","G3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Moist
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/J1M/"
    # CellList = ["", "R3", "S3", "", "T3", "", "U3", "", "V3"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Moist 90 PACK
    # counter = [1, 2, 4]
    # Url = "https://www.lensmode.com/goods/index/gc/J1M90/"
    # CellList = ["", "W3", "X3", "", "Y3"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Trueye
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/J1T/"
    # CellList = ["", "M3", "N3", "", "O3", "", "P3", "", "Q3"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Oasys
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/JOS1/"
    # CellList = ["", "H3", "I3", "", "J3", "", "K3", "", "L3"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # Myday
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/CM1/"
    # CellList = ["", "AE3", "AF3", "", "AG3", "", "AH3", "", "AI3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # Proclear 1 Day
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/CP1/"
    # CellList = ["", "AJ3", "AK3", "", "AL3", "", "AM3", "", "AN3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # Biomedics 1 Day
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/O1N/"
    # CellList = ["", "AO3", "AP3", "", "AQ3", "", "AR3", "", "AS3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # 1 Day Biotrue
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/B1T/"
    # CellList = ["", "AY3", "AZ3", "", "BA3", "", "BB3", "", "BC3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # Medalist 1 day plus
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/B1N/"
    # CellList = ["", "AT3", "AU3", "", "AV3", "", "AW3", "", "AX3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Define Vivid
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lensmode.com/goods/index/gc/J1MV/"
    # CellList = ["", "Z3", "AA3", "", "AB3", "", "AC3", "", "AD3"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # WEBSITE: Lenszero
    #
    # chrome.get("https://www.lenszero.com/auth/login/redirectUrl/%252Fmypage%252Findex%252F/")
    #
    # chrome.find_element_by_xpath(
    #     "/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[2]/td[2]/input[1]").send_keys(
    #     'lensmamajp@gmail.com')
    # chrome.find_element_by_xpath(
    #     "/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[3]/td[2]/input").send_keys('kk20201201')
    # chrome.find_element_by_xpath(
    #     '/html/body/div[3]/div[2]/div[1]/article/form[1]/table/tbody/tr[4]/td/input[2]').click()
    #
    # # 1 Day Acuvue Define RC
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/J1MC/"
    # CellList = ["", "Z5", "AA5", "", "AB5", "", "AC5", "", "AD5"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    #
    # # Dailies Total 1
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/C1T/"
    # CellList = ["","C5","D5","","E5","","F5","","G5"]
    # PowerOnlyContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Moist
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/J1M/"
    # CellList = ["", "R5", "S5", "", "T5", "", "U5", "", "V5"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Moist 90 PACK
    # counter = [1, 2, 4]
    # Url = "https://www.lenszero.com/goods/index/gc/J1M90/"
    # CellList = ["", "W5", "X5", "", "Y5"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Trueye
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/J1T/"
    # CellList = ["", "M5", "N5", "", "O5", "", "P5", "", "Q5"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # 1 Day Acuvue Oasys
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/JOS1/"
    # CellList = ["", "H5", "I5", "", "J5", "", "K5", "", "L5"]
    # PowerBCContactlens(counter, Url, CellList)
    #
    # # Myday 無出售
    # """
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/CM1/"
    # CellList = ["", "AE5", "AF5", "", "AG5", "", "AH5", "", "AI5"]
    # PowerOnlyContactlens(counter, Url, CellList)
    # """
    #
    # # Proclear 1 Day
    # """
    # counter = [1, 2, 4, 6, 8]
    # Url = "https://www.lenszero.com/goods/index/gc/CP1/"
    # CellList = ["", "AJ5", "AK5", "", "AL5", "", "AM5", "", "AN5"]
    # PowerOnlyContactlens(counter, Url, CellList)
    # """

    #WEBSITE: Lens-labo

    chrome.get("https://www.lens-labo.com/login")

    chrome.find_element_by_xpath(
        "/html/body/div[1]/div[1]/article/div[2]/div/form/div/section[1]/table/tbody/tr[1]/td[2]/input").send_keys(
        'lensmamajp@gmail.com')
    chrome.find_element_by_xpath(
        "/html/body/div[1]/div[1]/article/div[2]/div/form/div/section[1]/table/tbody/tr[2]/td[2]/input").send_keys('kk20201201')
    chrome.find_element_by_xpath(
        '/html/body/div[1]/div[1]/article/div[2]/div/form/div/section[1]/p[2]/input').click()

    # 1 Day Acuvue Define RC
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0005-1", "https://www.lens-labo.com/item/detail?itemcd=L0005-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0005-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0005-6", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0005-8"]
    CellList = ["", "Z7", "AA7", "", "AB7", "", "AC7", "", "AD7"]
    PowerBCColorContactlens(counter, Url, CellList)

    # Dailies Total 1
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0039-1", "https://www.lens-labo.com/item/detail?itemcd=L0039-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0039-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0039-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0039-8"]
    CellList = ["","C7","D7","","E7","","F7","","G7"]
    LaboPowerBCContactlens(counter, Url, CellList)

    # 1 Day Acuvue Moist
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0001-1", "https://www.lens-labo.com/item/detail?itemcd=L0001-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0001-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0001-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0001-8"]
    CellList = ["", "R7", "S7", "", "T7", "", "U7", "", "V7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)

    # 1 Day Acuvue Moist 90 PACK
    counter = [1, 2, 4]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0120-1", "https://www.lens-labo.com/item/detail?itemcd=L0120-2", "", "https://www.lens-labo.com/item/detail?itemcd=L0120-4"]
    CellList = ["", "W7", "X7", "", "Y7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)
    #
    # 1 Day Acuvue Trueye
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0002-1", "https://www.lens-labo.com/item/detail?itemcd=L0002-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0002-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0002-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0002-8"]
    CellList = ["", "M7", "N7", "", "O7", "", "P7", "", "Q7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)
    #
    # 1 Day Acuvue Oasys
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0004-1", "https://www.lens-labo.com/item/detail?itemcd=L0004-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0004-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0004-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0004-8"]
    CellList = ["", "H7", "I7", "", "J7", "", "K7", "", "L7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)

    # Myday 無出售

    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0014-1", "https://www.lens-labo.com/item/detail?itemcd=L0014-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0014-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0014-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0014-8"]
    CellList = ["", "AE7", "AF7", "", "AG7", "", "AH7", "", "AI7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)


    # Proclear 1 Day

    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0012-1", "https://www.lens-labo.com/item/detail?itemcd=L0012-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0012-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0012-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0012-8"]
    CellList = ["", "AJ7", "AK7", "", "AL7", "", "AM7", "", "AN7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)

    # 1 Day Biotrue
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0024-1", "https://www.lens-labo.com/item/detail?itemcd=L0024-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0024-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0024-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0024-8"]
    CellList = ["", "AY7", "AZ7", "", "BA7", "", "BB7", "", "BC7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)

    # Medalist 1 day plus
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0003-1", "https://www.lens-labo.com/item/detail?itemcd=L0003-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0003-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0003-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0003-8"]
    CellList = ["", "AT7", "AU7", "", "AV7", "", "AW7", "", "AX7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)

    # Biomedics 1 Day
    counter = [1, 2, 4, 6, 8]
    Url = ["", "https://www.lens-labo.com/item/detail?itemcd=L0013-1", "https://www.lens-labo.com/item/detail?itemcd=L0013-2", "",
           "https://www.lens-labo.com/item/detail?itemcd=L0013-4", "", "https://www.lens-labo.com/item/detail?itemcd=L0013-6",
           "", "https://www.lens-labo.com/item/detail?itemcd=L0013-8"]
    CellList = ["", "AO7", "AP7", "", "AQ7", "", "AR7", "", "AS7"]
    LaboPowerBCQtyContactlens(counter, Url, CellList)



