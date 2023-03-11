from _handyFunc.get_driver import *
from mainfuncSEO import getInfo
from decouple import config
import pandas as pd
import urllib

chrome = WebDriver(proxy=config('proxy'))

def getElement(title_xpath, url_xpath, des_xpath, page_xpath):
    print('Getting the elements')
    title_list = []
    ad_url = []
    des_list = []
    page_list = []
    counter = 0

    for tptitle in chrome.get_elements_by_xpath(title_xpath):
        title_list.append(tptitle.text)

    for url in chrome.get_elements_by_xpath(url_xpath):
        counter += 1
        ad_url.append(url.get_attribute('href'))

    for des in chrome.get_elements_by_xpath(des_xpath):
        des_list.append(des.text)

    for x in range(counter):
        x += 1
        try:
            page_list.append(chrome.get_element_by_xpath(page_xpath).text.replace('ページ ',''))
        except Exception:
            page_list.append('1')

    # nL = len(name_list)
    tL = len(title_list)
    return title_list, ad_url, des_list, page_list, tL, counter

def mainRun(kw_list):
    chrome.get_chrome_sp()
    url = 'https://www.google.co.jp/search?q='

    # CSS xpath
    title_xpath = "//*[@id='main']//*[@aria-level='3']"
    url_xpath = "//*[@class='uEierd']//*[@role='presentation']"
    des_xpath = "//*[@class='MUxGbd yDYNvb']/div"

    # if element not found -> page 1
    page_xpath = "//*[@class='SAez4c']"
    next_xpath = "//*[@id='main']/footer//*[@aria-label='次のページ']"

    keyword_main = []
    name_main = []
    title_main = []
    des_main = []
    page_main = []
    tel_main = []
    mail_main = []
    lpUrl_main = []

    for kw in kw_list:
        print(f'Target keyword: {kw}')
        query = urllib.parse.quote_plus(kw)
        chrome.driver.get(url + query)
        page = 1
        while page <= 3:
            print(f'Page: {page}')            
            tries = 0
            while tries <= 4:
                tries += 1
                print(f'Try counter - {tries}')
                
                data = getElement(title_xpath, url_xpath, des_xpath, page_xpath)
                if data[4] == data[5]:
                    for x in range(data[5]):
                        keyword_main.append(kw)
                        x += 1

                    print('Putting in main lists')
                    info = getInfo(chrome, data[1])

                    name_main.extend(info['name_list'])
                    tel_main.extend(info['tel_list'])
                    mail_main.extend(info['mail_list'])
                    lpUrl_main.extend(info['lpURL_list'])

                    title_main.extend(data[0])
                    des_main.extend(data[2])
                    page_main.extend(data[3])

                    page += 1
                    chrome.click_element_by_xpath(next_xpath)
                    break
            else: 
                page += 1
                chrome.click_element_by_xpath(next_xpath)
                break
    
    chrome.driver.close()
    
    dataPd = pd.DataFrame(
        {'keyword': keyword_main,
         'matchType': '完全一致',
         '媒体': 'Google',
         'device': 'SP',
         'name': name_main,
         'title': title_main,
         'description': des_main,
         'phoneNumber': tel_main,
         'email': mail_main,
         'url': lpUrl_main,
         'page': page_main
        }
    )
    return dataPd