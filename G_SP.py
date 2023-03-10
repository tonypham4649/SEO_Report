from _handyFunc.get_driver import *
from decouple import config
import urllib
import re
import pandas as pd
from statistics import mode
from urllib.parse import urlparse

chrome = WebDriver(proxy=config('proxy'))

def mostFrequent(list):
    try:
        if mode(list) != 'Info not found':
            return mode(list)
        else: 
            return next(item for item in list if item != 'Info not found')
    except:
        print('Most frequent item not found! Returning the first item!')
        return list[0]

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

def getInfo(ad_url):
    print('Getting ads info')
    lpURL_list = []
    name_list = []
    tel_list = []
    mail_list = []
    for y in range(len(ad_url)):
        chrome.driver.execute_script(script='window.open()')
        chrome.driver.switch_to.window(chrome.driver.window_handles[1])
        chrome.get(ad_url[y])

        html_text = chrome.driver.find_element(By.XPATH, '//body').text
        try:
            name = mostFrequent(re.findall(r"\株式会社\w+", html_text))
            name_list.append(name)
        except Exception:
            name_list.append('Info not found')
            pass

        try:
            tel = mostFrequent(re.findall("(\d{4}[-\.\s]??\d{3}[-\.\s]??\d{3}[-\.\s]??)", html_text))
            tel_list.append(tel)
        except Exception:
            tel_list.append('Info not found')
            pass

        try:
            mail = mostFrequent(re.findall("[\w\.-]+@[\w\.-]+\.\w+", html_text))
            mail_list.append(mail)
        except Exception:
            mail_list.append('Info not found')
            pass

        # get protocol + host name from ad_url
        parsed_url = urlparse(chrome.driver.current_url)
        lpURL_list.append('{uri.scheme}://{uri.netloc}/'.format(uri=parsed_url))
        
        chrome.driver.close()
        chrome.driver.switch_to.window(chrome.driver.window_handles[0])
    
    return name_list, tel_list, mail_list, lpURL_list

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
                    info = getInfo(data[1])

                    name_main.extend(info[0])
                    tel_main.extend(info[1])
                    mail_main.extend(info[2])
                    lpUrl_main.extend(info[3])

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