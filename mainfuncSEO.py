import re
from statistics import mode
from urllib.parse import urlparse
from _handyFunc.get_driver import *

def mostFrequent(list):
    try:
        if mode(list) != 'Info not found':
            return mode(list)
        else: 
            return next(item for item in list if item != 'Info not found')
    except:
        print('Most frequent item not found! Returning the first item!')
        return list[0]

def getInfo(chrome, ad_url):
    print('Getting ads info')

    lpURL_list = []
    name_list = []
    tel_list = []
    mail_list = []
    for y in range(len(ad_url)):
        chrome.driver.execute_script(script='window.open()')
        chrome.driver.switch_to.window(chrome.driver.window_handles[1])
        chrome.get(ad_url[y])
        
        chrome.driver.implicitly_wait(5)
        chrome.driver.execute_script("window.stop();")
        
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
    
    return {'name_list':name_list,
            'tel_list':tel_list, 
            'mail_list':mail_list, 
            'lpURL_list':lpURL_list}