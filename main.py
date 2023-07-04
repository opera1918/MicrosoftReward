import re
import os
import json
import time
import requests
import traceback
from sys import argv
from datetime import date, datetime, time as _time
from colorama import Fore, Style
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from terminate import kill_process
from clear_data import clear_chrome_data
from transactions import logOrders

BASE_URL = 'https://rewards.bing.com/'
ABS_EXTENSION  = "chrome-extension://ipbgaooglppjombmbgebgmaehjkfabme/popup.html"

def bingSignIn(driver):
    hold = 0
    while True:
        try:
            if "\"IsAuthenticated\":true" not in driver.page_source:
                sign_in_button = driver.find_element(By.ID, 'id_l')
                sign_in_button.click()
                time.sleep(1)
                hold += 1
            else:
                driver.close()
                break
            if hold % 3 == 0:
                driver.refresh()
        except:
            time.sleep(0.5)

def progress(driver, wait, device, level):
    try:
        ptrn = re.compile(r"<b>(\d{1,3})<\/b> \/ (\d{2,3})<\/p>", re.MULTILINE)
        driver.get(BASE_URL+"pointsbreakdown")
        if wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="modal-host"]/div[2]'))):
            instancePoint = []
            for i in driver.find_elements(By.CLASS_NAME, 'pointsDetail'):
                result = ptrn.findall(i.get_attribute('innerHTML'))
                if len(result) > 0:
                    instancePoint.append(result[0])
            if (level == 1 and device == 'Desktop') or (level == 2 and device == 'Desktop'):
                return instancePoint[0]
            elif (level == 2 and device == 'Mobile'):
                return instancePoint[1]
            else:
                return instancePoint[:2]
    except Exception as e:
        pass

def startSearch(driver, wait, exten, device, level):
    driver.execute_script(f"window.open('data:;')")
    driver.switch_to.window(driver.window_handles[1])
    driver.get(exten)
    spoofing = Select(driver.find_element(By.ID, 'platform-spoofing'))
    spoofing.select_by_visible_text(f'{device} only')
    wait.until(EC.element_to_be_clickable((By.ID, 'search'))).click()
    driver.switch_to.window(driver.window_handles[0])

    counter, prev_p1 = 0, None
    while True:
        time.sleep(8)
        p1, p2 = progress(driver, wait, device, level)
        if p1 == p2:
            break
        elif counter == 5:
            raise ValueError("Progress has not changed in the last 40 seconds.")
        elif p1 == prev_p1:
            counter += 1
        else:
            counter = 0
            prev_p1 = p1
        driver.refresh()

    driver.switch_to.window(driver.window_handles[1])
    driver.execute_script(f"window.open('data:;')")
    driver.switch_to.window(driver.window_handles[2])
    driver.get(exten)
    wait.until(EC.element_to_be_clickable((By.ID, 'stop'))).click()
    time.sleep(0.5)
    driver.close()
    time.sleep(0.5)
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    time.sleep(0.5)
    driver.switch_to.window(driver.window_handles[0])

# def getSessionBondRequest (driver):
    # cookies = driver.get_cookies()
    # user_agent = driver.execute_script("return navigator.userAgent;")
    # referer = driver.current_url
    # accept_language = driver.execute_script("return navigator.language;")

    # session = requests.Session()
    # session.headers.update({
    #     'User-Agent': user_agent,
    #     'Accept-Language': accept_language,
    #     'Referer': referer,
    # })
    # for cookie in cookies:
    #     session.cookies.set(cookie['name'], cookie['value'])
    # return session

def runMSreward(i, user, password, dateObj):
    userdatadir = r'C:\Users\Tushar Kumar\AppData\Local\Google\Chrome\User Data'
    chromeOptions = webdriver.ChromeOptions()

    if '-h' in argv:
        chromeOptions.add_argument("--headless=new")
        print('Starting chrome in headless mode - ')

    chromeOptions.add_argument(f"--user-data-dir={userdatadir}")
    chromeOptions.add_argument("profile-directory=Profile 3")
    chromeOptions.add_experimental_option("excludeSwitches", ["enable-logging"])

    with webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chromeOptions) as driver:
        driver.maximize_window()

        def load_get_with_retry(driver, url):
            while True:
                try:
                    driver.get(url)
                    if "icon icon-generic" in driver.page_source:
                        time.sleep(1)
                    else: break
                except:
                    time.sleep(2)

        def execute_with_retry(driver, load_url):
            driver.execute_script(f"window.open('{load_url}')")
            driver.switch_to.window(driver.window_handles[1])

            while True:
                try:
                    if "icon icon-generic" in driver.page_source:
                        time.sleep(1)
                        driver.get(load_url)
                    else: break
                except:
                    time.sleep(2)

        def send_keys_with_retry(element, keys):
            _try = 1
            while True:
                try:
                    wait.until(EC.element_to_be_clickable(element)).send_keys(keys)
                    break
                except:
                    print(f"Tryin' again to write email, {_try}")
                    time.sleep(0.5)
                    _try += 1

        def click_with_retry(element):
            _try = 1
            while True:
                try:
                    wait.until(EC.element_to_be_clickable(element)).click()
                    break
                except:
                    print(f"Tryin' again to click on next button, {_try}")
                    time.sleep(0.5)
                    _try += 1

        wait = WebDriverWait(driver, 10)
        reward_sign = BASE_URL+'signin'

        load_get_with_retry(driver, reward_sign)

        if driver.current_url == "https://rewards.bing.com/":
            load_get_with_retry(driver, "https://rewards.bing.com/signout") 

        print(f'[N-{i}] - Signing in to {user} -')
        EMAILFIELD = (By.ID, "i0116")
        PASSWFIELD = (By.ID, "i0118")
        NEXTBUTTON = (By.ID, "idSIButton9")

        print(' [LOGIN]', 'Writing email...')
        wait.until(EC.element_to_be_clickable(EMAILFIELD)).send_keys(user)

        print(' [LOGIN]', 'Clicking next button...')
        click_with_retry(NEXTBUTTON)

        print(' [LOGIN]', 'Writing password...')
        wait.until(EC.element_to_be_clickable(PASSWFIELD)).send_keys(password)

        print(' [LOGIN]', 'Passing security checks...')
        click_with_retry(NEXTBUTTON)

        if 'Abuse' in driver.current_url:
            points = 'Blocked'
            fs = open('points.log', 'a', encoding='utf-8')
            fs.write(f"{dateObj[0]}-{dateObj[1]}-{dateObj[2]} :   {user : <40} :   {points : >10}\n")
            fs.close()
            return driver.quit()

        click_with_retry(NEXTBUTTON)

        if 'Rewards Error' in driver.title:
            points = 'Suspended'
            fs = open('points.log', 'a', encoding='utf-8')
            fs.write(f"{dateObj[0]}-{dateObj[1]}-{dateObj[2]} :   {user : <40} :   {points : >10}\n")
            fs.close()
            return driver.quit()
    
        print(' [LOGIN]', 'Logged-in !')
        execute_with_retry(driver, 'https://www.bing.com')

        print(' [LOGIN]', 'Ensuring login on Bing...')
        bingSignIn(driver)

        driver.switch_to.window(driver.window_handles[0])
        load_get_with_retry(driver, BASE_URL) if "icon icon-generic" in driver.page_source else None

        print(early_status := progress(driver, wait, None, None))
    
        if '--skip-spoint' not in argv:
            flag = False
            for i in early_status:  # check whether all the search is completed or not 
                if i[0] == i[1]:
                    flag = True
                else:
                    flag = False
                    break
                
            level = 0        
            if flag != True:
                for j in driver.find_elements(By.CSS_SELECTOR, '.c-caption-1'): # find the current reward level
                    reLevel = re.search(r"Level\s(\d)", j.get_attribute('innerHTML'))
                    if  reLevel != None:
                        level = int(reLevel.group(1))

                if level == 1 and (early_status[0][0] != early_status[0][1]):
                    startSearch(driver, wait, ABS_EXTENSION, 'Desktop', 1) # search for desktop level 1
                if level == 2:
                    if (early_status[0][0] != early_status[0][1]):
                        startSearch(driver, wait, ABS_EXTENSION, 'Desktop', 2) # secarch for desktop level 2
                    if (early_status[1][0] != early_status[1][1]):
                        startSearch(driver, wait, ABS_EXTENSION, 'Mobile', 2) #search for mobile level 2
                else:
                    print('unexpected search pattern')
            else:
                print('all daily search quota completed!')

        if '--skip-rpoint' not in argv:
            addSym = "AddMedium"
            reward_wait = WebDriverWait(driver, 5)
            driver.switch_to.window(driver.window_handles[0])
            load_get_with_retry(driver, reward_sign)

            # set the collectable value to 3 if collecting reward between 12:00 am to 9:00 am / skip the daily reward -
            # not_collectable_period = _time(hour=10, minute=0, second=0)
            # now = datetime.now().time()
            # isCollectable = 0
            # if now < not_collectable_period:
            #     isCollectable = 3     
            isCollectable = 0

            for reward in driver.find_elements(By.CSS_SELECTOR, f"span.mee-icon-{addSym}, span.mee-icon-HourGlass")[isCollectable:]:
                driver.switch_to.window(driver.window_handles[0])
                reward_wait.until(EC.element_to_be_clickable(reward)).click()

                time.sleep(0.5)
                if driver.current_url == BASE_URL+'legaltextbox':
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="legalTextBox"]/div/div/div[3]/a/span/ng-transclude'))).click()

                driver.switch_to.window(driver.window_handles[1])

                try:
                    driver.find_element(By.CSS_SELECTOR, ".rqText , .wk_BannerLine1New")
                    i = 3
                    while i > 1:
                        time.sleep(1)
                        i = len(driver.window_handles)
                        if i < 2:
                            break
                    driver.switch_to.window(driver.window_handles[0])
                except Exception as e:
                    time.sleep(5)
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
        
        if '--skip-plog' not in argv:    
            load_get_with_retry(driver, BASE_URL)
            time.sleep(2)
            points = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div[1]/div[2]/main/div/ui-view/mee-rewards-dashboard/main/mee-rewards-user-status-banner/div/div/div/div/div[2]/div[1]/mee-rewards-user-status-banner-item/mee-rewards-user-status-banner-balance/div/div/div/div/div/div/p/mee-rewards-counter-animation/span"))).text
            open('points.log', 'a', encoding='utf-8').write(f"{dateObj[0]}-{dateObj[1]}-{dateObj[2]} :   {user : <40} :   {points : >10}\n")
        
        if '--log-orders' in argv:
            isDone = logOrders(user, driver, BASE_URL+'redeem/orderhistory')
            if isDone == False:
                driver.close()
        time.sleep(5)
        driver.quit()

if __name__ == '__main__':

    kill_process('chrome')
    time.sleep(5)
    with open('data.json') as user_data:
        file_contents = user_data.read()
        data_ids = tuple(dict(json.loads(file_contents)).items())
    date_ = date.today()
    dateObj = [f'{date_.day:02d}' , f'{date_.month:02d}', f'{date_.year:04d}']

    terminal_len = os.get_terminal_size().columns

    i = 0
    for flag in argv:
        if flag.startswith('--start-'):
            i = int(flag.split('-')[-1])
            break
    user, password = (None, None)
    while i < len(data_ids):
        print(Fore.BLUE + "-"*(terminal_len-1), end=f"{Style.RESET_ALL} \n")
        clear_chrome_data()
        try:
            user, password = data_ids[i]
            i += 1
            if len(password.split(",")) == 2:
                print(f" ID - (flagged) skipping | {user} ")
                continue
            
            start = time.time()
            runMSreward(i, user, password, dateObj)
            end = time.time()
            print(f" Total execution time : {end - start}")
        except (Exception) as e:
            print(f' Something went wrong, while working with {user}\n {e}\nRetrying with same user data...')
            traceback.print_exc()
            i -= 1
        print(Fore.BLUE + "-"*(terminal_len-1), end=f"{Style.RESET_ALL} \n")
    fs = open('points.log', 'a', encoding='utf-8')
    fs.write('-'*90 + "\n")
    fs.close()
