from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time
import sys
import os
from dotenv import load_dotenv
import pandas as pd
from datetime import datetime
df = pd.read_csv('time_table.csv')
load_dotenv()
USERNAME= os.getenv('USERNAME')
PASSWORD= os.getenv('PASSWORD')
max_classes= len(df[df['Day']=="Tuesday"])
class_num = 0
while class_num<max_classes:
    now = datetime.now()
    time_now = now.strftime("%H:%M")
    if time_now[0] == "0":
        time_now = time_now[1:]
    day = now.strftime("%A")
    TEAM = df[df['Day'] == day].reset_index()['Teams Name'][class_num]
    CHANNEL = df[df['Day'] == day].reset_index()['Channel'][class_num]
    Time_class = df[df['Day'] == day].reset_index()['Time'][class_num]
    Duration = int(df[df['Day'] == day].reset_index()['Duration(in sec)'][class_num])
    hour = ""
    for i in Time_class:
        if i ==":":
            break
        else:
            hour = hour+i
    Time_class_datetime = now.replace(hour=int(hour), minute=int(Time_class[-2:]), second=0, microsecond=0)
    if time_now == Time_class:
        opt = Options()
        opt.add_argument("--disable-infobars")
        opt.add_argument("start-maximized")
        opt.add_argument("--disable-extensions")
        # Pass the argument 1 to allow and 2 to block
        opt.add_experimental_option("prefs", {
            "profile.default_content_setting_values.media_stream_mic": 1,
            "profile.default_content_setting_values.media_stream_camera": 1,
            "profile.default_content_setting_values.geolocation": 1,
            "profile.default_content_setting_values.notifications": 1
        })

        browser = webdriver.Chrome(executable_path='/Users/harshitkalra/Desktop/ml/chromedriver', options=opt)
        browser.get(
            "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=id_token&scope=openid%20profile&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=eyJpZCI6ImI3ZTU5N2Q0LWRmZjktNDQ4NC1iYzUxLWIxOTU4Y2JlZmM2YiIsInRzIjoxNjQxNzMyMTk4LCJtZXRob2QiOiJyZWRpcmVjdEludGVyYWN0aW9uIn0%3D&nonce=1da36419-1834-4b4a-a08f-c4de1a73505a&client_info=1&x-client-SKU=MSAL.JS&x-client-Ver=1.3.4&client-request-id=358ac014-9407-487a-ac54-ba477b401de0&response_mode=fragment")

        time.sleep(3)

        try:
            login = browser.find_element_by_id("i0116")
            next1 = browser.find_element_by_id('idSIButton9')
        except:
            login = None
            next1 = None
        i = 0
        while i < 3 and login is None:
            time.sleep(3)
            i += 1
            try:
                login = browser.find_element_by_id("i0116")
                next1 = browser.find_element_by_id('idSIButton9')
            except:
                login = None
                next1 = None
        if i == 3:
            sys.exit('Browser too slow')
        login.send_keys(USERNAME)
        next1.click()

        time.sleep(2)

        try:
            password = browser.find_element_by_id("i0118")
            next1 = browser.find_element_by_id('idSIButton9')
        except:
            password = None
            next1 = None
        i = 0
        while i < 3 and login is None:
            time.sleep(3)
            i += 1
            try:
                password = browser.find_element_by_id("i0118")
                next1 = browser.find_element_by_id('idSIButton9')
            except:
                password = None
                next1 = None

        if i == 3:
            sys.exit('Browser too slow')

        password.send_keys(PASSWORD)
        next1.click()

        time.sleep(2)

        try:
            dontshow = browser.find_element_by_id("KmsiCheckboxField")
            yes = browser.find_element_by_id("idSIButton9")
        except:
            dontshow = None
            yes = None
        i = 0
        while i < 3 and dontshow is None:
            time.sleep(3)
            i += 1
            try:
                dontshow = browser.find_element_by_id("KmsiCheckboxField")
                yes = browser.find_element_by_id("idSIButton9")
            except:
                dontshow = None
                yes = None

        if dontshow is not None and yes is not None:
            dontshow.click()
            yes.click()

        time.sleep(3)

        try:
            kerbros = browser.find_element_by_xpath("//div[@data-test-id=" + f"'{USERNAME}'" + "]")
        except:
            kerbros = None
        i = 0
        while i < 3 and kerbros is None:
            time.sleep(3)
            i += 1
            try:
                kerbros = browser.find_element_by_xpath("//div[@data-test-id=" + f"'{USERNAME}'" + "]")
            except:
                kerbros = None

        if i == 3:
            sys.exit('Browser too slow')

        kerbros.click()

        time.sleep(15)

        try:
            useonweb = browser.find_element_by_xpath("//a[@data-tid='early-desktop-promo-use-web']")
        except:
            useonweb = None
        i = 0
        while i < 3 and useonweb is None:
            time.sleep(3)
            i += 1
            try:
                useonweb = browser.find_element_by_xpath("//a[@data-tid='early-desktop-promo-use-web']")
            except:
                useonweb = None
        if useonweb is not None:
            useonweb.click()

        time.sleep(7)

        try:
            teams = browser.find_element_by_id("app-bar-2a84919f-59d8-4441-a975-2a8c2643b741")
        except:
            teams = None
        i = 0
        while i < 3 and teams is None:
            time.sleep(3)
            i += 1
            try:
                teams = browser.find_element_by_id("app-bar-2a84919f-59d8-4441-a975-2a8c2643b741")
            except:
                teams = None

        if i == 3:
            sys.exit('Browser too slow')

        teams.click()

        time.sleep(2)
        try:
            class_name = browser.find_element_by_xpath("//div[@data-tid='team-" + f"{TEAM}'" + "]")
        except:
            class_name = None
        i = 0
        while i < 2 and class_name is None:
            time.sleep(3)
            i += 1
            try:
                class_name = browser.find_element_by_xpath("//div[@data-tid='team-" + f"{TEAM}'" + "]")
            except:
                class_name = None
        scroll_height = browser.execute_script(
            "return document.querySelector('[data-tid=\"team-channel-list\"]').scrollHeight")
        i = 0
        while 200*i<scroll_height:
            time.sleep(1)
            i+=1
            browser.execute_script(f"document.querySelector('[data-tid=\"team-channel-list\"]').scrollTo(0,200*{i})")
            print(2)
            try:
                class_name = browser.find_element_by_xpath("//div[@data-tid='team-" + f"{TEAM}'" + "]")
            except:
                class_name = None

        if class_name is None:
            sys.exit("Teams not found")

        class_name.click()

        time.sleep(2)

        if CHANNEL.lower() != "general":
            try:
                channel_name = browser.find_element_by_xpath("//li[@data-tid='team-" + f"{TEAM}-channel-{CHANNEL}-li'" + "]")
            except:
                channel_name = None
            i = 0
            while i < 3 and channel_name is None:
                time.sleep(3)
                i += 1
                try:
                    channel_name = browser.find_element_by_xpath("//li[@data-tid='team-" + f"{TEAM}-channel-{CHANNEL}-li'" + "]")
                except:
                    channel_name = None
            if i ==3:
                sys.exit("Channel not found")
            channel_name.click()

        time.sleep(2)

        try:
            join1 = browser.find_element_by_xpath("//button[@title='Join call with video']")
        except:
            join1 = None
        i = 0
        while i < 3 and join1 is None:
            time.sleep(3)
            i += 1
            try:
                join1 = browser.find_element_by_xpath("//button[@title='Join call with video']")
            except:
                join1 = None
        i = 0
        while i < 3 and join1 is None:
            time.sleep(180)
            i += 1
            try:
                join1 = browser.find_element_by_xpath("//button[@title='Join call with video']")
            except:
                join1 = None

        if i == 3:
            sys.exit('Browser too slow')

        join1.click()

        time.sleep(2)

        video_off = browser.find_element_by_xpath("//span[@title='Turn camera off']")
        if video_off is not None:
            video_off.click()
        audio_off = browser.find_element_by_xpath("//span[@title='Mute microphone']")
        if audio_off is not None:
            audio_off.click()

        time.sleep(1)

        join2 = browser.find_element_by_xpath("//button[@data-tid='prejoin-join-button']")
        join2.click()

        time.sleep(Duration)
        browser.quit()
        class_num+=1
    elif now> Time_class_datetime:
        class_num+=1
    else:
        pass



