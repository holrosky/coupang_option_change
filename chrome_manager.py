import os
import shutil
import json
import sys
import time

import requests as requests
from selenium import webdriver
from selenium.webdriver import ActionChains

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from fake_useragent import UserAgent


class ChromeManager():
    def __init__(self):
        self.driver = None

    def open_chrome(self):
        try:
            shutil.rmtree(r'C:\chrometemp')
        except Exception as e:
            pass

        # subprocess.Popen(
        #     r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"')
        #
        user_ag = UserAgent().random
        #
        options = webdriver.ChromeOptions()

        with open("config.json", "r", encoding='utf-8') as st_json:
            json_data = json.load(st_json)

        background = json_data['background']

        if 'y' in background.lower():
            options.add_argument('headless')

        options.add_argument('window-size=1910x1080')
        options.add_argument('lang=ko_KR')
        options.add_argument('user-agent=%s' % user_ag)
        # options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        if not self.driver:
            self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)

        # self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        #     "source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})

        self.driver.implicitly_wait(3)

    def log_in(self, coupang_id, coupang_pwd):
        print('Logging in...')

        self.driver.get(url='https://with.coupang.com/tickets/management#')

        self.wait_until_clickable(15, "//input[@id='username']")
        time.sleep(1)

        with open("config.json", "r", encoding='utf-8') as st_json:
            json_data = json.load(st_json)

        sms_api_url = json_data['sms_api_url']
        sms_api_key = json_data['sms_api_key']

        self.send_key("//input[@id='username']", coupang_id)
        self.send_key("//input[@id='password']", coupang_pwd + Keys.ENTER)

        try:
            self.wait_until_clickable(15, "//input[@id='btnSms']")

            self.click("//input[@id='btnSms']")

            self.wait_until_clickable(15, "//input[@id='auth-mfa-code']")

            while True:
                try:
                    time.sleep(7)

                    response = requests.get(url=sms_api_url, params={'key': sms_api_key})

                    auth_code = response.json()
                    print('Auth code : ' + str(auth_code['auth']))

                    self.send_key("//input[@id='auth-mfa-code']", auth_code['auth'] + Keys.ENTER)

                    self.wait_until_clickable(15, "//img[@alt='쿠팡로고']")

                    break
                except Exception as e:
                    print(e)
                    self.wait_until_clickable(15, "//input[@id='auth-mfa-code']")


        except Exception as e:
            self.wait_until_clickable(15, "//img[@alt='쿠팡로고']")

        print('Logging in done!')

    def change_option(self, product_code, option_text):
        try:

            try:
                self.search_product(product_code)

                self.wait_until_clickable(15, "//button[@class='btn-s-ty03  big vendor-item-invalid-button']")
                self.click("//button[@class='btn-s-ty03  big vendor-item-invalid-button']")

                self.wait_alert()

                time.sleep(2)

            except Exception as e:
                print(e)
                self.wait_until_clickable(15, "//button[@class='btn-s-ty03 green big vendor-item-valid-button']")

            self.wait_until_clickable(15, "//button[@class='btn-s-ty03 small vendor-item-option-modify-button']")
            self.click("//button[@class='btn-s-ty03 small vendor-item-option-modify-button']")

            while len(self.driver.window_handles) < 2:
                time.sleep(0.5)

            self.driver.switch_to.window(self.driver.window_handles[-1])

            self.wait_until_clickable(15, "//i[@class='ico-edit']")

            while True:
                if self.driver.find_element(By.XPATH, "//div[@id='js-option-info-wrap']").is_displayed():
                    break
                else:
                    self.click("//i[@class='ico-edit']")

                time.sleep(0.5)

            self.wait_until_clickable(15,
                                      "//input[@class='form-control js-validation-text-max-length-input js-validation-empty-input js-border']")

            element = self.driver.find_element(By.XPATH,
                                               "//input[@class='form-control js-validation-text-max-length-input js-validation-empty-input js-border']")
            element.clear()
            element.send_keys(option_text)

            self.click("//button[@id='js-option-register-btn']")

            self.wait_until_clickable(15, "//button[@class='btn btn-st01 green js-continue-temp']")
            time.sleep(3)
            self.click("//button[@class='btn btn-st01 green js-continue-temp']")

            self.wait_until_clickable(15, "//div[@id='js-select-target-ticket-receipt']")

            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[-1])

            self.search_product(product_code)

            self.wait_until_clickable(15, "//button[@class='btn-s-ty03 green big vendor-item-valid-button']")
            self.click("//button[@class='btn-s-ty03 green big vendor-item-valid-button']")

            self.wait_alert()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(e, fname, exc_tb.tb_lineno)
            raise Exception

    def wait_alert(self):
        alert_count = 0

        while True:
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
                break
            except Exception:
                alert_count += 1
                if alert_count >= 10:
                    raise Exception
                time.sleep(1)

    def search_product(self, product_code):
        self.driver.refresh()

        self.wait_until_clickable(15, "//input[@id='travelName']")

        time.sleep(5)

        element = self.driver.find_element(By.XPATH, "//input[@id='travelName']")
        element.clear()
        element.send_keys(product_code + Keys.ENTER)

        time.sleep(1)

        while self.driver.find_element(By.XPATH, "//div[@id='blocking_wait_modal']").is_displayed():
            time.sleep(0.5)

    def wait_until_clickable(self, wait_time, xpath):
        WebDriverWait(self.driver, wait_time).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    def send_key(self, xpath, keys):
        self.driver.find_element(By.XPATH, xpath).send_keys(keys)

    def click(self, xpath):
        self.driver.find_element(By.XPATH, xpath).click()
        # action = ActionChains(self.driver)
        # action.move_to_element(self.driver.find_element(By.XPATH, xpath)).click().perform()

    def close(self):
        self.driver.quit()
        self.driver = None
