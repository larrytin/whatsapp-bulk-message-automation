# Program to send bulk customized message through WhatsApp web application

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from random import randrange
import pandas as pd
import time
import config
import argparse


class WhatsappMessage(object):
    """
    A class that encapsulates Whatsapp Message automation
    function and attributes
    """

    def __init__(self, **kwargs):
        self.sheet_name = kwargs.get('sheet_name')
        self.image_path = kwargs.get('image_path')
        self.url = kwargs.get('recipients_xlsx')
        self.excel_data = None
        self.driver = None
        self.driver_wait = None

    def start_process(self):
        try:
            self.read_data()
            self.load_driver()
            send_time = self.process_message()
            self.excel_data['发送时间'] = pd.Series(send_time)
            # 创建一个 ExcelWriter 对象
            with pd.ExcelWriter(self.url, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # 将新的 DataFrame 写入到特定的工作表中
                self.excel_data.reset_index('Contact').to_excel(writer, sheet_name=self.sheet_name, index=False)
        finally:
            self.close_driver()

    def read_data(self):
        # Read data from excel
        self.excel_data = pd.read_excel(self.url, sheet_name=self.sheet_name, engine='openpyxl', dtype=str).set_index('Contact')

    def load_driver(self):
        # Load the chrome driver
        options = webdriver.ChromeOptions()
        options.add_argument(config.CHROME_PROFILE_PATH)
        if config.os_name == 'Windows':
            self.driver = webdriver.Chrome(executable_path=r'C:\Users\Nityam\AppData\Local\Programs\Python\Python39\chromedriver.exe',
                                           options=options)
        else:
            self.driver = webdriver.Chrome(options=options)

        # Open WhatsApp URL in chrome browser
        self.driver.get("https://web.whatsapp.com")
        self.driver_wait = WebDriverWait(self.driver, 20)

    def process_message(self):
        count = 0
        send_time = {}
        self.driver.set_page_load_timeout(60)
        # Iterate excel rows till to finish
        for column in self.excel_data.index[self.excel_data['发送时间'].isna() | (self.excel_data['发送时间'] == '')]:
            # Assign customized message
            message = self.excel_data['Message'][0]

            # Locate search box through x_path
            search_box = '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div'
            person_title = self.driver_wait.until(
                lambda driver: driver.find_element("xpath", search_box))

            # Clear search box if any contact number is written in it
            person_title.clear()

            # Send contact number in search box
            contact_number = str(column)

            # if len(contact_number) == 10 or not contact_number.startswith('91'):
            #     contact_number = '91' + contact_number
            person_title.send_keys(contact_number)

            # Wait for 2 seconds to search contact number
            time.sleep(2)

            try:
                # Load error message in case unavailability of contact number
                self.driver.find_element("xpath", '//*[@id="pane-side"]/div[1]/div/span')
                user_url = f'https://web.whatsapp.com/send?phone={contact_number}'
                self.driver.get(user_url)

                # Wait for 5 seconds to load user chat message
                time.sleep(5)

            except NoSuchElementException:
                person_title.send_keys(Keys.ENTER)

            if self.image_path is not None:
                attachment_button_path = '//span[@data-icon="attach-menu-plus"]'
                attachment_button = self.driver_wait.until(lambda driver: driver.find_element("xpath",
                                                                                              attachment_button_path))
                attachment_button.click()
                image_button_path = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
                image_button = self.driver_wait.until(
                    lambda driver: driver.find_element("xpath", image_button_path))
                image_button.send_keys(self.image_path)
                time.sleep(2)
                self.send_message(message)

            else:
                self.send_message(message)

            timelapse = randrange(3, 10)
            time.sleep(timelapse)
            count = count + 1
            send_time[contact_number] = time.strftime("%Y-%m-%d %H:%M:%S")
        return send_time

    def send_message(self, message):
        # Format the message from excel sheet
        # message = message.replace('{customer_name}', str(self.excel_data['Name'][count]))
        actions = ActionChains(self.driver)
        message = message.replace("\n", '__new_line__')
        message = message.replace("\r", '__new_line__')
        msg_lines = message.split('__new_line__')
        msg_lines[:] = [msg for msg in msg_lines if msg.strip()]
        for msg in msg_lines:
            actions.send_keys(msg)
            actions.key_down(Keys.SHIFT).key_down(
                Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER)
            actions.key_down(Keys.SHIFT).key_down(
                Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER)
        actions.send_keys(Keys.ENTER)
        actions.perform()
        time.sleep(2)

    def close_driver(self):
        # Close chrome browser
        self.driver.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Whatsapp Bulk Message Automation with optional Attachment feature')
    parser.add_argument('sheet_name', help='Sheet name', type=str)
    parser.add_argument('recipients_xlsx', help='whatsapp_recipients.xlsx', type=str)
    parser.add_argument(
        '--image-path', help='Full path of image attachment', type=str, dest='image_path')
    parsed_args = parser.parse_args()
    args = vars(parsed_args)
    whatsapp = WhatsappMessage(**args)
    whatsapp.start_process()
