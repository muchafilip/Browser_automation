###############################################################################
###############################################################################
# First commit 1-30-2020
# CREATED: Jan 2020
# AUTHOR: filip.mucha@rws.com
# version
# v0.1: analyses download, apply MT, download target assets
# v0.2: added check for checkbox (var checkbox)
# v0.3: read Project # from HTML table => User does not need to input project numbers in arguemnts
# v0.4: added review assignment functionality (from csv)
# v0.5: added logging
# USAGE: python wsbot.py -a
# "-a" to download analysis 
# "-d" to download target assets
# "-m" to apply MT
# "-r" to review assignment
# TODO use internal xpaths in all functions
###############################################################################
###############################################################################
from datetime import datetime
from time import sleep
import re
import csv
import sys
import logging
from credentials import username, url, password, driverpath, log_path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# logger config
logger_file_name = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
log_path = log_path + '\\'+ str(logger_file_name) + '.log'
logging.basicConfig(level=logging.INFO, filename=log_path, filemode='w', format='%(asctime)s - %(levelname)s - %(message)s')

main_screen_xpath = "/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr/td[5]/a[2]"


class SchneiderWorldServerBot:

    # all_xpaths is return value of generate_xpaths functions.
    # Array containing xpaths of clickable "project numbers" - 
    # bot uses this for page navigation

    def __init__(self, url, username, password, all_xpaths=[], counter=0):
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("excludeSwitches", ['enable-automation'])
        self.driver = webdriver.Chrome(executable_path=driverpath, options=self.options)
        self.all_xpaths = all_xpaths
        self.counter = counter
        self.driver.get(url)

        username_box = self.driver.find_element_by_id("username")
        username_box.send_keys(username)
        password_box = self.driver.find_element_by_id("password")
        password_box.send_keys(password)
        password_box.submit()
        print('logged in')
        logging.info('logged in')


    def generate_xpaths(self, all_xpaths):
        """Return all xpaths for selected projects so other functions can use them to navigate page. """
        # get rows(elements) from html table
        rows = self.driver.find_elements_by_xpath('html/body/form/table/tbody/tr/td/table/tbody/tr')
        # get rid of first 5 rows (headers, buttons)
        rows = rows[5:-1]
        print(f"{len(rows)} projects selected:")
        # columns = self.driver.find_elements_by_xpath('html/body/form/table/tbody/tr/td/table/thead/tr/th')
        # For each row, take only 6 digit project number, pass it to xpath. Finally, append each xpath to the all_xpaths array.
        for row in rows:
            # store text of each element(row in variable)
            mytext = row.text
            splitted = mytext.split()
            [all_xpaths.append(f'//*[@id="{x}"]/td[2]/a') for x in splitted if len(x) == 6 and x.isdigit()]
        [print(pn[9:15]) for pn in all_xpaths] 
        logging.info('xpaths generated')

        return all_xpaths


    def assign_to_review(self, xpath, reviewer_dict, web_language):

        internal_xpaths = {
        'checkbox':'//*[@id="activeTasks"]',
        'complete button':'/html/body/form/table/tbody/tr[2]/td/table[1]/tbody/tr/td[9]/a',
        'dropdown_1': '/html/body/form/table/tbody/tr[2]/td/p/table[1]/tbody/tr[1]/td[2]/select',
        'dropdown_2': '/html/body/form/table/tbody/tr[2]/td/p/table[1]/tbody/tr[2]/td[2]/select',
        'dropdown_3': '/html/body/form/table/tbody/tr[2]/td/p/table[1]/tbody/tr[3]/td[2]/select',
        'ok_button_assignment': '/html/body/form/table/tbody/tr[3]/td/input[1]',
        'approve_cost_est_button': '//*[@id="__wsDialog_button_ok"]'
        }

        self.driver.find_element_by_xpath(xpath).click()
        main_window = self.driver.window_handles[0]
        print(f"Project name: {self.driver.title[53:]}")
        print("lang read from html=", web_language)
        # check if all tasks are selected/select all
        checkbox = self.driver.find_element_by_xpath(internal_xpaths['checkbox'])
        if not checkbox.is_selected():
            checkbox.click()
        main_window = self.driver.window_handles[0]
        self.driver.find_element_by_xpath(internal_xpaths['complete button']).click()
        popup_window = self.driver.window_handles[1]
        self.driver.switch_to.window(popup_window)
        # click vendor Moravia
        WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, internal_xpaths['dropdown_1'])))
        el = self.driver.find_element_by_xpath(internal_xpaths['dropdown_1'])
        for option in el.find_elements_by_tag_name('option'):
            if option.text == 'Moravia 2016 [Moravia Translation 2016]':
                option.click()
                print(f'Vendor selected')
                break
    
        WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, internal_xpaths['dropdown_2'])))
        el = self.driver.find_element_by_xpath(internal_xpaths['dropdown_2'])
        for option in el.find_elements_by_tag_name('option'):
            if option.text == 'Moravia 2016 [Moravia DTP]':
                option.click()
                print('DTP Vendor selected')
                break
        
        WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, internal_xpaths['dropdown_3'])))
        if web_language in reviewer_dict.keys():
            rev_name = reviewer_dict[web_language]
            try:
                if len(rev_name) > 1:
                    print('Multiple reviewers found')
                    for name in rev_name:
                        element_1 = self.driver.find_element_by_xpath(f"//select[@name='assignReviewer']/option[text()='{name}']")
                        ActionChains(self.driver).key_down(Keys.CONTROL).click(element_1).key_up(Keys.CONTROL).perform()
                        print(name, " found")               
                else:
                    name = rev_name[0]
                    print(name, " found")
                    element_1 = self.driver.find_element_by_xpath(f"//select[@name='assignReviewer']/option[text()='{name}']")
                    ActionChains(self.driver).key_down(Keys.CONTROL).click(element_1).key_up(Keys.CONTROL).perform()
            except Exception:
                logging.error("Exception occurred", exc_info=True)


        self.driver.find_element_by_xpath(internal_xpaths['ok_button_assignment']).click() #submit button
        logging.info(f'Step review assign completed for {web_language,rev_name}')
        self.driver.switch_to.window(main_window)
        sleep(1)
        # Approve cost estimate step
        # checkbox = self.driver.find_element_by_xpath(internal_xpaths['checkbox'])
        # if not checkbox.is_selected():
        #     checkbox.click()
        # self.driver.find_element_by_xpath(internal_xpaths['complete_button']).click()
        # popup_window = self.driver.window_handles[1]
        # self.driver.switch_to.window(popup_window)
        # sleep(2)
        # self.driver.find_element_by_xpath(internal_xpaths['approve_cost_est_button']).click()
        # self.driver.switch_to.window(main_window)
        self.driver.find_element_by_xpath(main_screen_xpath).click()
        print('assignment successfully completed')
        logging.info(f'Step approve cost completed for {web_language, rev_name}')


    def apply_MT(self, xpath):
        self.driver.find_element_by_xpath(xpath).click()
        print(self.driver.title[53:])
        main_window = self.driver.window_handles[0]
        # check if all tasks are selected/select all
        checkbox = self.driver.find_element_by_xpath('//*[@id="activeTasks"]')
        if not checkbox.is_selected():
            checkbox.click()
        # complete button
        self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table[1]/tbody/tr/td[9]/a').click()
        popup_window = self.driver.window_handles[1]
        self.driver.switch_to.window(popup_window)
        # click OK button
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="__wsDialog_button_ok"]'))).click()
        self.driver.switch_to.window(main_window)
        # click on projects
        self.driver.find_element_by_xpath(main_screen_xpath).click()
        print(f"Machine translation for task #{xpath[9:15]} enabled.")
        logging.info(f"task {xpath[9:15]} success - apply machine translation")


    def download_analysis(self, xpath):
        self.driver.find_element_by_xpath(xpath).click()
        title = self.driver.title[53:]
        print(title)
        main_window = self.driver.window_handles[0]
        self.driver.find_element_by_xpath("/html/body/form/table/tbody/tr[2]/td/a[2]").click()
        popup_window = self.driver.window_handles[1]
        self.driver.switch_to.window(popup_window)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/table/tbody/tr[3]/td/p[1]/a"))).click()
        self.driver.switch_to.window(main_window)
        sleep(0.5)
        self.driver.find_element_by_xpath(main_screen_xpath).click()
        print(f"Analysis #{xpath[9:15]} downloaded")
        logging.info(f"task {xpath[9:15]} success - download analysis")



    def download_assets(self, xpath):
        print(xpath[9:15])
        self.driver.find_element_by_xpath(xpath).click()
        print(self.driver.title[53:])
        # dropdown xpaths
        more_options_button = self.driver.find_element_by_xpath("//a[contains(text(),'More Options')]")
        asset_options_button = self.driver.find_element_by_xpath("//td[contains(text(),'Asset Options')]")
        download_button = self.driver.find_element_by_xpath("//a[contains(text(),'Download')]")
        # define main window
        main_window = self.driver.window_handles[0]
        # check if all tasks are selected/select all
        checkbox = self.driver.find_element_by_xpath('//*[@id="activeTasks"]')
        if not checkbox.is_selected():
            checkbox.click()
        # complete button
        actions = ActionChains(self.driver)
        actions.move_to_element(more_options_button)
        actions.move_to_element(asset_options_button)
        actions.move_to_element(download_button).click().perform()
        # define popup window & switch
        popup_window = self.driver.window_handles[1]
        self.driver.switch_to.window(popup_window)
        # click Target assets
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/table/tbody/tr[3]/td/nobr[2]/label'))).click()
        self.driver.find_element_by_xpath('//*[@id="__wsDialog_button_download"]').click()
        # switch back to main window
        self.driver.switch_to.window(main_window)
        # click on projects
        self.driver.find_element_by_xpath(main_screen_xpath).click()
        print(f"Downloaded target assets for task #{xpath[9:15]}")
        logging.info(f"task {xpath[9:15]} success - download target assets")

    def read_language(self, xpath):
            lang_xpath_number = xpath[9:15]
            language = f'//*[@id="{lang_xpath_number}"]/td[3]'
            language = self.driver.find_element_by_xpath(language)
            logging.info(language.text)
            return language.text

def get_dict_from_csv(inputfile):
    with open(inputfile, 'r', newline="\n", encoding="utf-8") as f:
        data = csv.reader(f, delimiter=',')
        reviewer_dict = {row[0]:row[1:] for row in data}
        logging.info('csv file read successfully')
        return reviewer_dict
        


# scripts starts here
bat_input = input("""Type:
# "a" to download analysis 
# "d" to download target assets
# "m" to apply Machine Translation
# "r" to review assignment\n
""")
bot = SchneiderWorldServerBot(url, username, password)
logging.info('bot init')
bot.generate_xpaths(bot.all_xpaths)
reviewer_dict = get_dict_from_csv('reviewers.csv')

for xpath in bot.all_xpaths:
    try:
        web_language = bot.read_language(xpath)
        if bat_input == "a":
            bot.download_analysis(xpath)
        if bat_input == "m":
            bot.apply_MT(xpath)
        if bat_input == "d":
            bot.download_assets(xpath)
        if bat_input == "r":
            bot.assign_to_review(xpath, reviewer_dict, web_language)  
        counter += 1
        print(".")
    except Exception:
        print('Something went wrong')
        logging.error(f"task {xpath[9:15]} fail", exc_info=True)
        
print(f'{counter} tasks processed')




