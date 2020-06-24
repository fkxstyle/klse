import os
import shutil
import unittest
import time
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as cond

class MorningStar(unittest.TestCase):

    def setUp(self):
        self.download_path = "/Users/fookianxiong/Documents/Projects/klse/downloads/{}/"
        self.income_statement_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div[1]/div[1]/a'
        self.balance_sheet_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div[2]/div[1]/a'
        self.cash_flow_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div[3]/div[1]/a'
        self.export_excel_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div[2]/div/div[2]/div/div[1]/div[4]/button'
        self.back_to_summary_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div[2]/div/div[2]/div/div[2]/div[1]/h4/a'

    def testExportExcel(self):
        # load json data
        with open('tutorial/spiders/item.json') as json_file:
            stocks = json.load(json_file)

        morningstar_url = 'https://www.morningstar.com/stocks/xkls/{}/financials'

        for stock in stocks:
            url = morningstar_url.format(stock['code'])
            download_path = self.download_path.format(stock['name'])
            
            # Create directory if not exist
            if not os.path.exists(download_path):
                os.makedirs(download_path)

            options = Options()
            options.add_experimental_option("prefs", {
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            })
            browser = webdriver.Chrome(options=options)

            # Access morningstar website
            time.sleep(3)
            browser.get(url)

            # Wait til Income statement appear
            WebDriverWait(browser, 10).until(
                cond.presence_of_element_located(
                    (By.XPATH, self.income_statement_xpath)
                )
            )

            # Click continue to site, if any
            elements = browser.find_elements_by_xpath('//*[@id="__layout"]/div/div[4]/div/div/header/div/div[3]/button')
            if len(elements) > 0:
                element = elements[0]
                element.click()
                time.sleep(3)

            # Click Income statement
            element = browser.find_element_by_xpath(self.income_statement_xpath)
            element.click()
            time.sleep(3)
            # Click Export to excel
            element = browser.find_element_by_xpath(self.export_excel_xpath)
            element.click()
            time.sleep(3)
            # Click Back to Summary View
            element = browser.find_element_by_xpath(self.back_to_summary_xpath)
            element.click()
            time.sleep(3)

            # Click Balance Sheet
            element = browser.find_element_by_xpath(self.balance_sheet_xpath)
            element.click()
            time.sleep(3)
            # Click Export to excel
            element = browser.find_element_by_xpath(self.export_excel_xpath)
            element.click()
            time.sleep(3)
            # Click Back to Summary View
            element = browser.find_element_by_xpath(self.back_to_summary_xpath)
            element.click()
            time.sleep(3)

            # Click Cash Flow
            element = browser.find_element_by_xpath(self.cash_flow_xpath)
            element.click()
            time.sleep(3)
            # Click Export to excel
            element = browser.find_element_by_xpath(self.export_excel_xpath)
            element.click()
            time.sleep(3)


if __name__ == '__main__':
    unittest.main(verbosity=4)