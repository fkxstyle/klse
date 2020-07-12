import os
import shutil
import glob
import unittest
import time
import json
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as cond

class MorningStar(unittest.TestCase):

    def setUp(self):
        sector_root = Path().absolute()
        ## Configuration
        self.download_path = os.path.join(sector_root,'downloads/')
        self.restated_xpath = "//div[@class='sal-component-body']//sal-components-segment-band[@class='report-type']//mds-button[2]//label[1]//input[1]"
        self.income_statement_xpath = "//div[@class='sal-tab-content ng-scope']//a[@class='mds-link ng-binding'][contains(text(),'Income Statement')]"
        self.balance_sheet_xpath = "//div[@class='sal-tab-content ng-scope']//a[@class='mds-link ng-binding'][contains(text(),'Balance Sheet')]"
        self.cash_flow_xpath = "//div[@class='sal-tab-content ng-scope']//a[@class='mds-link ng-binding'][contains(text(),'Cash Flow')]"
        self.export_excel_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div[2]/div/div[2]/div/div[1]/div[4]/button'
        self.back_to_summary_xpath = '//*[@id="__layout"]/div/div[3]/main/div[2]/div/div/div[1]/sal-components/section/div/div/div/div/div[2]/div/div[2]/div/div[2]/div/div[2]/div[1]/h4/a'
        self.valuation_xpath = "//a[contains(text(),'Valuation')]"
        self.operating_performance_xpath = "//a[contains(text(),'Operating Performance')]"

        self.json_file_path = 'stocks.json'

    def testExportExcel(self):
        # Delete all folders and files in downloads folder
        downloaded_files = glob.glob(self.download_path + '*')
        for fpath in downloaded_files:
            if os.path.isdir(fpath):
                shutil.rmtree(fpath)
            elif os.path.isfile(fpath):
                os.remove(fpath)

        # load json data
        with open(self.json_file_path) as json_file:
            stocks = json.load(json_file)

        morningstar_url = 'https://www.morningstar.com/stocks/xkls/{}/financials'
        
        # Instiantiate browser
        options = Options()
        options.add_experimental_option("prefs", {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        browser = webdriver.Chrome(options=options)

        for stock in stocks:
            url = morningstar_url.format(stock['code'])

            # Access morningstar website
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

            # # Click Restated
            # element = browser.find_element_by_xpath(self.restated_xpath)
            # element.click()
            # time.sleep(3)

            # # Click Income statement
            # element = browser.find_element_by_xpath(self.income_statement_xpath)
            # element.click()
            # time.sleep(3)
            
            # # Click Export to excel
            # element = browser.find_element_by_xpath(self.export_excel_xpath)
            # element.click()
            # time.sleep(3)
            # # Click Back to Summary View
            # element = browser.find_element_by_xpath(self.back_to_summary_xpath)
            # element.click()
            # time.sleep(3)

            # # Click Balance Sheet
            # element = browser.find_element_by_xpath(self.balance_sheet_xpath)
            # element.click()
            # time.sleep(3)
            # # Click Export to excel
            # element = browser.find_element_by_xpath(self.export_excel_xpath)
            # element.click()
            # time.sleep(3)
            # # Click Back to Summary View
            # element = browser.find_element_by_xpath(self.back_to_summary_xpath)
            # element.click()
            # time.sleep(3)

            # # Click Cash Flow
            # element = browser.find_element_by_xpath(self.cash_flow_xpath)
            # element.click()
            # time.sleep(3)
            # # Click Export to excel
            # element = browser.find_element_by_xpath(self.export_excel_xpath)
            # element.click()
            # time.sleep(6)

            # Go to Valuation 
            element = browser.find_element_by_xpath(self.valuation_xpath)
            element.click()
            time.sleep(3)

            # Dismiss pop-up
            elements = browser.find_elements_by_xpath('//*[@id="__layout"]/div/div[4]/div/div/div[1]/button')
            if len(elements) > 0:
                element = elements[0]
                element.click()
                time.sleep(3)

            element = browser.find_element_by_xpath("//tr[4]//td[12]")
            stock['current_pe_ratio'] = element.text
            element = browser.find_element_by_xpath("//tr[4]//td[11]")
            stock['last_year_pe_ratio'] = element.text
            element = browser.find_element_by_xpath("//tr[4]//td[10]")
            stock['last_2_years_pe_ratio'] = element.text
            
            # Go to Operating Performance 
            element = browser.find_element_by_xpath(self.operating_performance_xpath)
            element.click()
            time.sleep(3)

            stock['days_sales_outstanding'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[11]/td[11]/span[1]").text
            stock['days_inventory'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[12]/td[11]/span[1]").text
            stock['days_payables'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[13]/td[11]/span[1]").text
            stock['receivable_turnover'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[14]/td[11]/span[1]").text
            stock['inventory_turnover'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[15]/td[11]/span[1]").text
            stock['fixed_asset_turnover'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[16]/td[11]/span[1]").text
            stock['total_asset_turnover'] = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/main[1]/div[2]/div[1]/div[1]/div[1]/sal-components[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[17]/td[11]/span[1]").text

            # Create directory for stock if not exist
            if not os.path.exists(self.download_path + stock['name']):
                os.makedirs(self.download_path + stock['name'])
        
            # Get list of downloaded excel files
            excel_file_paths = glob.glob("downloads/*.xls")
            # move to stock folder
            for excel_file_path in excel_file_paths:
                shutil.move(excel_file_path, self.download_path + stock['name'])

            # add data.json in stock folder
            with open(self.download_path + stock['name'] + '/data.json', 'w') as outfile:
                json.dump(stock, outfile)

        # move json file to downloads folder
        # shutil.move(self.json_file_path, self.download_path)

if __name__ == '__main__':
    unittest.main(verbosity=4)
    # python tutorial/spiders/morningstar_selenium.py