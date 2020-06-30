import os
import scrapy
import json
import glob

class KlseSpider(scrapy.Spider):
    name = "klse"
   
    def start_requests(self):
        urls = []

        # delete content in stocks.json
        f = open('stocks.json', 'r+')
        f.truncate(0)

        # Take only first sector
        sector_json_file_path = glob.glob('sector/*')[0]

        # load json data
        with open(sector_json_file_path) as json_file:
            sector_stocks = json.load(json_file)
        
        # collect all urls
        for sector_stock in sector_stocks:
            urls.append(
                'https://www.klsescreener.com/v2/stocks/view/{}'.format(sector_stock['stock_code'])
            )
        
        # urls = [
        #     'https://www.klsescreener.com/v2/stocks/view/7087',
        # ]

        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        selected_data = [
            'Market Cap',
            'DY',
        ]
        result = {}

        name, code = response.xpath('//*[@class="col-lg-5"]/h1/text()').extract_first().replace(' ', '').split('(')
        result['name'] = name
        result['code'] = code.replace(')', '')
        result['category'] = response.xpath('//*[@id="page"]/div[2]/div[1]/div[2]/div[1]/span[2]/text()').extract_first()
        result['price'] = response.xpath('//*[@id="price"]/text()').extract_first()
        stock_details = response.xpath('//*[@class="stock_details table table-hover table-striped table-theme"]//td/text()')
        
        extract = False
        found_data = ''
        for stock_detail in stock_details:

            if extract == True:
                result[found_data] = stock_detail.get()
                extract = False

            if stock_detail.get() in selected_data:
                found_data = stock_detail.get()
                extract = True

        yield result

        # scrapy runspider tutorial/spiders/klse_spider.py -o stocks.json 
        # scrapy shell 'https://www.klsescreener.com/v2/stocks/view/7087'
        # git push https://fkxstyle@github.com/fkxstyle/klse.git
