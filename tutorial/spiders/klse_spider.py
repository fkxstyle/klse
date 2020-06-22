import scrapy


class KlseSpider(scrapy.Spider):
    name = "klse"

   
    def start_requests(self):

        urls = [
            'https://www.klsescreener.com/v2/stocks/view/7087'
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        selected_data = [
            'Market Cap',
            'DY',
        ]
        result = {}

        result['price'] = response.xpath('//*[@id="price"]/text()').extract_first()
        stock_details = response.xpath('//*[@class="stock_details table table-hover table-striped table-theme"]//td/text()')
        
        extract = False
        found_data = ''
        for stock_detail in stock_details:

            if extract == True:
                print('kena')
                result[found_data] = stock_detail.get()
                extract = False

            if stock_detail.get() in selected_data:
                found_data = stock_detail.get()
                extract = True

        yield result

        # scrapy runspider klse_spider.py -o item.json 
        # scrapy shell 'https://www.klsescreener.com/v2/stocks/view/7087'
