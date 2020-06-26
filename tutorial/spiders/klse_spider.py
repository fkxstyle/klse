import scrapy


class KlseSpider(scrapy.Spider):
    name = "klse"

   
    def start_requests(self):

        urls = [
            'https://www.klsescreener.com/v2/stocks/view/7087',
            'https://www.klsescreener.com/v2/stocks/view/7088',
            'https://www.klsescreener.com/v2/stocks/view/7089',
            'https://www.klsescreener.com/v2/stocks/view/7090',
            'https://www.klsescreener.com/v2/stocks/view/7091',
        ]
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
                print('kena')
                result[found_data] = stock_detail.get()
                extract = False

            if stock_detail.get() in selected_data:
                found_data = stock_detail.get()
                extract = True

        yield result

        # scrapy runspider tutorial/spiders/klse_spider.py -o item.json 
        # scrapy shell 'https://www.klsescreener.com/v2/stocks/view/7087'
        # git push https://fkxstyle@github.com/fkxstyle/klse.git
