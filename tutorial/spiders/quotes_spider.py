import scrapy


class QuotesSpider(scrapy.Spider):
    name = "quotes"

    def start_requests(self):
        keywords = ' '

        keywords = keywords.replace(' ', '%20')
        urls = [
            # 'https://sg.carousell.com/search/products/?query={}'.format(keywords)
            'https://www.lazada.sg/catalog/?q=pedigree+20kg&_keyori=ss&from=input&spm=a2o42.searchlist.search.go.13f32f31chachg{}'.format(keywords)
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        for quote in response.xpath('//div[@class="_-Z"]/figcaption'):
            yield {
                'title': quote.xpath('./a/div/div/h4/text()').extract_first(),
                'price': quote.xpath('./a/div/dl/dd/text()').extract_first(),
                'href': 'https://sg.carousell.com' + quote.xpath('./a//@href').extract_first()
            }

        # scrapy shell 'https://sg.carousell.com/search/products/?query=test'
