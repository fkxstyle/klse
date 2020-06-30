#!/bin/sh

virtualenv_path="$1"
echo "Starting virtualenv"
source $virtualenv_path

echo "Getting latest stocks details..."
scrapy runspider tutorial/spiders/klse_spider.py -o stocks.json  

echo "Downloading Financial Report..."
python tutorial/spiders/morningstar_selenium.py  

echo "Preparing Financial Data..."
python prepare_data.py  