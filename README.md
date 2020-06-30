# KLSE

## Installation
1. Download chrome driver (https://chromedriver.chromium.org/)  
2. Make sure chromedriver is found in OS:  
- add chromedriver.exe into PATH (for Windows)  
- move chromedriver into /usr/local/bin/ (for MacOS)
3. pip install -r requirements.txt
  
## Running the script
scrapy runspider tutorial/spiders/klse_spider.py -o stocks.json  
python tutorial/spiders/morningstar_selenium.py  
python prepare_data.py  
  
or 
  
sh start.sh ~/path/until/virtualenv/activate
