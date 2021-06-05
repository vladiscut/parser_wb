from shutil import which
import scrapy
from scrapy import Request
from scrapy_selenium import SeleniumRequest
from selenium import webdriver
import urllib3
import requests
import selenium


SELENIUM_DRIVER_NAME = 'chrome'
SELENIUM_DRIVER_EXECUTABLE_PATH = r'C:\Users\v.sotnikov\Desktop\парсер\chrome\chromedriver.exe'
SELENIUM_DRIVER_ARGUMENTS = ['--headless']



BOT_NAME = 'beta_wb'

SPIDER_MODULES = ['beta_wb.spiders']
NEWSPIDER_MODULE = 'beta_wb.spiders'



# Crawl responsibly by identifying yourself (and your website) on the user-agent
# USER_AGENT = 'vlad (+http://www.google.com)'

# Obey robots.txt rules
ROBOTSTXT_OBEY = False

DOWNLOAD_DELAY = 0
DOWNLOAD_TIMEOUT = 30
RANDOMIZE_DOWNLOAD_DELAY = True

REACTOR_THREADPOOL_MAXSIZE = 128
CONCURRENT_REQUESTS = 256
CONCURRENT_REQUESTS_PER_DOMAIN = 256
CONCURRENT_REQUESTS_PER_IP = 256

AUTOTHROTTLE_ENABLED = True
AUTOTHROTTLE_START_DELAY = 1
AUTOTHROTTLE_MAX_DELAY = 0.25
AUTOTHROTTLE_TARGET_CONCURRENCY = 128
AUTOTHROTTLE_DEBUG = True

RETRY_ENABLED = True
RETRY_TIMES = 3
RETRY_HTTP_CODES = [500, 502, 503, 504, 400, 401, 403, 404, 405, 406, 407, 408, 409, 410, 429]

PROXY_POOL_ENABLED = True


DOWNLOADER_MIDDLEWARES = {

    'scrapy_proxy_pool.middlewares.ProxyPoolMiddleware': 610,
    'scrapy_proxy_pool.middlewares.BanDetectionMiddleware': 620,
    'scrapy_selenium.SeleniumMiddleware': 800,
    'scrapy.downloadermiddlewares.useragent.UserAgentMiddleware': None,
    'scrapy.spidermiddlewares.referer.RefererMiddleware': 80,
    'scrapy.downloadermiddlewares.retry.RetryMiddleware': 90,
    'scrapy_fake_useragent.middleware.RandomUserAgentMiddleware': 120,
    'scrapy.downloadermiddlewares.cookies.CookiesMiddleware': 130,
    'scrapy.downloadermiddlewares.httpcompression.HttpCompressionMiddleware': 810,
    'scrapy.downloadermiddlewares.redirect.RedirectMiddleware': 900

}

