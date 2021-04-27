import scrapy
from scrapy import Request
from scrapy_selenium import SeleniumRequest
from selenium import webdriver
import urllib3
import selenium
import requests
import time
import re
import json
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import openpyxl
import math

class WB(scrapy.Spider):
    name = 'wb_cat'
    allowed_domains = ['wildberries.ru']

    def start_requests(self):
        global URL
        driver = response.request.meta['driver']
        time.sleep(1)
        driver.find_element_by_xpath("//button[@class='nav-element__burger j-menu-burger-btn']").click()
        time.sleep(2)
        cats = response.selector.xpath("//li[@class='menu-burger__main-list-item menu-burger__main-list-item--subcategory']/a/@href").extract()
        for cat in cats:
            URL = cat
            yield SeleniumRequest(url=URL,callback=self.parse)


    def parse(self,response):
        driver = response.request.meta['driver']
        time.sleep(1)
        driver.find_element_by_xpath("//button[@class='nav-element__burger j-menu-burger-btn']").click()
        time.sleep(2)
        k = response.selector.xpath("//li[@class='menu-burger__main-list-item menu-burger__main-list-item--subcategory']/a/@href").extract_first()
        print(k, '__________________')








