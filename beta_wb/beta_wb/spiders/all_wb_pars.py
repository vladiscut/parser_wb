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
# from .otchet import*


class WB(scrapy.Spider):
    def closed(self, reason):
        df.to_excel(r'global_pars_wb.xlsx', index=False)



    name = 'all_cats_wb'
    allowed_domains = ['wildberries.ru']
    start_urls = ['https://www.wildberries.ru/catalog/9131176/detail.aspx?targetUrl=GP']
    global df
    df = pd.DataFrame(
        columns=['Артикул', 'Направление', 'Группа','Раздел', 'Подгруппа', 'Продавец','Ссылка на товар'])
    global df_vse
    df_vse = pd.read_excel(r'C:\Users\v.sotnikov\PycharmProjects\pythonProject8\venv\wb_scrapy\beta_wb\beta_wb\spiders\isxod\product_links_00.xlsx')

    def start_requests(self):
        global i
        global df_vse
        for i  in range(len(df_vse)):
            print(i, 'HERE BITCH')
            global URL
            URL = df_vse['URL'][i]
            print(URL)
            yield SeleniumRequest( url=URL,callback=self.parse, cb_kwargs={'index':i, 'URL':URL})

    def parse(self,response,index,URL):
        global i
        global df
        driver = response.request.meta['driver']
        print(index, '_________',index)
        # time.sleep(1)
        seller = str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[3]/div[2]/div[10]/p/span[2]/text()').get())
        print(seller)
        if (seller=='None'):
            time.sleep(3)
            seller = str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[3]/div[2]/div[10]/p/span[2]/text()').get())
        df = df.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                         'Направление' : df_vse['Направление'].iloc[index],
                         'Группа': df_vse['Категория'].iloc[index],
                         'Раздел': df_vse['Раздел'].iloc[index],
                         'Подгруппа': df_vse['Подкатегория'].iloc[index],
                         'Продавец': seller,
                         'Ссылка на товар': URL},ignore_index=True )

        if index==100000:
            df.to_excel('stages/100000.xlsx', index=False)
        if index==200000:
            df.to_excel('stages/200000.xlsx', index=False)
        if index == 300000:
            df.to_excel('stages/300000.xlsx', index=False)
        if index == 400000:
            df.to_excel('stages/400000.xlsx', index=False)
        if index == 500000:
            df.to_excel('stages/500000.xlsx', index=False)
        if index == 600000:
            df.to_excel('stages/600000.xlsx', index=False)
        if index == 700000:
            df.to_excel('stages/700000.xlsx', index=False)
        if index == 800000:
            df.to_excel('stages/800000.xlsx', index=False)
        if index == 900000:
            df.to_excel('stages/900000.xlsx', index=False)











