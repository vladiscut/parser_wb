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
    name = 'wb'
    allowed_domains = ['wildberries.ru']
    # start_urls = ['https://www.wildberries.ru/catalog/9131176/detail.aspx?targetUrl=GP']
    global df
    df = pd.DataFrame(
        columns=['Артикул', 'Артикул Кари', 'Байер', 'Направление', 'Группа', 'Подгруппа', 'ТМ', 'Название',
                 'Первая цена', 'Текущая цена', 'Отзывов', '+ к отзывам', 'Купили раз', '+к купили раз',
                 'Кол-во звезд', 'Поразмерная дистрибуция'])
    global df_vse
    df_vse = pd.read_excel(r'C:\Users\v.sotnikov\PycharmProjects\parser_wb_categ\beta_wb\beta_wb\spiders\isxod\все.xlsx')

    def start_requests(self):
        global i
        global df_vse
        for i in range(len(df_vse)):
            print(i, 'HERE BITCH')
            art = df_vse['Артикул'][i]
            URL_1 = 'https://www.wildberries.ru/catalog/'
            URL_2 = str(art)
            URL_3 = '/detail.aspx?targetUrl=GP'
            global URL
            URL = URL_1 + URL_2 + URL_3
            yield SeleniumRequest( url=URL,callback=self.parse, cb_kwargs={'index':i})
            # if i==5:
            #     df.to_excel('all_all.xlsx', index=False)
            #     break

    def parse(self,response,index):
        global i
        global df
        count_able=0
        count_disable=0
        driver = response.request.meta['driver']
        # import ipdb; ipdb.set_trace()
        # time.sleep(1)
        # print(driver.page_source)
        # driver.find_element_by_class_name('order-quantity j-orders-count-wrapper').extract()
        for k in response.selector.xpath('//label[@class="j-size active"]/span[1]/text()'):
            print(k.get())
            count_able = count_able + 1
        for k in response.selector.xpath('//label[@class="j-size"]/span[1]/text()'):
            print(k.get())
            count_able = count_able + 1
        for kk in response.selector.xpath('//label[@class="j-size disabled"]/span[1]/text()'):
            print(kk.get())
            count_disable = count_disable + 1
        if count_able!=0:
            distribution = round((count_able/(count_able+count_disable))*100)
        else:
            distribution = 0
        print(index, '_________',index)
        df = df.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                         'Направление' : df_vse['Направление'].iloc[index],
                         'Группа': df_vse['Категория'].iloc[index],
                         'Подгруппа': df_vse['Подкатегория'].iloc[index],
                         'Первая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/span/del/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                         'Текущая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/div/span/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                         'Отзывов' : str(re.sub('\D', '',str(response.selector.xpath("//*[@id='a-Comments']/text()").extract()))),
                         'ТМ' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[1]/div[1]/span[1]/text()").extract()).strip().replace('[','')),
                         'Купили раз' : str(re.sub('\D', '',str(response.selector.xpath("//p[@class='order-quantity j-orders-count-wrapper']/span[1]/text()").extract()))),
                         'Название' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[1]/div[1]/span[2]/text()').extract())),
                         'Кол-во звезд' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[2]/div[2]/p/span/text()').extract())),
                         'Поразмерная дистрибуция':str(distribution)+'%'},ignore_index=True )
        if index==10000:
            df.to_excel('10000.xlsx', index=False)
        if index == 20000:
            df.to_excel('20000.xlsx', index=False)
        if index==len(df_vse)-1:
            df.to_excel('all_all.xlsx', index=False)








