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
from .otchet import*


class WB(scrapy.Spider):
    name = 'wb'
    allowed_domains = ['wildberries.ru']
    # start_urls = ['https://www.wildberries.ru/catalog/9131176/detail.aspx?targetUrl=GP']
    global df_obuv
    global df_odezhda
    global df_igrushki
    global df_sumki
    global df_uvelirka
    df_obuv = pd.DataFrame(
        columns=['Артикул', 'Артикул Кари', 'Байер', 'Направление', 'Группа', 'Подгруппа', 'ТМ', 'Название',
                 'Первая цена', 'Текущая цена', 'Отзывов', '+ к отзывам', 'Купили раз', '+к купили раз',
                 'Кол-во звезд', 'Поразмерная дистрибуция', 'Ссылка на товар'])
    df_odezhda = pd.DataFrame(
        columns=['Артикул', 'Артикул Кари', 'Байер', 'Направление', 'Группа', 'Подгруппа', 'ТМ', 'Название',
                 'Первая цена', 'Текущая цена', 'Отзывов', '+ к отзывам', 'Купили раз', '+к купили раз',
                 'Кол-во звезд', 'Поразмерная дистрибуция', 'Ссылка на товар'])
    df_igrushki = pd.DataFrame(
        columns=['Артикул', 'Артикул Кари', 'Байер', 'Направление', 'Группа', 'Подгруппа', 'ТМ', 'Название',
                 'Первая цена', 'Текущая цена', 'Отзывов', '+ к отзывам', 'Купили раз', '+к купили раз',
                 'Кол-во звезд', 'Поразмерная дистрибуция', 'Ссылка на товар'])
    df_sumki = pd.DataFrame(
        columns=['Артикул', 'Артикул Кари', 'Байер', 'Направление', 'Группа', 'Подгруппа', 'ТМ', 'Название',
                 'Первая цена', 'Текущая цена', 'Отзывов', '+ к отзывам', 'Купили раз', '+к купили раз',
                 'Кол-во звезд', 'Поразмерная дистрибуция', 'Ссылка на товар'])
    df_uvelirka = pd.DataFrame(
        columns=['Артикул', 'Артикул Кари', 'Байер', 'Направление', 'Группа', 'Подгруппа', 'ТМ', 'Название',
                 'Первая цена', 'Текущая цена', 'Отзывов', '+ к отзывам', 'Купили раз', '+к купили раз',
                 'Кол-во звезд', 'Поразмерная дистрибуция', 'Ссылка на товар'])
    global df_vse
    df_vse = pd.read_excel(r'C:\Users\v.sotnikov\PycharmProjects\pythonProject8\venv\wb_scrapy\beta_wb\beta_wb\spiders\isxod\все.xlsx')

    def start_requests(self):
        global i
        global df_vse
        for i  in range(len(df_vse)):
            print(i, 'HERE BITCH')
            art = df_vse['Артикул'][i]
            URL_1 = 'https://www.wildberries.ru/catalog/'
            URL_2 = str(art)
            URL_3 = '/detail.aspx?targetUrl=GP'
            global URL
            URL = URL_1 + URL_2 + URL_3
            yield SeleniumRequest( url=URL,callback=self.parse, cb_kwargs={'index':i, 'URL':URL})
            if i == len(df_vse)-1:
                df_obuv.to_excel('all_obuv.xlsx', index=False)
                df_odezhda.to_excel('all_odezhda.xlsx', index=False)
                df_sumki.to_excel('all_sumki.xlsx', index=False)
                df_igrushki.to_excel('all_igrushki.xlsx', index=False)
                df_uvelirka.to_excel('all_uvelirka.xlsx', index=False)
                Otchet()

    def parse(self,response,index,URL):
        global i
        global df_obuv
        global df_odezhda
        global df_igrushki
        global df_sumki
        global df_uvelirka
        count_able=0
        count_disable=0
        driver = response.request.meta['driver']
        print(len(df_vse))
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
        if str('Обувь') in str(df_vse['Направление'][index]):
            df_obuv = df_obuv.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                             'Направление' : df_vse['Направление'].iloc[index],
                             'Группа': df_vse['Категория'].iloc[index],
                             'Подгруппа': df_vse['Подкатегория'].iloc[index],
                             'Первая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/span/del/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Текущая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/div/span/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Отзывов' : str(re.sub('\D', '',str(response.selector.xpath("//*[@id='a-Comments']/text()").extract()))),
                             'ТМ' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[1]/div[1]/span[1]/text()").extract()).strip().replace('[','')),
                             'Купили раз' : str(re.sub('\D', '',str(response.selector.xpath("//p[@class='order-quantity j-orders-count-wrapper']/span[1]/text()").extract()).replace('xa0',''))),
                             'Название' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[1]/div[1]/span[2]/text()').extract())),
                             'Кол-во звезд' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[2]/div[2]/p/span/text()').extract())),
                             'Поразмерная дистрибуция':str(distribution)+'%',
                             'Ссылка на товар': URL},ignore_index=True )

        if str('Одежда') in str(df_vse['Направление'][index]):
            df_odezhda = df_odezhda.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                             'Направление' : df_vse['Направление'].iloc[index],
                             'Группа': df_vse['Категория'].iloc[index],
                             'Подгруппа': df_vse['Подкатегория'].iloc[index],
                             'Первая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/span/del/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Текущая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/div/span/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Отзывов' : str(re.sub('\D', '',str(response.selector.xpath("//*[@id='a-Comments']/text()").extract()))),
                             'ТМ' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[1]/div[1]/span[1]/text()").extract()).strip().replace('[','')),
                             'Купили раз' : str(re.sub('\D', '',str(response.selector.xpath("//p[@class='order-quantity j-orders-count-wrapper']/span[1]/text()").extract()).replace('xa0',''))),
                             'Название' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[1]/div[1]/span[2]/text()').extract())),
                             'Кол-во звезд' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[2]/div[2]/p/span/text()').extract())),
                             'Поразмерная дистрибуция':str(distribution)+'%',
                             'Ссылка на товар': URL},ignore_index=True )

        if str('Игрушки') in str(df_vse['Направление'][index]):
            df_igrushki = df_igrushki.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                             'Направление' : df_vse['Направление'].iloc[index],
                             'Группа': df_vse['Категория'].iloc[index],
                             'Подгруппа': df_vse['Подкатегория'].iloc[index],
                             'Первая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/span/del/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Текущая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/div/span/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Отзывов' : str(re.sub('\D', '',str(response.selector.xpath("//*[@id='a-Comments']/text()").extract()))),
                             'ТМ' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[1]/div[1]/span[1]/text()").extract()).strip().replace('[','')),
                             'Купили раз' : str(re.sub('\D', '',str(response.selector.xpath("//p[@class='order-quantity j-orders-count-wrapper']/span[1]/text()").extract()).replace('xa0',''))),
                             'Название' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[1]/div[1]/span[2]/text()').extract())),
                             'Кол-во звезд' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[2]/div[2]/p/span/text()').extract())),
                             'Поразмерная дистрибуция':str(distribution)+'%',
                             'Ссылка на товар': URL},ignore_index=True )
        if str('Ювелир') in str(df_vse['Направление'][index]):
            df_uvelirka = df_uvelirka.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                             'Направление' : df_vse['Направление'].iloc[index],
                             'Группа': df_vse['Категория'].iloc[index],
                             'Подгруппа': df_vse['Подкатегория'].iloc[index],
                             'Первая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/span/del/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Текущая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/div/span/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Отзывов' : str(re.sub('\D', '',str(response.selector.xpath("//*[@id='a-Comments']/text()").extract()))),
                             'ТМ' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[1]/div[1]/span[1]/text()").extract()).strip().replace('[','')),
                             'Купили раз' : str(re.sub('\D', '',str(response.selector.xpath("//p[@class='order-quantity j-orders-count-wrapper']/span[1]/text()").extract()).replace('xa0',''))),
                             'Название' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[1]/div[1]/span[2]/text()').extract())),
                             'Кол-во звезд' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[2]/div[2]/p/span/text()').extract())),
                             'Поразмерная дистрибуция':str(distribution)+'%',
                             'Ссылка на товар': URL},ignore_index=True )

        if str('Сумк') in str(df_vse['Направление'][index]):
            df_sumki = df_sumki.append( {'Артикул' : df_vse['Артикул'].iloc[index],
                             'Направление' : df_vse['Направление'].iloc[index],
                             'Группа': df_vse['Категория'].iloc[index],
                             'Подгруппа': df_vse['Подкатегория'].iloc[index],
                             'Первая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/span/del/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Текущая цена' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[4]/div[2]/div/div/div/span/text()").extract()).replace(' ','').replace('\\xa0', '').replace('₽','')),
                             'Отзывов' : str(re.sub('\D', '',str(response.selector.xpath("//*[@id='a-Comments']/text()").extract()))),
                             'ТМ' : re.sub('[\[\'\]]','',str(response.selector.xpath("//*[@id='container']/div[1]/div[2]/div[1]/div[1]/span[1]/text()").extract()).strip().replace('[','')),
                             'Купили раз' : str(re.sub('\D', '',str(response.selector.xpath("//p[@class='order-quantity j-orders-count-wrapper']/span[1]/text()").extract()).replace('xa0',''))),
                             'Название' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[1]/div[1]/span[2]/text()').extract())),
                             'Кол-во звезд' : re.sub('[\[\'\]]','',str(response.selector.xpath('//*[@id="container"]/div[1]/div[2]/div[2]/div[2]/p/span/text()').extract())),
                             'Поразмерная дистрибуция':str(distribution)+'%',
                             'Ссылка на товар': str(URL)},ignore_index=True )
        if index==10000:
            df_obuv.to_excel('stages/10000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/10000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/10000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/10000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/10000_uvelirka.xlsx', index=False)
        if index==20000:
            df_obuv.to_excel('stages/20000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/20000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/20000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/20000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/20000_uvelirka.xlsx', index=False)
        if index == 30000:
            df_obuv.to_excel('stages/30000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/30000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/30000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/30000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/30000_uvelirka.xlsx', index=False)
        if index == 40000:
            df_obuv.to_excel('stages/40000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/40000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/40000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/40000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/40000_uvelirka.xlsx', index=False)
        if index == 50000:
            df_obuv.to_excel('stages/50000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/50000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/50000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/50000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/50000_uvelirka.xlsx', index=False)
        if index == 60000:
            df_obuv.to_excel('stages/60000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/60000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/60000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/60000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/60000_uvelirka.xlsx', index=False)
        if index == 70000:
            df_obuv.to_excel('stages/70000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/70000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/70000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/70000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/70000_uvelirka.xlsx', index=False)
        if index == 80000:
            df_obuv.to_excel('stages/80000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/80000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/80000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/80000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/80000_uvelirka.xlsx', index=False)
        if index == 90000:
            df_obuv.to_excel('stages/90000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/90000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/90000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/90000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/90000_uvelirka.xlsx', index=False)
        if index == 100000:
            df_obuv.to_excel('stages/100000_obuv.xlsx', index=False)
            df_odezhda.to_excel('stages/100000_odezhda.xlsx', index=False)
            df_sumki.to_excel('stages/100000_sumki.xlsx', index=False)
            df_igrushki.to_excel('stages/100000_igrushki.xlsx', index=False)
            df_uvelirka.to_excel('stages/100000_uvelirka.xlsx', index=False)











