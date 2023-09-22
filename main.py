import openpyxl
import undetected_chromedriver as uc
from selenium_stealth import stealth
from webdriver_manager.chrome import ChromeDriverManager
from scrapy import Selector
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import re
from selenium.webdriver.chrome.options import Options
from GF import convert_to_number, extract_value, getdriver, get_tickers_from_excel, update_excel_with_data, scraping_time_series_graph, getting_yearly_price_from_Scraping_time_series_func, rearrange_string


def fetch_data_from_GF(ticker,driver):
    print(ticker)
    driver.get(f'https://www.google.com/finance/quote/{ticker}?hl=en&window=MAX')
    driver.maximize_window()
    price_data_from_graph = scraping_time_series_graph(driver)
    if price_data_from_graph is not None:
        yearly_price_dict = getting_yearly_price_from_Scraping_time_series_func(price_data_from_graph)
        price_2022 = yearly_price_dict['2022']
        price_2021 = yearly_price_dict['2021']
        price_2020 = yearly_price_dict['2020']
        price_2019 = yearly_price_dict['2019']

    else:
        price_2022 = ''
        price_2021 = ''
        price_2020 = ''
        price_2019 = ''
    
    time.sleep(4)
    response = Selector(text=driver.page_source)
    try:
        driver.execute_script('document.getElementById("annual2").click();')
    except:
        print('error on click')
    
    print('sleep 10')
    time.sleep(7)
    response = Selector(text=driver.page_source)
    market = response.xpath('//div[@class="PxxJne"]/following-sibling::div[@aria-selected="true"]/text()').get()
    founded = response.xpath('(//div[contains(text(),"Founded")]/following-sibling::div/text())[1]').get()
    try:
        currency = response.xpath('//div[@aria-labelledby="key-stats-heading"]/div/span/div[contains(text(),"Market cap")]/parent::span/parent::div/div/text()').getall()[-1]
        try: 
            currency = currency.split()[-1]
        except:
            currency = currency
    except:
        currency = ''
    revenue = response.xpath('(//td/span/div[contains(text(),"Revenue")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    revenue = extract_value(revenue)
    Ebitda = response.xpath('(//td/span/div[contains(text(),"EBITDA")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    Ebitda = extract_value(Ebitda)
    Net_Income = response.xpath('(//td/span/div[contains(text(),"Net income")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    Net_Income = extract_value(Net_Income)
    Net_Profit_margin = response.xpath('(//td/span/div[contains(text(),"Net profit margin")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    if not Net_Profit_margin:
        Net_Profit_margin = response.xpath('(//div[contains(text(),"Net profit margin")])[2]/ancestor::td/following-sibling::td/text()').get()
    Return_On_Assets = response.xpath('(//td/span/div[contains(text(),"Return on assets")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    Price_to_book = response.xpath('(//td/span/div[contains(text(),"Price to book")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    market_Cap = response.xpath('(//div[contains(text(),"Market cap")]/parent::span/following-sibling::div/text())[1]').get()
    market_Cap = extract_value(market_Cap)
    Outstanding_Shares = response.xpath('(//td/span/div[contains(text(),"Shares outstanding")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    Outstanding_Shares = extract_value(Outstanding_Shares)
    Average_daily_vol = response.xpath('(//div[contains(text(),"Avg Volume")]/parent::span/following-sibling::div/text())[1]').get()
    Average_daily_vol = extract_value(Average_daily_vol)
    PE_Ratio = response.xpath('//div[@class="eYanAe"]/div/span/div[contains(text(),"P/E ratio")]/parent::span/following-sibling::div/text()').get()
    Dividend_Yield = response.xpath('(//div[contains(text(),"Dividend yield")]/parent::span/following-sibling::div/text())[1]').get()

    Earnings_per_share = response.xpath('(//td/span/div[contains(text(),"Earnings per share")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
    if not Earnings_per_share:
        Earnings_per_share = response.xpath('(//div[contains(text(),"Earnings per share")])[2]/ancestor::tr/td/text()').get()

    price = response.xpath('//div[@jsname="LXPcOd"]/div/span/div/div/text()').get()

    try:
        driver.execute_script('document.getElementById("option-1").click();')
        time.sleep(2)
        response = Selector(text=driver.page_source)
        Earnings_per_share_2021 = response.xpath('(//td/span/div[contains(text(),"Earnings per share")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
        if not Earnings_per_share_2021:
            Earnings_per_share_2021 = response.xpath('(//div[contains(text(),"Earnings per share")])[2]/ancestor::tr/td/text()').get()
    except:
        print('error on click')
        Earnings_per_share_2021 = ''
    time.sleep(3)
    print('time.sleep 3')
    try:
        driver.execute_script('document.getElementById("option-2").click();')
        time.sleep(2)
        response = Selector(text=driver.page_source)
        Earnings_per_share_2020 = response.xpath('(//td/span/div[contains(text(),"Earnings per share")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
        if not Earnings_per_share_2020:
            Earnings_per_share_2020 = response.xpath('(//div[contains(text(),"Earnings per share")])[2]/ancestor::tr/td/text()').get()
    except:
        print('error on click')
        Earnings_per_share_2020 = ''
    time.sleep(3)
    print('time.sleep 4')
    try:
        driver.execute_script('document.getElementById("option-3").click();')
        time.sleep(2)
        response = Selector(text=driver.page_source)
        Earnings_per_share_2019 = response.xpath('(//td/span/div[contains(text(),"Earnings per share")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
        if not Earnings_per_share_2019:
            Earnings_per_share_2019 = response.xpath('(//div[contains(text(),"Earnings per share")])[2]/ancestor::tr/td/text()').get()
    except:
        print('error on click')
        Earnings_per_share_2019 = ''

    time.sleep(3)
    print('time.sleep 5')
    try:
        driver.execute_script('document.getElementById("option-4").click();')
        time.sleep(2)
        response = Selector(text=driver.page_source)
        Earnings_per_share_2018 = response.xpath('(//td/span/div[contains(text(),"Earnings per share")]/parent::span/parent::td/following-sibling::td)[1]/text()').get()
        if not Earnings_per_share_2018:
            Earnings_per_share_2018 = response.xpath('(//div[contains(text(),"Earnings per share")])[2]/ancestor::tr/td/text()').get()
    except:
        print('error on click')
        Earnings_per_share_2018 = ''

    data = {
        "market": market,
        "founded": founded,
        "currency": currency,
        "Revenue": revenue,
        "Ebitda": Ebitda,
        "Net_income": Net_Income,
        "Net_Profit_margin": Net_Profit_margin,
        "Return_On_Assets": Return_On_Assets,
        "Price_to_book": Price_to_book,
        "market_Cap": market_Cap,
        "Outstanding_Shares": Outstanding_Shares,
        "Average_daily_vol": Average_daily_vol,
        "P/E Ratio": PE_Ratio,
        "Earnings Per Share": Earnings_per_share,
        "Price": price,
        "Dividend_Yield": Dividend_Yield,
        "Earnings_per_share_2021": Earnings_per_share_2021,
        "Earnings_per_share_2020": Earnings_per_share_2020,
        "Earnings_per_share_2019": Earnings_per_share_2019,
        "Earnings_per_share_2018": Earnings_per_share_2018,
        'price_2022': price_2022,
        'price_2021': price_2021,
        'price_2020': price_2020,
        'price_2019': price_2019,
    }

    print(ticker)
    print(data)
    return data

file_path = "C:/Users/shahz/Downloads/GoogleFinance_master_sheet.xlsx"
tickers = get_tickers_from_excel(file_path)
mydriver = getdriver()

print(file_path)
for ticker in tickers:
    myticker = rearrange_string(ticker)
    print(f"{myticker} & {ticker}")
    data = fetch_data_from_GF(myticker,mydriver)
    update_excel_with_data(file_path, data, ticker)
