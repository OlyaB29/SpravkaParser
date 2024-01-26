import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

chrome_options = Options()
chrome_options.add_argument("--headless=new")
driver = webdriver.Chrome(options=chrome_options)

base_url = "https://spravka.by"
url = "https://spravka.by/organizations"


def save_to_excel(sheet, title, array, row):
    ds = pd.Series(title)
    df = pd.DataFrame(array, index=pd.RangeIndex(start=1, stop=len(array) + 1))
    with pd.ExcelWriter("./spravka.xlsx", mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
        ds.to_excel(writer, sheet_name=sheet, startrow=row, index=False, header=False)
        df.to_excel(writer, sheet_name=sheet, startrow=row + 1)


def parse_sub_category(url):
    companies = []
    driver.get(url)
    try:
        while True:
            button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="companies"]/div[2]/button'))
            )
            button.click()
    except:
        time.sleep(5)
        html = BeautifulSoup(driver.page_source, 'html.parser')
        all = html.find_all('div', class_="items-widget-item")
        for item in all:
            title = item.find('div', class_="title").find('span').text.strip()
            meta = item.find('a').find_all('meta')
            phones = " "
            address = ""
            email = ""
            site = ""
            for el in meta:
                if el.get('itemprop') == "telephone":
                    phones += el.get('content') + "\n"
                elif el.get('itemprop') == "email":
                    email = el.get('content')
            info = item.find_all('div', class_="info")
            for el in info:
                if el.get('itemprop') == "address":
                    address = el.text.strip()
                else:
                    span = el.find('span')
                    if span:
                        site = span.text.strip()
            link = item.find('a').find('link')['href']

            companies.append(
                {'Наименование': title, 'Адрес': address, 'Телефоны': phones[:-1], 'Электронная почта': email, 'Сайт': site,
                 'Ссылка': link})
    print(len(companies))
    return companies


def run_parsing():
    resp = requests.get(url)
    categories = BeautifulSoup(resp.content, 'html.parser').find_all('h2')

    for category in categories[6:7]:
        a = category.find('a')
        title = a.text.strip()
        category_url = base_url + a['href']
        category_resp = requests.get(category_url)
        sub_categories = BeautifulSoup(category_resp.content, 'html.parser').find_all('div', class_="widget-item")

        startrow = 0
        for sub_category in sub_categories:
            sub_a = sub_category.find('a')
            sub_title = sub_a.text.strip()
            sub_category_url = base_url + sub_a['href']
            companies = parse_sub_category(sub_category_url)
            save_to_excel(title, sub_title, companies, startrow)
            startrow += len(companies) + 3


run_parsing()