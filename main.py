import openpyxl
import time
import json
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

import pandas as pd


# Load Excel file

df = pd.read_excel('dd.xlsx', skiprows=1)

driver = webdriver.Chrome()
driver.get('http://www.snapir.co.il/leviconvert/PublicConversion.aspx')
data = []
for index, row in df.iterrows():
    dict = {}
    a = str(row.iloc[0])
    b = str(row.iloc[1])
    firstInput = driver.find_element(By.ID, 'txtSearchCodeLeviYats')
    firstInput.send_keys(a)
    firstInput.send_keys(Keys.ENTER)
    time.sleep(0.5)
    secondInput = driver.find_element(By.ID, 'txtSearchCodeLeviDegem')
    secondInput.send_keys(b)
    secondInput.send_keys(Keys.ENTER)
    time.sleep(0.5)
    select_element = driver.find_element(By.NAME, 'ddlYear')
    select = Select(select_element)
    time.sleep(1)
    id_str = f"{int(a):d}-{int(b):04d}"
    options_count = len(select_element.find_elements(By.TAG_NAME, 'option'))
    if options_count == 1:
        dict[id_str] = {}
        dict[id_str] = {
            "data": "null"
        }
    else:
        dict[id_str] = {}         
        for index in range(options_count):
            if index > 0:
                select.select_by_index(index)
                year = select.options[index].text
                dict[id_str][year] = {}
                driver.find_element(By.NAME, "btnConvertCodeRishui").click()
                time.sleep(1)
                try:
                    table = driver.find_element(By.ID, "gvLeviList")
                    trs = table.find_elements(By.TAG_NAME, "tr")
                    td_elements = trs[1].find_elements(By.TAG_NAME, "td")
                except NoSuchElementException:
                    td_elements = None
                if td_elements:
                    dict[id_str][year] = {
                        "code": td_elements[0].text,
                        "description": td_elements[1].text,
                        "commercial_status": td_elements[2].text
                    }
                else:
                    dict[id_str][year] = {
                        "data": "null"
                    }
    data.append(dict)
    driver.find_element(By.NAME, "btnClean").click()
with open("data.json", "w", encoding="utf-8") as outfile:
    json.dump(data, outfile, ensure_ascii=False, indent=4)
driver.close()

