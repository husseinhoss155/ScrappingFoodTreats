import os
import shutil
import numpy as np
from pathlib import Path
from docx import Document
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import pandas as pd
import time
from selenium.webdriver.chrome.options import Options

driver = webdriver.Chrome('chromedriver.exe',desired_capabilities={'pageLoadStrategy':'eager'})
driver.get('https://www.petloverscentre.com/dog/dog-food-treats/food')
driver.implicitly_wait(1)

#Waiting for a random element to load for preventing getting error loading
element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, 'product-box.col-1-3.tab-col-1-2.mobile-col-1-1'))
        )
brands=[]
newpirces=[]
oldprices=[]
saves=[]
descs=[]
infos=[]
i=0
#
while True:
    i+=1
    if i==10:
        break
    for product in driver.find_elements(By.CLASS_NAME,'product-box.col-1-3.tab-col-1-2.mobile-col-1-1'):
        brand = product.find_element(By.CLASS_NAME,'brand-name').text
        brands.append(brand)
        descs.append(driver.find_element(By.CLASS_NAME,'prod-name').text.replace(brand,'').strip())
        prices = driver.find_element(By.CLASS_NAME,'prod-price').find_elements(By.XPATH,'.//*')
        oldprices.append(prices[0].text)
        newpirces.append(prices[1].text)
        saves.append(prices[2].text.split(' ')[1])
        infos.append(product.find_element(By.XPATH,'.//ul[@class="info"]/li').text.replace('*',''))
    # Navigating to next page if possible
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 '//*[@id="maincontent"]/div[1]/div[3]/div[7]/ul[3]/li/a'))
        )
        element.click()
    except:
        break
driver.quit()
lst = []
lst.append(brands)
lst.append(newpirces)
lst.append(oldprices)
lst.append(saves)
lst.append(descs)
lst.append(infos)
lst = (np.array(lst).transpose()).tolist()

df = pd.DataFrame(data=lst,columns=['Brand name','New price','Old price','Saving','Description','Info'])

# Exporting the dataframe in excel file with auto adjusting excel columns
with pd.ExcelWriter("Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="MySheet")
    auto_adjust_xlsx_column_width(df, writer, sheet_name="MySheet", margin=0)

# Openeing the excel file
absolutePath = Path('Output.xlsx').resolve()
os.system(f'start Output.xlsx "{absolutePath}"')
