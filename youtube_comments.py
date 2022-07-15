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
from selenium.webdriver.common.keys import Keys
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://www.youtube.com/watch?v=ZbZSe6N_BXs')

element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="container"]/h1/yt-formatted-string'))
        )
body = driver.find_element(By.TAG_NAME,'body')
while True:
    body.send_keys(Keys.PAGE_DOWN)
    try:
        last = driver.find_element(By.XPATH,'//*[@id="contents"]/ytd-comment-thread-renderer[5000]')
        break
    except:
        continue
usernames = driver.find_elements(By.XPATH,'//*[@id="author-text"]/span')
comments = driver.find_elements(By.XPATH,'//*[@id="content-text"]')
u=[]
c=[]
for username in usernames:
    u.append(username.text)
for comm in comments:
    c.append(comm.text)
driver.quit()
lst = []
lst.append(u)
lst.append(c)
lst = (np.array(lst).transpose()).tolist()
df = pd.DataFrame(data=lst,columns=['Name','Comment'])

# Exporting the dataframe in excel file with auto adjusting excel columns
with pd.ExcelWriter("Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="MySheet")
    auto_adjust_xlsx_column_width(df, writer, sheet_name="MySheet", margin=0)

# Openeing the excel file
absolutePath = Path('Output.xlsx').resolve()
os.system(f'start Output.xlsx "{absolutePath}"')

