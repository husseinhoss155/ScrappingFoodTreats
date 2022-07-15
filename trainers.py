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

from keyboard import press

from selenium.webdriver.common.keys import Keys
subreddit = input('Enter the needed subreddit:')
option = Options()

option.add_argument("--disable-infobars")
option.add_argument("start-maximized")
option.add_argument("--disable-extensions")

# Pass the argument 1 to allow and 2 to block
option.add_experimental_option(
    "prefs", {"profile.default_content_setting_values.notifications": 1}
)

driver = webdriver.Chrome(
    chrome_options=option, executable_path="chromedriver.exe"
)
driver.get('https://www.reddit.com')
# element = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located(
#                 (By.XPATH, '//*[@id="js-country-field"]'))
#         )
# element.click()
# time.sleep(2)
element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="header-search-bar"]'))
        )
element.send_keys(subreddit)
press('enter')
element.send_keys(Keys.RETURN)
# element = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located(
#                 (By.XPATH, '//*[@id="js-state-field"]'))
#         )
body = driver.find_element(By.TAG_NAME,'body')
while True:
    body.send_keys(Keys.PAGE_DOWN)
    time.sleep(0.1)
    try:
        last = driver.find_element(By.XPATH,'//*[@id="AppRouter-main-content"]/div/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div[100]')
        break
    except:
        continue

firstnames=[]
lastnames=[]
emails=[]
xpath = '//*[@id="AppRouter-main-content"]/div/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div['
i=1
usernames=driver.find_elements(By.CLASS_NAME,'_2tbHP6ZydRpjI44J3syuqC._23wugcdiaj44hdfugIAlnX.oQctV4n0yUb0uiHDdGnmE')
usernames_lst = []
for user in usernames:
    if user.text=='':
        continue
    usernames_lst.append(user.text)
driver.quit()
lst = usernames_lst
df = pd.DataFrame(data=lst,columns=['User Names'])

# Exporting the dataframe in excel file with auto adjusting excel columns
with pd.ExcelWriter("Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="MySheet")
    auto_adjust_xlsx_column_width(df, writer, sheet_name="MySheet", margin=0)

# Openeing the excel file
absolutePath = Path('Output.xlsx').resolve()
os.system(f'start Output.xlsx "{absolutePath}"')



