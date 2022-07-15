import os
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import pandas as pd
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://www.amazon.com.au/s?k=smart+door+lock&crid=1HAEK186JU3G1&qid=1657776184&sprefix=smart+%2Caps%2C446&ref=sr_pg_1')

#List of dictionaries for each product
products=[]

#Waiting for the page to load random element
element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//div[@data-component-type="s-search-result"]'))
        )
df = pd.DataFrame()
page=1
while True:
    i=2
    print('Now scraping page number '+str(page))
    page+=1
    while True:
        product={}
        #XPATH for each product
        xpath = '//div[@data-index="'+str(i)+'"]'
        i += 1
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, xpath))
            )
            item = element

            #This is the last section in the page so it means that the products have been scraped in this page
            if 'RELATED' in item.text.split():
                break

            #A middle section that is similar to the products in XPATH but not a product (skip it)
            if item.text.strip() == 'MORE RESULTS':
                continue

            #Scraping Product URL
            try:
                url = item.find_element(By.XPATH,'.//h2/a')
            except:
                try:
                    url = item.find_element(By.XPATH,'.//div/a')
                except:
                    continue
            link = url.get_attribute('href')
            url.click()

            #Waiting for the product image to load and scrape it
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="landingImage"]'))
            )
            img = element.get_attribute('src')

            #Scraping product name and price
            name = driver.find_element(By.ID,'productTitle').text
            try:
                price = '$' + str(driver.find_element(By.XPATH, '//span[@class="a-price-whole"]').text)
            except:
                price='Currently Unavailable'
            product['Name'] = name
            product['Price'] =price
            product['URL'] = link
            product['Image'] = img

            #Scraping technical details from tables in the product page
            tables = driver.find_elements(By.TAG_NAME,'tbody')
            for table in tables:
                try:
                    there_is_th = False
                    tags = table.find_elements(By.XPATH,'.//*')
                    for tag in tags:
                        if tag.tag_name=='th':
                            there_is_th=True
                            break
                    if not there_is_th:
                        for tr in table.find_elements(By.TAG_NAME,'tr'):
                            tds = tr.find_elements(By.TAG_NAME,'td')
                            if tds[0].text=='' or tds[1].text=='':
                                break
                            product[tds[0].text] = tds[1].text
                    else:
                        for tr in table.find_elements(By.TAG_NAME, 'tr'):
                            th = tr.find_element(By.TAG_NAME,'th')
                            td = tr.find_element(By.TAG_NAME,'td')
                            if th.text=='' or td.text=='':
                                break
                            product[th.text] = td.text
                except:
                    break
            #Collecting all the products dictionaries in one list
            products.append(product)

            #Get back to products page
            driver.execute_script("window.history.go(-1)")
        except:
            break

    #Navigating to next page if possible
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 '//a[@class="s-pagination-item s-pagination-next s-pagination-button s-pagination-separator"]'))
        )
        element.click()
    except:
        break

#Converting the dictionaries to one data frame
df = df.append(products,ignore_index=True)
driver.quit()

#Special features column is present no need for that
df.drop(columns='Special feature',inplace=True)

#Dropping two unneccessary columns
for index,row in df.iterrows():
    if row['Are batteries included?']=='':
        continue
    row['Batteries Included?'] = row['Are batteries included?']
    break
df = df.drop(columns=['Are batteries included?','Batteries Included'])

# Exporting the dataframe in excel file with auto adjusting excel columns
with pd.ExcelWriter("Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="MySheet")
    auto_adjust_xlsx_column_width(df, writer, sheet_name="MySheet", margin=0)

# Openeing the excel file
absolutePath = Path('Output.xlsx').resolve()
os.system(f'start Output.xlsx "{absolutePath}"')

