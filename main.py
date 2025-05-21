# Created by Nikita Sinclair
# Last modified 2025-05-05

# Downloading SDSs from Sigma-Aldrich website from HSIS list. 

# Importing modules
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pymupdf
import urllib.request

import time
import xlsxwriter

import csv
timestart = time.time()
# Enables download
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
# "download.default_directory": "C:\\Users\\nikit\\OneDrive\\Documents\\2025\\ENG4701 FINAL YEAR PROJECT A\\substance_scraper", #Set directory to save your downloaded files.
"download.prompt_for_download": False, #Downloads the file without confirmation.
"download.directory_upgrade": True,
"plugins.always_open_pdf_externally": True #Disable PDF opening.
})
driver = webdriver.Chrome(options=options)

# Opens HSIS file to refer to CAS numbers (abridged file for testing/time saving)
input_file = csv.DictReader(open("HSIS_abridged.csv", encoding="utf8"))
cas_numbers = []

for row in input_file:
        if row["CAS-No."] == '':
                pass
        else:
               cas_numbers.append(row['CAS-No.'])


dictionary = {} # Initialising dictionary for CAS numbers
search_query = "http://sigmaaldrich.com/AU/en?srsltid=AfmBOoq3xZoJGH013S7A6KXD_9wrAXjOtUr4fe1zVJCS9Wt2QXx-jcFo"

for number in cas_numbers:
        driver.get(search_query) # Opens search page
        time.sleep(3)
        searchbar = driver.find_element(By.CLASS_NAME,"jss325") # Finds search bar
        searchbar.click()

        actions = ActionChains(driver).send_keys(number + Keys.ENTER).perform() # Input CAS number into search bar.
        time.sleep(2)

        openproduct = driver.find_elements(By.CLASS_NAME, "MuiTypography-root.MuiLink-root.MuiLink-underlineNone.MuiTypography-colorPrimary")

        actions = ActionChains(driver).click(openproduct[3]).perform() # Opens first product entry
        time.sleep(1)

        try: 
                productname = driver.find_element(By.ID, "product-number")
        except:
                dictionary[number] =  'No information available.'
                continue
        else: 
                driver.execute_script("arguments[0].scrollIntoView();", productname)
                productID = "".join([productname.text])
                time.sleep(0.5)
                opensds = driver.find_element(By.XPATH, "//*[text()='SDS']")
                opensds.click()
                endownload = driver.find_element(By.LINK_TEXT, "English - EN")
                endownload.click() # Downloads SDS
        
                time.sleep(5)
                
                doc = pymupdf.open(productID + '.pdf')
                text = ""
                for page in doc:
                        text+=page.get_text()
                index = text.find("Material: ")

                textfil = ''
                for x in range(len(text)):
                        if x < index:
                                continue
                        else:
                                textfil += text[x]

                indexend = textfil.find("Material tested:")
                textint = ''
                for x in range(indexend-1):
                        textint += textfil[x]
                
                indexstart = textfil.find("Material:")
                indexmat = textfil.find("Minimum")
                textmat = ''
                for x in range((indexstart+len("Material:")+1),indexmat-1):
                        textmat += textfil[x]
                if textmat == "":
                        dictionary[number] =  'No information available.'
                else:
                        dictionary[number] =  textmat

           # Need to add something here to delete PDFs after they have been used.     


# Creating new workbook
workbook = xlsxwriter.Workbook("HSIS_material_recommendations.xlsx")
worksheet = workbook.add_worksheet("Material recommendations")

# Creating headers
worksheet.write(0,0,'CAS Number')
worksheet.write(0,1,'Glove material')

# Writing dictionary to workbook
index = 1
for number in cas_numbers:
    worksheet.write(index, 0, number)
    worksheet.write(index, 1, dictionary[number])
    index += 1
workbook.close()
timeend = time.time()
duration = timeend-timestart
print('Duration: ' + str(duration) + ' seconds')