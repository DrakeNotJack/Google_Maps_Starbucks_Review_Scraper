from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from parsel import Selector
from time import sleep
from bs4 import BeautifulSoup
import unittest
import math
import xlwt
import xlrd
import xlutils.copy


# # Create Excel File
# book = xlwt.Workbook(encoding='utf-8', style_compression=0)
# sheet = book.add_sheet('Starbucks', cell_overwrite_ok=True)


# # Open the first Worksheet
excel_path = 'OuputExcelPath.xls' # The path of the Excel saved
open_file = xlrd.open_workbook(excel_path) # Open Excel file using xlrd mode
file_copy = xlutils.copy.copy(open_file) # Create the copy of Excel file using xlutils mode
sheet = file_copy.get_sheet(0) # Select the first Worksheet

# # Create titles
# col = ('No', 'Name','Time','Stars','Review')
# for m in range(5):
#     sheet.write(0, m+1, col[m])

# # Driver of the browser(using Chrome here)
driver = webdriver.Chrome(".\chromedriver_win32\chromedriver") #driver
# driver.get('https://www.google.com/maps/search/starbucks/@1.3176312,103.7636757,15z/data=!3m1!4b1') #Open the search page of Starbucks stores on Google Maps
# url = driver.find_elements(By.CLASS_NAME,'hfpxzc') # Through classname'hfpxzc' find the url of each store

# # Incomplete for loop, exit the review page & scroll down on the store search page remain to be solved
# for o in range(3):
# url = driver.find_elements(By.CLASS_NAME, 'hfpxzc')
# shop_link = url[0].get_attribute('href')

# # Place need manual work, not necessary if the two problems above were solved
r = (input_row_number)-1) # row number of each record starts
url='https://www.google.com/maps/place/Starbucks/data=!4m7!3m6!1s0x31da181708e5c9c3:0x12b5061eba19a057!8m2!3d1.3151237!4d103.8944134!16s%2Fg%2F11b6ck1zd5!19sChIJw8nlCBcY2jERV6AZuh4GtRI?authuser=0&hl=en&rclk=1'
driver.get(url) # Open the url of the store
sleep(1) # Wait for 1 sec to load the page
# print("Store" + str())
print("Address：", driver.find_element(By.CLASS_NAME,'rogA2c').text) # Crawl the address of the store
print(" ")
sheet.write(r, 0, "Store" + str(input_store_number)) # sheet.write writes data to Excel file，r is row number，0 is column number，"Store" + str(input_store_number) is the cell restricted by the row and column
sheet.write(r+1, 0, "Address：" + driver.find_element(By.CLASS_NAME,'rogA2c').text)

sleep(2)

review_number_element = driver.find_elements(By.CLASS_NAME, 'DkEaL')[0] # Find the review numbers，add [0] since the result of find_elements is list，list doesn't have attribute: text
reviews_number = review_number_element.text.replace(" reviews", "").replace(",", "") # Delete all other text aside the number
reviews_num = int(reviews_number)


# Find the review button and click, enter the review page
review = driver.find_element(By.CLASS_NAME,'DkEaL')
review.click()

sleep(2)

# Set scrollbar, since the reviews load dynamically, need to keep scroll down to crawl all reviews
js = "var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop=10000"  # getElementsByClassName means getting the list of class='main'
driver.execute_script(js)
sleep(1)
js = "var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop=0"
driver.execute_script(js)
sleep(1.5)


# Starting srolling down, until all elements are loaded
comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
for i in range(0, math.ceil((reviews_num-20)/10)): # After scrolling math.ceil((reviews_num-20)/10) times，all reviews will be completely loaded
    js = f"var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop={10000+i*5000}"  
    driver.execute_script(js)
    sleep(1.5)
    i = i+1
    comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
    if comment[10*(i-3)].text.replace("（Translated by Google）", "").replace("…...", "").replace("（Orignal）", "").strip() == "":
        break # When scroll down to blank text reviews, stop scrolling to save time
        
all_reviews = driver.find_elements(By.CLASS_NAME, 'jJc9Ad') # The framework
time_frame = driver.find_elements(By.CLASS_NAME, 'rsqaWe') # Time
comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd') # Content of the review
name = driver.find_elements(By.CLASS_NAME, 'd4r55') # User
star = driver.find_elements(By.CLASS_NAME, 'kvMYJc') # Stars


# Get all the information of the titles and write them into Excel
l = 0
for j in range(len(time_frame)):
    if comment[j].text.replace("（Translated by Google）", "").replace("…...", "").replace("（Orignal）", "").strip() == "":
        pass # Blank review will not be writen into Excel 
    else:
        l = l + 1
        comment_text = comment[j].text.replace("（Translated by Google）", "").replace("…...", "").replace("（Orignal）", "").strip()
        print(l)
        print("Name:", name[j].text)
        print("Time:", time_frame[j].text)
        print("Stars:", star[j].get_attribute('aria-label').strip())
        print("Review:", "".join([s for s in comment_text.splitlines() if s.strip()])) # Since review has several paragraphs, blank lines needs to be deleted
        print(" ")
        review_list = [l, name[j].text, time_frame[j].text, star[j].get_attribute('aria-label').strip(), "".join([s for s in comment_text.splitlines() if s.strip()])]
        for n in range(5):
            sheet.write(r+l-1, n+1, review_list[n]) # Write the information into Excel

# Back to last step
# driver.back()

save_path = 'C:/..' # The path of the Excel file to be saved
file_copy.save(save_path) # Save the file, copying the file and cover it with the new fil
