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


# 创建Excel文件
# book = xlwt.Workbook(encoding='utf-8', style_compression=0)
# sheet = book.add_sheet('Starbucks', cell_overwrite_ok=True)
excel_path = 'C:/Users/xieqi/Desktop/starbucks.xls'
open_file = xlrd.open_workbook(excel_path)
file_copy = xlutils.copy.copy(open_file)
sheet = file_copy.get_sheet(0)
# col = ('序数', '姓名','时间','星数','评论')
# for m in range(5):
#     sheet.write(0, m+1, col[m])

driver = webdriver.Chrome(".\chromedriver_win32\chromedriver")
# driver.get('https://www.google.com/maps/search/starbucks/@1.3176312,103.7636757,15z/data=!3m1!4b1')
# url = driver.find_elements(By.CLASS_NAME,'hfpxzc')
# for o in range(3):
# url = driver.find_elements(By.CLASS_NAME, 'hfpxzc')
# shop_link = url[0].get_attribute('href')
r = (10516-1) #门店几所在那行,这里和下面一行url都要改
url='https://www.google.com/maps/place/Starbucks/data=!4m7!3m6!1s0x31da181708e5c9c3:0x12b5061eba19a057!8m2!3d1.3151237!4d103.8944134!16s%2Fg%2F11b6ck1zd5!19sChIJw8nlCBcY2jERV6AZuh4GtRI?authuser=0&hl=en&rclk=1'
driver.get(url)
sleep(1)
# print("门店" + str()) #这里要改
print("地址：", driver.find_element(By.CLASS_NAME,'rogA2c').text)
print(" ")
sheet.write(r, 0, "门店" + str(116)) #这里要改
sheet.write(r+1, 0, "地址：" + driver.find_element(By.CLASS_NAME,'rogA2c').text)

sleep(2)

review_number_element = driver.find_elements(By.CLASS_NAME, 'DkEaL')[0]
reviews_number = review_number_element.text.replace(" reviews", "").replace(",", "")
reviews_num = int(reviews_number)


review = driver.find_element(By.CLASS_NAME,'DkEaL')
review.click()

sleep(2)

js = "var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop=10000"  # getElementsByClassName表示获取class='main'的元素列表，0表示第一个，所以使用的时候要加索引
driver.execute_script(js)
sleep(1)
js = "var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop=0"  # getElementsByClassName表示获取class='main'的元素列表，0表示第一个，所以使用的时候要加索引
driver.execute_script(js)
sleep(1.5)

comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
for i in range(0, math.ceil((reviews_num-20)/10)):
# for i in range(0, 1):
    # 没懂为什么这里框架数量是20，和下面完全不一样，之后再看看
    js = f"var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop={10000+i*5000}"  # getElementsByClassName表示获取class='main'的元素列表，0表示第一个，所以使用的时候要加索引
    driver.execute_script(js)
    sleep(1.5)
    i = i+1
    comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
    if comment[10*(i-3)].text.replace("（由 Google 翻译）", "").replace("…...", "").replace("（原文）", "").strip() == "":
        break

# 因为需要加载出全部所以下面用不了
# WebDriverWait(driver, 20, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME, 'kvMYJc')))

all_reviews = driver.find_elements(By.CLASS_NAME, 'jJc9Ad')
time_frame = driver.find_elements(By.CLASS_NAME, 'rsqaWe')
comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
name = driver.find_elements(By.CLASS_NAME, 'd4r55')
star = driver.find_elements(By.CLASS_NAME, 'kvMYJc')

# k = 0
l = 0
for j in range(len(time_frame)):
    # k = k + 1
    # print("计数：", k)
    if comment[j].text.replace("（由 Google 翻译）", "").replace("…...", "").replace("（原文）", "").strip() == "":
        pass
    else:
        l = l + 1
        comment_text = comment[j].text.replace("（由 Google 翻译）", "").replace("…...", "").replace("（原文）", "").strip()
        print(l)
        print("姓名:", name[j].text)
        print("时间:", time_frame[j].text)
        print("星数:", star[j].get_attribute('aria-label').strip())
        print("评论:", "".join([s for s in comment_text.splitlines() if s.strip()]))
        print(" ")
        review_list = [l, name[j].text, time_frame[j].text, star[j].get_attribute('aria-label').strip(), "".join([s for s in comment_text.splitlines() if s.strip()])]
        for n in range(5):
            sheet.write(r+l-1, n+1, review_list[n])

print("到第几行了：", r+l+2)

#下面也可以用按钮
# driver.back()
# sleep(3)
# driver.back()

save_path = 'C:/Users/xieqi/Desktop/starbucks.xls'
file_copy.save(save_path)