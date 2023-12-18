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


# # 创建Excel文件，若已创建则不需要以下两行
# book = xlwt.Workbook(encoding='utf-8', style_compression=0)
# sheet = book.add_sheet('Starbucks', cell_overwrite_ok=True)


# # 创建Excel文件，打开第一张工作表
excel_path = 'OuputExcelPath.xls' #Excel文件要保存在电脑上的路径
open_file = xlrd.open_workbook(excel_path) #用xlrd模块打开Excel文件
file_copy = xlutils.copy.copy(open_file) #用xlutils模块创建文件的复制
sheet = file_copy.get_sheet(0) #选中Excel文件的第一张工作表，如果还未创建的话用add_sheet，如上

# # 创建表头
# col = ('序数', '姓名','时间','星数','评论')
# for m in range(5):
#     sheet.write(0, m+1, col[m])

# # 浏览器驱动（这里选用Chrome，也可以用其他浏览器）
driver = webdriver.Chrome(".\chromedriver_win32\chromedriver") #driver
# driver.get('https://www.google.com/maps/search/starbucks/@1.3176312,103.7636757,15z/data=!3m1!4b1') #打开google maps上的星巴克搜索页面
# url = driver.find_elements(By.CLASS_NAME,'hfpxzc') #通过classname'hfpxzc'找到每一家门店的链接

# # 失败的for循环，原本想一口气爬每家店，但是退出评论和门店搜索界面下滑两大难点始终解决不了，遂放弃，改为一家家门店点
# for o in range(3):
# url = driver.find_elements(By.CLASS_NAME, 'hfpxzc')
# shop_link = url[0].get_attribute('href')

# # 需要人工操作的地方
r = (10516-1) #门店几所在那行,即每家店的数据开始行，这里和下面一行url都要改，url在href属性中
url='https://www.google.com/maps/place/Starbucks/data=!4m7!3m6!1s0x31da181708e5c9c3:0x12b5061eba19a057!8m2!3d1.3151237!4d103.8944134!16s%2Fg%2F11b6ck1zd5!19sChIJw8nlCBcY2jERV6AZuh4GtRI?authuser=0&hl=en&rclk=1'
driver.get(url) #打开门店
sleep(1) #页面待机1秒钟，目的是让网页加载
# print("门店" + str())
print("地址：", driver.find_element(By.CLASS_NAME,'rogA2c').text) #这行爬取店铺地址
print(" ")
sheet.write(r, 0, "门店" + str(116)) #这里要改，sheet.write将数据输入Excel，r是行数，0是列数，"门店" + str(116)是由前面行列数限定的格子要输入的信息
sheet.write(r+1, 0, "地址：" + driver.find_element(By.CLASS_NAME,'rogA2c').text)

sleep(2) #页面待机2秒钟，后面同

review_number_element = driver.find_elements(By.CLASS_NAME, 'DkEaL')[0] #找到门店的评论数，[0]是因为find_elements的结果是列表，列表没有如下的text属性
reviews_number = review_number_element.text.replace(" reviews", "").replace(",", "") #将评论数以外的文本去除，因为原本是如156 reviews，只需要156
reviews_num = int(reviews_number) #将评论数由文本转为数字


# 找到评论的按钮并点击，进入评论页面
review = driver.find_element(By.CLASS_NAME,'DkEaL')
review.click()

sleep(2)

# 设置滚动条，因为评论是动态加载的，需要不断往下滚动才能取到所有评论
js = "var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop=10000"  # getElementsByClassName表示获取class='main'的元素列表，0表示第一个，所以使用的时候要加索引
driver.execute_script(js)
sleep(1)
js = "var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop=0"  # getElementsByClassName表示获取class='main'的元素列表，0表示第一个，所以使用的时候要加索引
driver.execute_script(js)
sleep(1.5)


# 开始滚动，直到所有评论加载出来
comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
for i in range(0, math.ceil((reviews_num-20)/10)): #滚动math.ceil((reviews_num-20)/10)次以后，所有评论都会被加载出来；math.ceil向上取整
    js = f"var q=document.getElementsByClassName('m6QErb DxyBCb kA9KIf dS8AEf')[0].scrollTop={10000+i*5000}"  # getElementsByClassName表示获取class='main'的元素列表，0表示第一个，所以使用的时候要加索引
    driver.execute_script(js)
    sleep(1.5)
    i = i+1
    comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd')
    if comment[10*(i-3)].text.replace("（由 Google 翻译）", "").replace("…...", "").replace("（原文）", "").strip() == "": #replace去除括号中字段，可替换为英文
        break #当滑动到出现空白评论时，退出循环，不再滑动，节约时间

# 因为需要加载出全部所以下面用不了
# WebDriverWait(driver, 20, 0.5).until(EC.presence_of_element_located((By.CLASS_NAME, 'kvMYJc')))

all_reviews = driver.find_elements(By.CLASS_NAME, 'jJc9Ad') #评论所在大框架
time_frame = driver.find_elements(By.CLASS_NAME, 'rsqaWe') #时间
comment = driver.find_elements(By.CLASS_NAME, 'wiI7pd') #评论内容
name = driver.find_elements(By.CLASS_NAME, 'd4r55') #评论用户
star = driver.find_elements(By.CLASS_NAME, 'kvMYJc') #星数


# 获取时间、评论内容、评论用户、星数信息，并写入Excel
# k = 0
l = 0
for j in range(len(time_frame)):
    # k = k + 1
    # print("计数：", k)
    if comment[j].text.replace("（由 Google 翻译）", "").replace("…...", "").replace("（原文）", "").strip() == "":
        pass #如果评论为空，则不写入
    else:
        l = l + 1
        comment_text = comment[j].text.replace("（由 Google 翻译）", "").replace("…...", "").replace("（原文）", "").strip()
        print(l)
        print("姓名:", name[j].text)
        print("时间:", time_frame[j].text)
        print("星数:", star[j].get_attribute('aria-label').strip())
        print("评论:", "".join([s for s in comment_text.splitlines() if s.strip()])) #因为评论由多段组成，中间有空白行，需要去除
        print(" ")
        review_list = [l, name[j].text, time_frame[j].text, star[j].get_attribute('aria-label').strip(), "".join([s for s in comment_text.splitlines() if s.strip()])]
        for n in range(5):
            sheet.write(r+l-1, n+1, review_list[n]) #写入Excel

print("到第几行了：", r+l+2) #为了上面输入方便加的代码，下一次门店输入需从哪一行开始

#下面也可以用按钮
# driver.back()
# sleep(3)
# driver.back()

save_path = 'C:/Users/xieqi/Desktop/starbucks.xls' #Excel文件路径
file_copy.save(save_path) #文件保存，采用先创建Excel文件复制而后保存在相同路径来覆盖的方式
