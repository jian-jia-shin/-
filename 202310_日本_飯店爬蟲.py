from selenium.common import NoSuchElementException
from datetime import datetime  ###顯示當下時間之模組

from selenium.webdriver import Keys

jp_cityName = {1:"札幌",2:"東京",3:"橫濱",4:"名古屋",5:"京都",
            6:"奈良",7:"大阪",8:"神戶",9:"廣島",10:"沖繩"}
city_search = {1:"札幌",2:"東京關東",3:"橫濱市",4:"名古屋",5:"京都府京都市",
            6:"奈良",7:"大阪",8:"神戶",9:"廣島",10:"沖繩縣"}
print(datetime.now())
def choose_city():
    city_num = int(input(f"目前有：{jp_cityName}\n請問您要選哪一個都市?(輸入數字(可配合景點之都市位置選定輸入))"))
    print("========================================\n")
    return city_num
city = choose_city()
if city>=1 and city<=10:
    cityN = city
dYear=str(input("請輸入預定年(西元)："))
dMonth=str(input("請輸入預定月(只有個位數請加0，ex.五月-->輸入時為 05 )："))
dDay=str(input("請輸入預定日(只有個位數請加0，ex.七日-->輸入時為 07 )："))

living_days = int(input("請問要住1晚，或是2晚?(輸入數字1、2)"))
nights=''
if (living_days == 1):
    nights = '1晚'
if (living_days == 2):
    nights = '2晚'

add_closerJR=""
closer_JR = int(input("需要靠近JR車站及地鐵嗎?\n(輸入1:是，輸入0:否)"))
if closer_JR==1:
    add_closerJR = "+近JR與地鐵"
if closer_JR==0:
    add_closerJR = ""

add_with_breakfast=""
with_breakfast = int(input("需要附早餐嗎?\n(輸入1:是，輸入0:否)"))
if with_breakfast==1:
    add_with_breakfast = "+附早餐"
if with_breakfast==0:
    add_with_breakfast = ""

url=""
if living_days ==1:
    url="https://www.google.com/maps/search/日本+"+city_search[city]+"+飯店"+add_closerJR+add_with_breakfast+"/@34.9986042,135.7384476,15z/data=!3m1!4b1!4m7!2m6!5m4!5m3!1s"+dYear+"-"+dMonth+"-"+dDay+"!4m1!1i2!6e3?entry=ttu"
if living_days ==2:
    url="https://www.google.com/maps/search/日本+"+city_search[city]+"+飯店"+add_closerJR+add_with_breakfast+"/@34.9985699,135.7384476,15z/data=!3m1!4b1!4m8!2m7!5m5!5m4!1s"+dYear+"-"+dMonth+"-"+dDay+"!2i2!4m1!1i2!6e3?entry=ttu"

############################################
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
#############################################
hotel_not_score = ["飯店","京都","沖繩","酒店"]
######################################################################################
s = Service(r"C:\chromedriver\chromedriver.exe")  # 驅動器位置(需確認chromedriver.exe放置的位置)
option = webdriver.ChromeOptions()
#option.add_argument("headless")  ##執行爬蟲時不開啟瀏覽器

driver = webdriver.Chrome(service=s, options=option)
driver.implicitly_wait(20)
###################################################
from openpyxl import Workbook

wb_result = Workbook()  ##建立Excel檔案
wb_result.create_sheet(jp_cityName[cityN] + "_飯店")
wb_result.remove(wb_result['Sheet'])
ws_result = wb_result[jp_cityName[cityN] + "_飯店"]
list_top = ["飯店名稱", "簡介", "評分", "討論人數", "目前金額", "訂房網站", "特色", "地點", "入住與退房時間", "連結網址"]
ws_result.append(list_top)
filename2 = dYear+dMonth+dDay+"_"+nights+"_"+jp_cityName[cityN] +'_附近飯店列表(詳細地點).xlsx'

######################################################
driver.get(url)
driver.implicitly_wait(20)
time.sleep(5)

def scroll_times(times):
    for scrolltime in range(0,times+1):
        #driver.implicitly_wait(15)
        #print("scroll ",scrolltime+1," times.")
        click_to_scroll=driver.find_element(By.XPATH, "//div[@id='QA0Szd']/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]")
        #/html/body/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]
        #driver.execute_script("arguments[0].click();", click_to_scroll)
        #print("scrolling...")
        click_to_scroll.send_keys(Keys.END)
        time.sleep(7)
scroll_times(5)
hotel_name = driver.find_elements(By.XPATH, "//div[@class='Z8fK3b']/div[2]/div[1]")
hotel_profile = driver.find_elements(By.XPATH,"//div[@class='y7PRA']/div/div[@class='Z8fK3b']/div[2]/div[4]")
hotel_score = driver.find_elements(By.XPATH,"//div/div/div[2]/div[3]/div[@class='AJB7ye']/span[2]/span[@class='ZkP5Je']/span[1]")
hotel_people = driver.find_elements(By.XPATH,"//div/div/div[2]/div[3]/div[@class='AJB7ye']/span[2]/span[@class='ZkP5Je']/span[2]")
hotel_cost = driver.find_elements(By.XPATH,"//div/div[2]/div[4]/div[3]/div/div[2]/div/div[@class='Tc0rEd fT414d RfDVvc OXYjof']/div[2]")
hotel_web = driver.find_elements(By.XPATH,"//div[@class='Nv2PK THOPZb CpccDe ']/a")


###某一飯店未評分，先檢查數量一致，若無則補字串至景點位置
if len(hotel_name) != len(hotel_score):
    for f in range(len(hotel_name)):
        for m in range(len(hotel_not_score)):
            if hotel_name[f].text==hotel_not_score[m] :
                hotel_score.insert(f,"0.0")
                hotel_people.insert(f,"0")
if len(hotel_name) != len(hotel_cost):
    for f in range(len(hotel_name)):
        for m in range(len(hotel_not_score)):
            if hotel_name[f].text==hotel_not_score[m] :
                hotel_cost.insert(f,"X")

HT_name_list=[]
HT_profile_list=[]
HT_score_list=[]
HT_people_list=[]
HT_web_list=[]

##篩選有評分之飯店
for add1 in range(len(hotel_name)):
    if hotel_score[add1]!="0.0":
        HT_name_list.append(hotel_name[add1].text)
        HT_profile_list.append(hotel_profile[add1].text)
        HT_score_list.append(hotel_score[add1].text)
        HT_people_list.append(hotel_people[add1].text)
        HT_web_list.append(str(hotel_web[add1].get_attribute('href')))
    else:
        continue
#print(HT_web_list)
#####################細部抓取#########################################

fullItem = []
find_feature = "//div[@class='m6QErb WNBkOb ']/div[11]/div[2]/div[2]"
####
#find_price = "//div[@class='m6QErb ']/div[5]/div[2]/div/div[2]/div/div/a/div/div[1]/div/div"
find_price = "//div[@class='m6QErb ']/div[3]/div[3]/div[1]/div/div[2]/div/div/a/div/div[1]/div/div"
####
#find_orderN = "//div[@class='m6QErb ']/div[5]/div[2]/div/div[2]/div/div/a/div/div[1]/span[1]/span"
find_orderN = "//div[@class='m6QErb ']/div[3]/div[3]/div[1]/div/div[2]/div/div/a/div/div[1]/span[1]/span"
####
#find_orderWeb= "//div[@class='m6QErb ']/div[5]/div[2]/div/div[2]/div/div/a"
find_orderWeb= "//div[@class='m6QErb ']/div[3]/div[3]/div[1]/div/div[2]/div/div/a"
####
find_address1 = "//div[@class='m6QErb WNBkOb ']/div[15]/div[3]/button/div/div[2]/div[1]"
find_address2 = "//div[@class='m6QErb WNBkOb ']/div[15]/div[4]/button/div/div[2]/div[1]"
find_address3 = "//div[@class='m6QErb WNBkOb ']/div[17]/div[3]/button/div/div[2]/div[1]"
####
find_checkIn_checkOut1 = "//div[@class='m6QErb WNBkOb ']/div[15]/div[10]/div[1]/div[2]/div[1]"
find_checkIn_checkOut2 = "//div[@class='m6QErb WNBkOb ']/div[15]/div[9]/div[1]/div[2]/div[1]"
find_checkIn_checkOut3 = "//div[@class='m6QErb WNBkOb ']/div[15]/div[8]/div[1]/div[2]/div[1]"
####
#pick_font = "//div[@class='m6QErb WNBkOb ']/div[2]/div/div[1]/div[1]"
#pick_font="//img[@class='Liguzb']"
pick_font = "//div[@id='QA0Szd']/div/div/div[1]/div[2]/div/div[1]"
rest_step_url="https://www.google.com/"

def scroll_step():
    time.sleep(5)
    scroll_click=driver.find_element(By.XPATH, pick_font)
    #driver.execute_script("window.scrollBy(0, 500);")
    #driver.execute_script("arguments[0].click();", scroll_click)
    scroll_click.send_keys(Keys.PAGE_DOWN)
    time.sleep(5)
for go_link in range(len(HT_web_list)):
    driver.get(rest_step_url)
    time.sleep(4)
    url = HT_web_list[go_link]
    driver.get(url)
    driver.implicitly_wait(20)
    time.sleep(4)
    scroll_step()

    try:
        hotelFull_feature = driver.find_element(By.XPATH,find_feature)
        hotelFull_feature = hotelFull_feature.text
    except NoSuchElementException:
        hotelFull_feature = "無顯示特色"

    try:
        hotelFull_price = driver.find_element(By.XPATH, find_price)
        hotelFull_price = hotelFull_price.text.replace("$","").replace(",","")
    except NoSuchElementException:
        hotelFull_price = "9999999999"

    try:
        hotelFull_orderN = driver.find_element(By.XPATH, find_orderN)
        hotelFull_orderN = hotelFull_orderN.text
    except NoSuchElementException:
        hotelFull_orderN = "無訂房資訊"

    try:
        hotelFull_orderWeb = driver.find_element(By.XPATH, find_orderWeb)
        hotelFull_orderWeb = hotelFull_orderWeb.get_attribute('href')
    except NoSuchElementException:
        hotelFull_orderWeb = "none"

    try:
        hotelFull_address = driver.find_element(By.XPATH,find_address1)
        hotelFull_address = hotelFull_address.text
    except NoSuchElementException:
        try:
            hotelFull_address = driver.find_element(By.XPATH, find_address2)
            hotelFull_address = hotelFull_address.text
        except NoSuchElementException:
            try:
                hotelFull_address = driver.find_element(By.XPATH, find_address3)
                hotelFull_address = hotelFull_address.text
            except NoSuchElementException:
                hotelFull_address = "無顯示地點"

    try:
        hotelFull_checkInOut = driver.find_element(By.XPATH,find_checkIn_checkOut1)
        hotelFull_checkInOut = hotelFull_checkInOut.text
    except NoSuchElementException:
        try:
            hotelFull_checkInOut = driver.find_element(By.XPATH, find_checkIn_checkOut2)
            hotelFull_checkInOut = hotelFull_checkInOut.text
        except NoSuchElementException:
            try:
                hotelFull_checkInOut = driver.find_element(By.XPATH, find_checkIn_checkOut3)
                hotelFull_checkInOut = hotelFull_checkInOut.text
            except NoSuchElementException:
                hotelFull_checkInOut = "無顯示入住、退房時間"

    Items = [hotelFull_price,hotelFull_orderN,hotelFull_feature, hotelFull_address, hotelFull_checkInOut,hotelFull_orderWeb]
    fullItem += [Items]
    Item1=[HT_name_list[go_link],HT_profile_list[go_link],HT_score_list[go_link],HT_people_list[go_link],
           hotelFull_price,hotelFull_orderN,hotelFull_feature, hotelFull_address, hotelFull_checkInOut,hotelFull_orderWeb]
    ws_result.append(Item1)
    wb_result.save(filename2)
    time.sleep(4)


# for row in range(len(HT_name_list)):
#     bank=[HT_name_list[row],HT_profile_list[row],HT_score_list[row],HT_people_list[row]]
#     for addcol in range(0,6):
#         bank += [fullItem[row][addcol]]
#     ws_result.append(bank)
# wb_result.save(filename2)
driver.close()

###製作時間10分鐘(有參考資料)
###程式執行時間5分鐘
import pandas as pd
df = pd.read_excel(filename2, sheet_name=jp_cityName[cityN]+"_飯店")
wb_value = df.sort_values(by=["目前金額","評分", "討論人數"], ascending=[True,False,False])

writer = pd.ExcelWriter(filename2)
wb_value.to_excel(writer, sheet_name=jp_cityName[cityN]+"_飯店", index=False)
writer.save()
print(datetime.now())