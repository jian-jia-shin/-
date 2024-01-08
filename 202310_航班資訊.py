jp_cityName = {1:"札幌",2:"東京",3:"橫濱",4:"名古屋",5:"京都",
            6:"奈良",7:"大阪",8:"神戶",9:"廣島",10:"沖繩"}
city_search = {1:"札幌",2:"東京都",3:"橫濱市",4:"名古屋",5:"京都府京都市",
            6:"奈良",7:"大阪",8:"神戶",9:"廣島",10:"沖繩島"}

def choose_city():
    city_num = int(input(f"目前有：{jp_cityName}\n請問您要選哪一個都市?(輸入數字)"))
    print("========================================\n")
    return city_num
city = choose_city()
if city>=1 and city<=10:
    cityN = city

dYearS=str(input("請輸入\"去程\"預定年(西元)："))
dMonthS=str(input("請輸入\"去程\"預定月(只有個位數請加0，ex.五月-->輸入時為 05 )："))
dDayS=str(input("請輸入\"去程\"預定日(只有個位數請加0，ex.七日-->輸入時為 07 )："))
dDateS = dYearS+dMonthS+dDayS
dYearE=str(input("請輸入\"回程\"預定年(西元)："))
dMonthE=str(input("請輸入\"回程\"預定月(只有個位數請加0，ex.五月-->輸入時為 05 )："))
dDayE=str(input("請輸入\"回程\"預定日(只有個位數請加0，ex.七日-->輸入時為 07 )："))
dDateE = dYearE+dMonthE+dDayE

###############進入介面，取得個別資料##########################
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
###################Chrome有頭爬蟲##########################################################
s = Service(r"C:\chromedriver\chromedriver.exe")  # 驅動器位置(需確認chromedriver.exe放置的位置)
option = webdriver.ChromeOptions()
#option.add_argument("headless")  ##執行爬蟲時不開啟瀏覽器
prefs = {
    'profile.default_content_setting_values':{
        'notifications':2
    }
}##2024/01/03出現通知許可時，設定為拒絕
driver = webdriver.Chrome(service=s, options=option)
##################firefox無頭爬蟲############################################
# s = Service("C:/geckodriver/geckodriver.exe")
# option = webdriver.FirefoxOptions()
# option.add_argument(argument="--headless")  ##執行爬蟲時不開啟瀏覽器
# option.add_argument(argument="--no-sandbox")#2024/1/8不開瀏覽器爬蟲須追加
# option.add_argument(argument='--disable-gpu')#2024/1/8不開瀏覽器爬蟲須追加
# option.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")##避免被視為機器人
# driver = webdriver.Firefox(service=s, options=option)
##########################################################
url = "https://www.google.com/travel/flights/search?tfs=CBwQAhooEgoyMDI0LTAyLTA4agwIAxIIL20vMGZ0a3hyDAgDEggvbS8wZ3FmeRooEgoyMDI0LTAyLTEwagwIAxIIL20vMGdxZnlyDAgDEggvbS8wZnRreEABSAFwAYIBCwj___________8BmAEB&hl=zh-TW&gl=tw&curr=TWD"
driver.get(url)
driver.implicitly_wait(20)
time.sleep(5)

###選定日期
driver.implicitly_wait(20)
click_data = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[2]/div/div/div[1]/div/div/div[1]/div/div[1]/div/input")
click_data.click()
time.sleep(3)

driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/button/span").click()

input_dateS = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div[1]/div[1]/div[1]/div/input")
input_dateS.click()
input_dateS.send_keys(dYearS+"/"+dMonthS+"/"+dDayS)
time.sleep(5)
driver.implicitly_wait(20)
input_dateE = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div/input")
input_dateE.click()
input_dateE.send_keys(dYearE+"/"+dMonthE+"/"+dDayE)
time.sleep(5)
driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[2]/div/div/div[2]/div/div[3]/div[3]").click()
input_dateFin = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[2]/div/div/div[2]/div/div[3]/div[3]/div/button/span")
time.sleep(5)
input_dateFin.click()
time.sleep(5)
#####選定日本都市
input_city = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[1]/div[4]/div/div/div[1]/div/div/input")
input_city.click()
time.sleep(3)
input_cityC = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[1]/div[6]/div[2]/div[2]/div[1]/div/input")
input_cityC.clear()
input_cityC.send_keys(city_search[cityN])
time.sleep(3)
input_city_ck = driver.find_element(By.XPATH,"/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div[2]/div[1]/div/div[2]/div[1]/div[6]/div[3]/ul/li[1]/div[2]")
input_city_ck.click()
time.sleep(8)

airplane_list = []
#driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
driver.execute_script("window.scrollTo(0,5)")
time.sleep(3)
driver.implicitly_wait(20)
how_many_plane = driver.find_elements(By.XPATH,"//ul/li[1]/div/div[@class='JMc5Xc']")
time.sleep(7)

for step in range(1,len(how_many_plane)+2):
    i = str(step)
    driver.execute_script("window.scrollTo(0,5)")
    time.sleep(7)

    arrow = driver.find_element(By.XPATH,"//ul/li["+i+"]/div/div[3]/div/div/button/div[3]")
    driver.execute_script("arguments[0].click();", arrow)
    time.sleep(5)

    GO_start_time = driver.find_element(By.XPATH,"//ul/li["+i+"]/div/div[4]/div/div[1]/div[4]/div[1]/span[1]/span/span/span")
    GO_start_place = driver.find_element(By.XPATH,"//ul/li["+i+"]/div/div[4]/div/div[1]/div[4]/div[1]/span[3]")
    GO_spend_times = driver.find_element(By.XPATH,"//ul/li["+i+"]/div/div[4]/div/div[1]/div[4]/div[2]")
    GO_end_time = driver.find_element(By.XPATH, "//ul/li["+i+"]/div/div[4]/div/div[1]/div[4]/div[3]/span[1]/span/span/span")
    GO_end_place = driver.find_element(By.XPATH, "//ul/li["+i+"]/div/div[4]/div/div[1]/div[4]/div[3]/span[3]")
    plane_comp = driver.find_element(By.XPATH, "//ul/li["+i+"]/div/div[4]/div/div[1]/div[10]/span[2]")
    GO_plane_classN = driver.find_element(By.XPATH, "//ul/li["+i+"]/div/div[4]/div/div[1]/div[10]/span[10]")
    #price_cost = driver.find_element(By.XPATH,"//c-wiz/div[2]/div[2]/div[3]/ul/li[" + i + "]/div/div[2]/div/div[3]/div[3]/div[1]/div[2]/span")
    #cost = price_cost.text.replace("$","").replace(",","")
    comp = plane_comp.text
    GsTime = GO_start_time.text
    GsPlace = GO_start_place.text
    GspendT = GO_spend_times.text
    GeTime = GO_end_time.text
    GePlace = GO_end_place.text
    GpClass = GO_plane_classN.text

    time.sleep(5)
    next_page=driver.find_element(By.XPATH, "//div[2]/div[2]/div[3]/ul/li["+i+"]/div/div[2]/div/div[2]/div[4]/div[1]")###
    driver.execute_script("arguments[0].click();",next_page)
    time.sleep(5)
    driver.implicitly_wait(20)
    driver.find_element(By.XPATH, "//ul/li[1]/div/div[3]/div/div/button/div[3]").click()
    time.sleep(5)
    driver.execute_script("window.scrollTo(0,5)")
    time.sleep(5)
    driver.implicitly_wait(20)
    BACK_start_time = driver.find_element(By.XPATH, "//ul/li[1]/div/div[4]/div/div[1]/div[4]/div[1]/span[1]/span/span/span")
    BACK_start_place = driver.find_element(By.XPATH, "//ul/li[1]/div/div[4]/div/div[1]/div[4]/div[1]/span[3]")
    BACK_spend_times= driver.find_element(By.XPATH, "//ul/li[1]/div/div[4]/div/div[1]/div[4]/div[2]")
    BACK_end_time = driver.find_element(By.XPATH, "//ul/li[1]/div/div[4]/div/div[1]/div[4]/div[3]/span[1]/span/span/span")
    BACK_end_place = driver.find_element(By.XPATH, "//ul/li[1]/div/div[4]/div/div[1]/div[4]/div[3]/span[3]")
    BACK_plane_classN = driver.find_element(By.XPATH, "//ul/li[1]/div/div[4]/div/div[1]/div[10]/span[10]")
    plane_comp2 = driver.find_element(By.XPATH,"//ul/li[1]/div/div[4]/div/div[1]/div[10]/span[2]")
    BsTime = BACK_start_time.text
    BsPlace = BACK_start_place.text
    BspendT = BACK_spend_times.text
    BeTime = BACK_end_time.text
    BePlace = BACK_end_place.text
    BpClass = BACK_plane_classN.text
    comp2 = plane_comp2.text
    time.sleep(7)
    next_page2=driver.find_element(By.XPATH,"//ul/li/div/div[2]/div/div[3]/div[1]/div")
    driver.execute_script("arguments[0].click();", next_page2)

    ###################
    time.sleep(7)
    driver.execute_script("window.scrollTo(0,5)")
    time.sleep(5)
    driver.implicitly_wait(20)
    order_webs = driver.find_elements(By.XPATH,"//div[@class='zwo4Rd']/div[2]/div[1]/div")
    order_costs = driver.find_elements(By.XPATH,"//div[@class='zwo4Rd']/div[4]/div/div/span")
    order_web = order_webs[0].text
    order_cost = order_costs[0].text.replace("$","").replace(",","")
    air_class = [order_web,order_cost,
                 comp,
                 GsTime,GsPlace,GspendT,GeTime,GePlace,GpClass,
                 comp2,
                 BsTime,BsPlace,BspendT,BeTime,BePlace,BpClass]
    airplane_list.append(air_class)
    driver.back()
    time.sleep(5)
    driver.back()
    time.sleep(5)
print(airplane_list)
############################################
from openpyxl import Workbook
from openpyxl.styles import Alignment
wb = Workbook()  ##建立Excel檔案
sheet_name = city_search[cityN]+"_航班(來回)"
wb.create_sheet(sheet_name)
ws = wb[sheet_name]
ws.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

list_top = ["訂票網站","訂票金額",
            "去程_航空公司",
            "去程_出發時間", "去程_出發地點", "去程_路徑時間", "去程_抵達時間", "去程_抵達地點", "去程_班次",
            "回程_航空公司",
            "回程_出發時間", "回程_出發地點", "回程_路徑時間", "回程_抵達時間", "回程_抵達地點", "回程_班次"
            ]
ws.append(list_top)
for i in range(len(airplane_list)):
    addup=[airplane_list[i][0],airplane_list[i][1],
           airplane_list[i][2],
           airplane_list[i][3],airplane_list[i][4],airplane_list[i][5],airplane_list[i][6],airplane_list[i][7],airplane_list[i][8],
           airplane_list[i][9],
           airplane_list[i][10],airplane_list[i][11],airplane_list[i][12],airplane_list[i][13],airplane_list[i][14],airplane_list[i][15]
           ]
    ws.append(addup)
time.sleep(6)
wb.remove(wb['Sheet'])
filename = dDateS + "-" + dDateE + jp_cityName[cityN] +'_航班(來回)列表.xlsx'
wb.save(filename)

driver.close()