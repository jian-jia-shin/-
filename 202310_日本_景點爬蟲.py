#1236,1237.   1240,1251. 0233,0235.  0249,0255.0307
from selenium.common import NoSuchElementException
from datetime import datetime  ###顯示當下時間之模組

from selenium.webdriver import Keys

print(datetime.now())
jp_cityName = {1:"札幌",2:"東京",3:"橫濱",4:"名古屋",5:"京都",
            6:"奈良",7:"大阪",8:"神戶",9:"廣島",10:"沖繩"}
city_search = {1:"札幌",2:"東京市",3:"横浜市",4:"名古屋",5:"京都府京都市",
            6:"奈良",7:"大阪市",8:"神戶",9:"廣島",10:"沖繩縣"}
city_intro = {
 "札幌":"札幌位於日本北海道的交通中心，更是北海道首府的所在地，擁有豐富的人文及觀光資源，常吸引遊客前往旅行。\n 札幌也是日本啤酒發源地之一，因此而有日本唯一一座啤酒觀光博物館。\n 札幌又常有大型祭典像是2月的冬季札幌雪祭，或是夏季的索朗祭，都是札幌知名的祭典，更有並列日本三大夜景的藻岩山觀景台，吸引旅人駐足於此。"
,"東京":"擁有上千萬人口的東京除了是商業文化的超級融匯地，也是日本與世界接軌的中心點。\n 不論是原宿鮮明的時尚氣息、機器人餐廳、女僕Cafe 或熱血的動漫粉絲（御宅族），都是媒體爭相報導的焦點。\n 瞬息萬變的東京透過對於別具歷史意義的庭園、神社與寺廟的維護，展現守護傳統的精神。"
,"橫濱":"橫濱就在東京都旁邊，是一個美麗的海港城市，擁有壯麗優美的海灣景觀，\n同時洋溢著極具魅力的異國情調，橫濱因港口發展歷史遺留下大量的西洋建築和歐式街道，更有日本最大的中華街(唐人街)，還有許多風景秀麗的公園、購物中心等等，觀光資源十分豐富，相當值得一探"
,"名古屋":"名古屋擁有悠久的歷史。 在日本第二受尊崇的熱田神宮，於1,900年前在這裡建成，供奉著日本三大神器之一的草薙劍。\n 名古屋是首位將軍源賴朝的故鄉，跟愛知縣出身的織田信長、豐臣秀吉、德川家康三英傑也很有淵源，他們和眾多的武士一起奠定了現代日本的堅實基礎。\n 因為是戰國時代的中心，在名古屋以及周邊也發生過歷史上著名的戰役。"
,"京都":"作為日本過去的首都，京都以精緻文化、美食料理和迷人的鄉村風光聞名於世 \n每年數以百萬計的國內外遊客紛紛慕名前來京都市，體驗日本的傳統文化。 不僅清水寺、金閣寺吸引了大批遊客前來觀賞，附近的嵐山竹林也是一個熱門的觀光景點。\n 下榻傳統的日式旅館，泡著療癒身心的溫泉，感受櫻花和斑斕的秋葉帶來的季節變化。"
,"奈良":"奈良縣擁有諸多朝拜聖山、神社靈場和日本古早工藝大阪和京都是古代日本政治和宗教中心，這裡有日本最古老的神社和寺院。\n 奈良公園內，馴鹿自在閒逛，相鄰的東大寺中供奉著世界上最大的青銅佛像。\n 重要的朝聖路線穿過吉野山地區，作為世界遺產地的吉野山，更以春天漫山遍野燦爛綻放的櫻花而著名。"
,"大阪":"大阪是個因地域不同而有不同氣氛的城市\n 南區的心齋橋高級品牌店面林立，難波有百貨公司、地下街及商店街等各種型態的店鋪，五花八門的熱鬧氣氛非常具有大阪風情。 此外，有著歷史淵源的大阪城、住吉大社，以及大阪環球影城和海遊館等王道遊樂景點亦不在少數。"
,"神戶":"神戶的主要景點皆與城市景觀和歷史緊密結合，\n除了展現過往榮光的博物館，更有擁抱高科技與新潮流的時尚藝術展覽會，古今交相輝映。 於1868 年開港的神戶，是日本首批開放國際貿易的港口之一，不久後更迎來了歐洲、美國、中國以及其他地區的移民，先後在城市裡構築出各式各色的異國區；而這些充滿異國風情的住家和商店，則留存至今。"
,"廣島":"廣島人口約280萬人。 整年穩定的氣候是它的特徵。\n 除了被列為世界遺產的嚴島神社和原子彈爆炸遺址之外，還有多樣的觀光景點，是國內外觀光客經常到訪的觀光勝地。 這裡還有什錦燒、拉麵，以及新鮮牡蠣等，長期受到喜愛且只有在當地才能品嘗到的高水準美食。"
,"沖繩":"沖繩縣位於日本最南端，是一個由島嶼組成的島鏈，歷史上曾是一個獨立的王國，擁有獨特的亞熱帶氣候，也是空手道的發源地。\n 探索歷代琉球國王曾居住過的宮殿遺跡和修復城堡，參觀美麗的海灘和海岸線，欣賞令人歎為觀止的珊瑚叢和海底生物。\n 來這裡賞鯨、賽龍舟、觀賞稀有動植物、 體驗海島風情，讓您忘記時間。"
}

city_wd_1 = ["觀光工廠","特色建築"]
city_wd_2 = ["動漫","特色街區","傳統建築","百貨商場"]
city_wd_3 = ["海灣景觀","中華街","風景公園","購物中心"]
city_wd_4 = ["傳統建築","藝術文化","休閒場所"]
city_wd_5 = ["傳統文化","美食料理","鄉村風光","溫泉"]
city_wd_6 = ["神社 寺院","馴鹿","賞櫻花"]
city_wd_7 = ["百貨與商店街","傳統建築","遊樂園"]
city_wd_8 = ["博物館","藝術文化","異國風情"]
city_wd_9 = ["歷史建築","觀光景點","美食料理"]
city_wd_10 = ["歷史建築","參觀海景","賞鯨","琉球群島"]
cityWords=[city_wd_1,city_wd_2,city_wd_3,city_wd_4,city_wd_5,
           city_wd_6,city_wd_7,city_wd_8,city_wd_9,city_wd_10]

city_yd_1 = ["年齡皆可，想體驗或參觀者推薦","年齡皆可，觀賞特色建築推薦"]
city_yd_2 = ["20~30歲，喜歡動漫推薦","年齡皆可，逛街推薦","30~50歲，喜歡觀賞歷史建築推薦","年齡皆可，逛百貨推薦"]
city_yd_3 = ["50歲以上，喜歡觀賞海景推薦","20~30歲，體驗中華風格推薦","50歲以上，喜歡觀賞風景、遊走公園推薦","20~30歲，喜愛購物者推薦"]
city_yd_4 = ["30~50歲，喜歡觀賞傳統建築推薦","30~50歲，喜歡觀賞藝術文化推薦","年齡皆可，逛街推薦"]
city_yd_5 = ["30~50歲，體驗傳統文化推薦","30~50歲，享受京都美食推薦","30~50歲，體驗鄉村風景推薦","30~50歲，喜歡泡溫泉者推薦"]
city_yd_6 = ["30~50歲，喜歡觀賞傳統建築推薦","年齡皆可，喜歡馴鹿推薦","年齡皆可，賞櫻推薦"]
city_yd_7 = ["20~30歲，喜愛購物者推薦","30~50歲，喜歡觀賞傳統建築推薦","20~30歲，想體驗遊樂設施者推薦"]
city_yd_8 = ["30~50歲，喜歡參觀博物館推薦","30~50歲，喜歡觀賞藝術文化推薦","年齡皆可，想體驗異國風情推薦"]
city_yd_9 = ["30~50歲，喜歡觀賞歷史建築推薦","30~50歲，喜歡觀賞當地文化風景推薦","30~50歲，喜歡大阪特色美食推薦"]
city_yd_10 = ["30~50歲，喜歡觀賞歷史建築推薦","50歲以上，喜歡觀賞海景推薦","年齡皆可，想目睹鯨魚風采者推薦","年齡皆可，想尋找琉球推薦的景點"]
cityyears=[city_yd_1,city_yd_2,city_yd_3,city_yd_4,city_yd_5,
           city_yd_6,city_yd_7,city_yd_8,city_yd_9,city_yd_10]

closerJR_search = {0:"",1:"近JR"}

cityN = 0
wordN = 0
JRN = 0
def choose_city():
    city_num = int(input(f"目前有：{jp_cityName}\n請問您要選哪一個都市?(輸入數字)"))
    print("========================================")
    return city_num

for_sure = "n"
while for_sure=="n":
    city = choose_city()
    if city >= 1 and city <= 10:
        print("以下是這個都市的景點關鍵字:")
        for i in range(len(cityWords[int(city) - 1])):
            print(i, ". ", cityWords[int(city) - 1][i],"(",cityyears[int(city) - 1][i], ")\t")
    word = int(input("請問您對哪一個關鍵字有興趣?(輸入數字)"))
    if word >= 0 and word <= len(cityWords[int(city)-1]):
        print()
    closer_JR = int(input("要靠近JR車站嗎?(輸入數字1:是，0:否)"))

    check_sure = str(input("確定嗎?y:是,n:否"))
    if check_sure == "y":
        cityN = city
        wordN = word
        JRN = closer_JR
        for_sure = "y"
        break
    else:
        cityN = 0
        wordN = 0
        JRN = 0

#print(cityN,wordN)
############################################
place_not_score = ["蟻丘之塔（あり塚の塔）","晚上の東京鐵塔","森之宮口噴水廣場","往有馬","大阪城坐船"]
############################################
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time

######################################################################################
s = Service(r"C:\chromedriver\chromedriver.exe")  # 驅動器位置(需確認chromedriver.exe放置的位置)
option = webdriver.ChromeOptions()
#option.add_argument("headless")  ##執行爬蟲時不開啟瀏覽器

driver = webdriver.Chrome(service=s, options=option)
driver.implicitly_wait(4)
#############################################################################
from openpyxl import Workbook
wb_result = Workbook()  ##建立Excel檔案
wb_result.create_sheet(jp_cityName[cityN])
wb_result.remove(wb_result['Sheet'])
ws_result = wb_result[jp_cityName[cityN]]
list_top = ["景點名稱","評分","討論人數","簡介","門票","地址","時間"]
ws_result.append(list_top)
filename = jp_cityName[cityN] +'_日本景點列表(詳細地點).xlsx'
##############################################################################
url = 'https://www.google.com/maps/search/日本+'+city_search[cityN]+"+"+cityWords[cityN-1][wordN]+'+'+closerJR_search[JRN]+'+景點'
driver.get(url)
driver.implicitly_wait(20)
time.sleep(7)

def scroll_times(times):
    for scrolltime in range(0, times + 1):
        #driver.implicitly_wait(10)
        click_to_scroll=driver.find_element(By.XPATH, "/html/body/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]")
        #driver.execute_script("arguments[0].click();", click_to_scroll)
        click_to_scroll.send_keys(Keys.END)
        time.sleep(5)
scroll_times(4)
place_name = driver.find_elements(By.XPATH, "//div[@class='UaQhfb fontBodyMedium']/div[@class='NrDZNb']/div[2]")
place_score = driver.find_elements(By.XPATH, "//div/div/div[2]/div[3]/div/span[2]/span/span[1]")
place_people = driver.find_elements(By.XPATH, "//div/div/div[2]/div[3]/div/span[2]/span/span[2]")
place_profile = driver.find_elements(By.XPATH, "//div/div[2]/div[4]/div[1]/div/div/div[2]/div[4]")
place_web = driver.find_elements(By.XPATH, "//div/a[@class='hfpxzc']")

###某一景點未評分，先檢查數量一致，若無則補字串至景點位置
if len(place_name) != len(place_score):
    for f in range(len(place_name)):
        for m in range(len(place_not_score)):
            if place_name[f].text==place_not_score[m] :
                place_score.insert(f,"0.0")
                place_people.insert(f,"0")
PL_name_list=[]
PL_score_list=[]
PL_people_list=[]
PL_profile_list=[]
PL_web_list=[]
for item in range(len(place_name)):
    pName = place_name[item].text
    if place_score[item] != "0.0":
        pScore = place_score[item].text
    else:
        pScore = place_score[item]
    if place_people[item] != "0":
        pPeople = place_people[item].text.replace(",", "").replace("(", "").replace(")", "")
    else:
        pPeople = place_people[item]
    pProfile = place_profile[item].text
    pWeb = str(place_web[item].get_attribute('href'))
    PL_name_list.append(pName)
    PL_score_list.append(pScore)
    PL_people_list.append(pPeople)
    PL_profile_list.append(pProfile)
    PL_web_list.append(pWeb)
print(len(PL_name_list))
fullItem = []
find_profile = "//div[@class='m6QErb WNBkOb ']/div[6]/button/div[2]/div[1]/div[1]"

find_address1 = "//div[@class='m6QErb WNBkOb ']/div[16]/div[3]/button/div/div[2]/div[1]"
find_address2 = "//div[@class='m6QErb WNBkOb ']/div[15]/div[3]/button/div/div[2]/div[1]"
find_address3 = "//div[@class='m6QErb WNBkOb ']/div[11]/div[3]/button/div/div[2]/div[1]"
find_address4 = "//div[@class='m6QErb WNBkOb ']/div[7]/div[3]/button/div/div[2]/div[1]"
######
find_ticket1 = "//div[@class='m6QErb WNBkOb ']/div[13]/div[2]/div/a/div[1]/div[2]"
find_ticket2 = "//div[@class='m6QErb WNBkOb ']/div[12]/div[2]/div/a/div[1]/div[2]"
find_ticket3 = "//div[@class='m6QErb WNBkOb ']/div[8]/div[2]/div/a/div[1]/div[2]"
find_ticket4 = "//div[@class='m6QErb WNBkOb ']/div[8]/div[2]/div/a/div[1]/div[1]"
######
full_time1_click="//div[@class='m6QErb WNBkOb ']/div[15]/div[5]/div[1]/div[2]/div/span[2]"
full_time1="//div[@class='m6QErb WNBkOb ']/div[15]/div[5]/div[2]/div/table/tbody"
full_time2_click="//div[@class='m6QErb WNBkOb ']/div[16]/div[4]/div[1]/div[2]/div/span[2]"
full_time2="//div[@class='m6QErb WNBkOb ']/div[16]/div[4]/div[2]/div/table/tbody"
full_time4_click="//div[@class='m6QErb WNBkOb ']/div[11]/div[4]/div[1]/div[2]/div/span[2]"
full_time4="//div[@class='m6QErb WNBkOb ']/div[11]/div[4]/div[2]/div/table/tbody"
full_time5_click="//div[@class='m6QErb WNBkOb ']/div[11]/div[5]/div[1]/div[2]/div/span[2]"
full_time5="//div[@class='m6QErb WNBkOb ']/div[11]/div[5]/div[2]/div/table/tbody"
full_time3_click="//div[@class='m6QErb WNBkOb ']/div[7]/div[4]/div[1]/div[2]/div/span[2]"
full_time3="//div[@class='m6QErb WNBkOb ']/div[7]/div[4]/div[2]/div/table/tbody"

full_timeEnt_click="//div[@class='m6QErb WNBkOb ']/div[11]/div[4]/button"
full_timeEnt_click_catch="//div[@class='m6QErb WNBkOb ']/div[2]/div[1]/div[2]/div/table/tbody"
full_timeEnt2_click="//div[@class='m6QErb WNBkOb ']/div[11]/div[5]/button"
full_timeEnt2_click_catch="//div[@class='m6QErb WNBkOb ']/div[2]/div[1]/div[2]/div/table/tbody"

pick_font = "//div[@id='QA0Szd']/div/div/div[1]/div[2]/div/div[1]"

#pick_font="/html/body/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div"
rest_step_url="https://tw.yahoo.com/"
def scroll_step():
    time.sleep(5)
    scroll_click=driver.find_element(By.XPATH, pick_font)
    scroll_click.send_keys(Keys.PAGE_DOWN)
    #driver.execute_script("window.scrollBy(0, 500);")
    time.sleep(5)

for go_link in range(len(PL_name_list)):
    driver.get(rest_step_url)
    time.sleep(4)
    url = PL_web_list[go_link]
    driver.get(url)
    driver.implicitly_wait(20)
    time.sleep(6)
    try:
        placeFull_profile = driver.find_element(By.XPATH, find_profile)
        placeFull_profile = placeFull_profile.text
    except NoSuchElementException:
        placeFull_profile = "無顯示簡介"

    scroll_step()

    try:
        placeFull_ticket = driver.find_element(By.XPATH, find_ticket1)
        placeFull_ticket = placeFull_ticket.text.replace("$","").replace(",","")
    except NoSuchElementException:
        try:
            placeFull_ticket = driver.find_element(By.XPATH, find_ticket2)
            placeFull_ticket = placeFull_ticket.text.replace("$","").replace(",","")
        except NoSuchElementException:
            try:
                placeFull_ticket = driver.find_element(By.XPATH, find_ticket3)
                placeFull_ticket = placeFull_ticket.text.replace("$","").replace(",","")
            except NoSuchElementException:
                try:
                    placeFull_ticket = driver.find_element(By.XPATH, find_ticket4)
                    placeFull_ticket = placeFull_ticket.text.replace("$", "").replace(",", "")
                except NoSuchElementException:
                    placeFull_ticket = "無顯示門票"

    try:
        placeFull_address = driver.find_element(By.XPATH, find_address1)
        placeFull_address = placeFull_address.text
    except NoSuchElementException:
        try:
            placeFull_address = driver.find_element(By.XPATH, find_address2)
            placeFull_address = placeFull_address.text
        except NoSuchElementException:
            try:
                placeFull_address = driver.find_element(By.XPATH, find_address3)
                placeFull_address = placeFull_address.text
            except NoSuchElementException:
                try:
                    placeFull_address = driver.find_element(By.XPATH, find_address4)
                    placeFull_address = placeFull_address.text
                except NoSuchElementException:
                    placeFull_address = "無顯示地址"

    time.sleep(5)
    driver.get(rest_step_url)
    time.sleep(5)
    driver.get(url)
    driver.implicitly_wait(20)
    time.sleep(5)

    def need_scroll():
        scroll_step()

    need_scroll()
    try:

        clickto=driver.find_element(By.XPATH, full_time1_click)
        driver.execute_script("arguments[0].click();", clickto)
        placeFull_time = driver.find_element(By.XPATH, full_time1)
        placeFull_time = placeFull_time.text
    except NoSuchElementException:
        try:
            #need_scroll()
            clickto=driver.find_element(By.XPATH, full_time2_click)
            driver.execute_script("arguments[0].click();", clickto)
            placeFull_time = driver.find_element(By.XPATH, full_time2)
            placeFull_time = placeFull_time.text
        except NoSuchElementException:
            try:
                #need_scroll()
                clickto=driver.find_element(By.XPATH, full_time3_click)
                driver.execute_script("arguments[0].click();", clickto)
                placeFull_time = driver.find_element(By.XPATH, full_time3)
                placeFull_time = placeFull_time.text
            except NoSuchElementException:
                try:
                    #need_scroll()
                    clickto=driver.find_element(By.XPATH, full_time4_click)
                    driver.execute_script("arguments[0].click();", clickto)
                    placeFull_time = driver.find_element(By.XPATH, full_time4)
                    placeFull_time = placeFull_time.text
                except NoSuchElementException:
                    try:
                        #need_scroll()
                        clickto = driver.find_element(By.XPATH, full_time5_click)
                        driver.execute_script("arguments[0].click();", clickto)
                        placeFull_time = driver.find_element(By.XPATH, full_time5)
                        placeFull_time = placeFull_time.text
                    except NoSuchElementException:
                        try:
                            #need_scroll()
                            #scroll_step()
                            clickto=driver.find_element(By.XPATH, full_timeEnt_click)
                            driver.execute_script("arguments[0].click();", clickto)
                            time.sleep(5)
                            placeFull_time = driver.find_element(By.XPATH, full_timeEnt_click_catch)
                            placeFull_time = placeFull_time.text
                        except NoSuchElementException:
                            try:
                                #need_scroll()
                                #scroll_step()
                                clickto=driver.find_element(By.XPATH, full_timeEnt2_click)
                                driver.execute_script("arguments[0].click();", clickto)
                                time.sleep(5)
                                placeFull_time = driver.find_element(By.XPATH, full_timeEnt2_click_catch)
                                placeFull_time = placeFull_time.text
                            except NoSuchElementException:
                                placeFull_time = "無顯示時間"

    # Items = [placeFull_profile, placeFull_ticket, placeFull_address, placeFull_time]
    # fullItem += [Items]
    #print(placeFull_profile+", "+placeFull_ticket+", "+placeFull_address+", "+placeFull_time)
    Items1=[PL_name_list[go_link],PL_score_list[go_link],PL_people_list[go_link],
            placeFull_profile, placeFull_ticket, placeFull_address, placeFull_time]
    ws_result.append(Items1)
    wb_result.save(filename)
    time.sleep(3)
    driver.implicitly_wait(20)

# for row in range(len(PL_name_list)):
#     bank=[PL_name_list[row],PL_score_list[row],PL_people_list[row]]
#     for addcol in range(0,4):
#         bank += [fullItem[row][addcol]]
#     ws_result.append(bank)


driver.close()
print(datetime.now())
###2023-10-11 15:17:43.756393
###2023-10-11 15:53:30.212655
###36mins

###2023-10-11 16:06:15.479783
###2023-10-11 16:38:34.429529
###32mins
