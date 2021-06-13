from bs4 import BeautifulSoup
from selenium import webdriver
import time
import sys
import pandas as pd
import os
import urllib

# 검색어 받기
print("=" *30)
print("     아마존 닷컴의 분야별 Best Seller 상품 정보 추출하기     ")
print("=" *30)
print('\n')

sec = input('''
    1.Amazon Devices & Accessories     2.Amazon Launchpad            3.Appliances
    4.Apps & Games                     5.Arts, Crafts & Sewing       6.Audible Books & Originals
    7.Automotive                       8.Baby                        9.Beauty & Personal Care      
    10.Books                           11.CDs & Vinyl                12.Camera & Photo             
    13.Cell Phones & Accessories       14.Clothing, Shoes & Jewelry  15.Collectible Currencies       
    16.Computers & Accessories         17.Digital Music              18.Electronics                
    19.Entertainment Collectibles      20.Gift Cards                 21.Grocery & Gourmet Food     
    22.Handmade Products               23.Health & Household         24.Home & Kitchen             
    25.Industrial & Scientific         26.Kindle Store               27.Kitchen & Dining           
    28.Magazine Subscriptions          29.Movies & TV                30.Musical Instruments        
    31.Office Products                 32.Patio, Lawn & Garden       33.Pet Supplies               
    34.Prime Pantry                    35.Smart Home                 36.Software                   
    37.Sports & Outdoors               38.Sports Collectibles        39.Tools & Home Improvement   
    40.Toys & Games                    41.Video Games

1.위 분야 중에서 자료를 수집할 분야의 번호를 선택하세요 : ''')
    
cnt = int(input('2. 해당 분야에서 크롤링 할 건수는 몇건입니까? (1-100건 사이로 입력하세요) : '))
while cnt > 100 : 
    cnt = int(input(" =============== 1-100건 사이로 입력하세요. : "))
    
f_dir = input("3. 파일을 저장할 폴더명만 쓰세요(예: C:\\Users\\User\Desktop\PP\snspr) : ")

#파일 이름 지정, 폴더 생성    
now = time.localtime()
s =  '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

if sec == '1' :
      sec_name='Amazon Devices and Accessories'
elif sec =='2' :
      sec_name='Amazon Launchpad'
elif sec =='3' :
      sec_name='Appliances'
elif sec =='4' :
      sec_name='Apps and Games'
elif sec =='5' :
      sec_name='Arts and Crafts and Sewing'
elif sec =='6' :
      sec_name='Audible Books and Originals'        
elif sec =='7' :
      sec_name='Automotive'        
elif sec =='8' :
      sec_name='Baby'
elif sec =='9' :
      sec_name='Beauty and Personal Care'
elif sec =='10' :
      sec_name='Books'
elif sec =='11' :
      sec_name='CDs and Vinyl'
elif sec =='12' :
      sec_name='Camera and Photo'
elif sec =='13' :
      sec_name='Cell Phones and Accessories'
elif sec =='14' :
      sec_name='Clothing and Shoes and Jewelry'
elif sec =='15' :
      sec_name='Collectible Currencies'
elif sec =='16' :
      sec_name='Computers and Accessories'
elif sec =='17' :
      sec_name='Digital Music'
elif sec =='18' :
      sec_name='Electronics'
elif sec =='19' :
      sec_name='Entertainment Collectibles'
elif sec =='20' :
      sec_name='Gift Cards'
elif sec =='21' :
      sec_name='Grocery and Gourmet Food'
elif sec =='22' :
      sec_name='Handmade Products'
elif sec =='23' :
      sec_name='Health and Household'
elif sec =='24' :
      sec_name='Home and Kitchen'
elif sec =='25' :
      sec_name='Industrial and Scientific'
elif sec =='26' :
      sec_name='Kindle Store'
elif sec =='27' :
      sec_name='Kitchen and Dining'
elif sec =='28' :
      sec_name='Magazine Subscriptions'
elif sec =='29' :
      sec_name='Movies and TV'
elif sec =='30' :
      sec_name='Musical Instruments'
elif sec =='31' :
      sec_name='Office Products'
elif sec =='32' :
      sec_name='Patio and Lawn and Garden'
elif sec =='33' :
      sec_name='Pet Supplies'
elif sec =='34' :
      sec_name='Prime Pantry'
elif sec =='35' :
      sec_name='Smart Home'
elif sec =='36' :
      sec_name='Software'
elif sec =='37' :
      sec_name='Sports and Outdoors'
elif sec =='38' :
      sec_name='Sports Collectibles'
elif sec =='39' :
      sec_name='Tools and Home Improvemen'
elif sec =='40' :
      sec_name='Toys and Games'
elif sec =='41' :
      sec_name='Video Games'
else : 
    print('범위 내의 숫자를 입력해주세요 ^^')

os.makedirs(f_dir+s+'-Amazon-'+sec_name)
os.chdir(f_dir+s+'-Amazon-'+sec_name)

ff_dir=f_dir+s+'-Amazon-'+sec_name
ff_name=f_dir+s+'-Amazon-'+sec_name+'\\'+s+'-Amazon-'+sec_name+'.txt'
fc_name=f_dir+s+'-Amazon-'+sec_name+'\\'+s+'-Amazon-'+sec_name+'.csv'
fx_name=f_dir+s+'-Amazon-'+sec_name+'\\'+s+'-Amazon-'+sec_name+'.xls'

#페이지 열기
startt = time.time()

path = "C:\py_temp\chromedriver_90\chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get('https://www.amazon.com/bestsellers?ld=NSGoogle')
time.sleep(3)

#분야별 더보기 클릭
if sec == '1' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[1]/a""").click( )
elif sec == '2' :                    
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[2]/a""").click( )
elif sec == '3' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[3]/a""").click( )
elif sec == '4' : 
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[4]/a""").click( )
elif sec == '5' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[5]/a""").click( )
elif sec == '6' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[6]/a""").click( )
elif sec == '7' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[7]/a""").click( )  
elif sec == '8' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[8]/a""").click( )
elif sec == '9' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[9]/a""").click( )
elif sec == '10' : 
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[10]/a""").click( )
elif sec == '11' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[11]/a""").click( )
elif sec == '12' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[12]/a""").click( )
elif sec == '13' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[13]/a""").click( )
elif sec == '14' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[14]/a""").click( )
elif sec == '15' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[15]/a""").click( )
elif sec == '16' : 
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[16]/a""").click( )
elif sec == '17' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[17]/a""").click( )
elif sec == '18' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[18]/a""").click( )
elif sec == '19' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[19]/a""").click( )
elif sec == '20' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[20]/a""").click( )
elif sec == '21' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[21]/a""").click( )
elif sec == '22' : 
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[22]/a""").click( )
elif sec == '23' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[23]/a""").click( )
elif sec == '24' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[24]/a""").click( )
elif sec == '25' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[25]/a""").click( )
elif sec == '26' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[26]/a""").click( )
elif sec == '27' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[27]/a""").click( )
elif sec == '28' : 
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[28]/a""").click( )
elif sec == '29' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[29]/a""").click( )
elif sec == '30' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[30]/a""").click( )
elif sec == '31' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[31]/a""").click( )
elif sec == '32' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[32]/a""").click( )
elif sec == '33' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[33]/a""").click( )
elif sec == '34' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[34]/a""").click( )
elif sec == '35' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[35]/a""").click( )
elif sec == '36' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[36]/a""").click( )
elif sec == '37' : 
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[37]/a""").click( )
elif sec == '38' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[38]/a""").click( )
elif sec == '39' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[39]/a""").click( )
elif sec == '40' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[40]/a""").click( )
elif sec == '41' :
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[41]/a""").click( )
time.sleep(3)

#자동 스크롤
def scroll_down(driver):      
      driver.execute_script("window.scrollBy(0,9300);")
      time.sleep(2)
      
scroll_down(driver)

#이모티콘 -> 아래 딕셔너리로 대체함
bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

#
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
bssection = soup.find('div', id='zg-center-div').find_all('li')
#slist = soup.select('#zg-center-div > #zg-ordered-list')
#bssection = slist[0].find_all('li')

if cnt < 51 :
    
    no = 0
    no2 =[]

    ranking2=[]
    title2=[]
    price2=[]
    review2=[]
    rating2=[]

    img2 = []
    img_dir = ff_dir+"-images" #이미지 저장할 폴더 이름
    os.makedirs(img_dir)
    os.chdir(img_dir) #이미지 저장하기 위해 폴더로 이동해있자
    img_no = 0
    
    for i in bssection: 
        f = open(ff_name, 'a', encoding = 'UTF-8')
        print("=" *30)
        no += 1
        print(" %d 번째 상품 정보입니다. " %no)
        f.write(str(no) + '번째 상품 정보입니다. ')
        
        try: 
            ranking = i.find('span',class_='zg-badge-text').get_text().replace("#","")
        except AttributeError: 
            ranking = ""
         
        print("1. 판매순위 : ", ranking)
        f.write('\n' + '1. 판매순위 : ' + ranking + '\n')
         
        ranking2.append(ranking)
        print('='*30)
        
        try:
            title = i.find('div',class_='p13n-sc-truncated').get_text().replace("\n","")
        except AttributeError:
            title = ""

        print("2. 제품소개 : ", title)
        f.write('2. 제품소개 : '+ title +'\n')
         
        title2.append(title)
        print("=" *30)
        
        try: 
            price = i.find('span','p13n-sc-price').get_text().replace("\n","")
        except AttributeError:
            price = ""

        print("="*30)
        print("3. 가격 : ", price)
        f.write('3. 가격 : ' + price + '\n')
        price2.append(price)
        
        try:
            review = i.find('a','a-size-small a-link-normal').get_text().replace(",","")
        except AttributeError:
            review = ""

        print('='*30)
        print("4. 상품평 갯수 : ", review)
        f.write('4. 상품평 갯수 : ' + review + '\n')
        review2.append(review)
        
        try:
            rating = i.find('span','a-icon-alt').get_text()
        except AttributeError:
            rating = ""

        print('='*30)
        print('5. 상품 평점 : ', rating)
        f.write('5. 상품 평점 : ' + rating + '\n')
        rating2.append(rating)
        
        
        try:
            img = i.find('div','a-section a-spacing-small').find('img')['src']
        except AttributeError :
              continue
        
        img_no += 1
        urllib.request.urlretrieve(img, str(img_no) + '.jpg') #이미지 저장
        time.sleep(1)

        if cnt == img_no :
            break
          
        no2.append(no) 
        time.sleep(2)
        
        if cnt + 1 == no : 
            time.sleep(5)
            f.close()
            break
        
elif cnt >= 51 : #1p 크롤링하고, 2p 넘어가서 51-100 크롤링해야함. 
    
    no = 0
    no2 =[]

    ranking2=[]
    title2=[]
    price2=[]
    review2=[]
    rating2=[]

    img2 = []
    img_dir = ff_dir+"-images"
    os.makedirs(img_dir)
    os.chdir(img_dir)
    img_no = 0
    
    for i in bssection: 
        f = open(ff_name, 'a', encoding = 'UTF-8')
        print("=" *30)
        no += 1
        print(" %d 번째 상품 정보입니다. " %no)
        f.write(str(no) + '번째 상품 정보입니다. ')
        
        try: 
            ranking = i.find('span',class_='zg-badge-text').get_text().replace("#","")
        except AttributeError: 
            ranking = ""
         
        print("1. 판매순위 : ", ranking)
        f.write('\n' + '1. 판매순위 : ' + ranking + '\n')
         
        ranking2.append(ranking)
        print('='*30)
        
        try:
            title = i.find('div',class_='p13n-sc-truncated').get_text().replace("\n","")
        except AttributeError:
            title = ""

        print("2. 제품소개 : ", title)
        f.write('2. 제품소개 : '+ title +'\n')
         
        title2.append(title)
        print("=" *30)
        
        try: 
            price = i.find('span','p13n-sc-price').get_text().replace("\n","")
        except AttributeError:
            price = ""

        print("="*30)
        print("3. 가격 : ", price)
        f.write('3. 가격 : ' + price + '\n')
        price2.append(price)
        
        try:
            review = i.find('a','a-size-small a-link-normal').get_text().replace(",","")
        except AttributeError:
            review = ""

        print('='*30)
        print("4. 상품평 갯수 : ", review)
        f.write('4. 상품평 갯수 : ' + review + '\n')
        review2.append(review)
        
        try:
            rating = i.find('span','a-icon-alt').get_text()
        except AttributeError:
            rating = ""

        print('='*30)
        print('5. 상품 평점 : ', rating)
        f.write('5. 상품 평점 : ' + rating + '\n')
        rating2.append(rating)
        
        
        try:
            img = i.find('div','a-section a-spacing-small').find('img')['src']
        except AttributeError :
              continue
        
        img_no += 1
        urllib.request.urlretrieve(img, str(img_no) + '.jpg')
        time.sleep(1)

        if img_no == 50 :
            break
          
        no2.append(no) 
        time.sleep(2)
        
        if no == 50 :
            time.sleep(5)
            f.close()
            break
    
    #페이지 넘기기
    driver.find_element_by_xpath("""//*[@id="zg-center-div"]/div[2]/div/ul/li[3]/a""").click( )
    time.sleep(3)
    
    for i in bssection: 
        f = open(ff_name, 'a', encoding = 'UTF-8')
        no += 1
        f.write(str(no) + '번째 상품 정보입니다. ')
        
        try: 
            ranking = i.find('span',class_='zg-badge-text').get_text().replace("#","")
        except AttributeError: 
            ranking = ""
    
        print("=" *30)
        print(" %d 번째 상품 정보입니다. " %no)
        ranking50 = int(ranking) + 50
        print("1. 판매순위 : ", str(ranking50))
        f.write('\n' + ' 1. 판매순위 : ' + str(ranking50) + '\n')
        ranking2.append(ranking50)
        
        try:
            title = i.find('div',class_='p13n-sc-truncated').get_text().replace("\n","")
        except AttributeError:
            title = ""
      
        print('='*30)
        print("2. 제품소개 : ", title)
        f.write('2. 제품소개 : '+title+'\n')
        title2.append(title)
        
        try: 
            price = i.find('span','p13n-sc-price').get_text().replace("\n","")
        except AttributeError:
            price = ""
        
        print("="*30)
        print("3. 가격 : ", price)
        f.write('3. 가격 : ' + price + '\n')
        price2.append(price)
        
        try:
            review = i.find('a','a-size-small a-link-normal').get_text().replace(",","")
        except AttributeError:
            review = ""
        
        print('='*30)
        print("4. 상품평 갯수 : ", review)
        f.write('4. 상품평 갯수 : ' + review + '\n')
        review2.append(review)
        
        try:
            rating = i.find('span','a-icon-alt').get_text()
        except AttributeError:
            rating = ""
        
        print('='*30)
        print('5. 상품 평점 : ', rating)
        f.write('5. 상품 평점 : ' + rating + '\n')
        rating2.append(rating)
        
        try:
            img = i.find('div','a-section a-spacing-small').find('img')['src']
        except AttributeError :
              continue
        
        img_no += 1
        urllib.request.urlretrieve(img, str(img_no) + '.jpg') #이미지 저장
        time.sleep(1)

        if cnt == img_no :
            time.sleep(3)
            break
                  
        no2.append(no)
        time.sleep(5)
        
        if cnt + 1 == no : 
            f.close()
            break

driver.close()

#표로 저장
amazon_bestseller = pd.DataFrame()
amazon_bestseller['판매순위']=ranking2
amazon_bestseller['제품소개']=pd.Series(title2)
amazon_bestseller['판매가격']=pd.Series(price2)
amazon_bestseller['상품평 갯수']=pd.Series(review2)
amazon_bestseller['상품평점']=pd.Series(rating2)
        
amazon_bestseller.to_csv(fc_name, encoding="utf-8-sig", index=True)
amazon_bestseller.to_excel(fx_name, index=True)


import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fx_name)
ws = wb.ActiveSheet
    
for r in range(1, cnt+1):  
    cell_num = "C" + str(r+1)
    img_path = img_dir + '\\' + str(r) + '.jpg'
    
    rng = ws.Range(cell_num)
    image = ws.Shapes.AddPicture(img_path, False, True, rng.Left, rng.Top, -1, -1)
    excel.Visible = True
    
    if r == cnt+1 :
        #wb.Close(True)
        #excel.quit()
        break

ws.Columns.AutoFit()
rows_no = cnt+1
ws.Rows('2:%s' %rows_no).RowHeight = 200
wb.Close(SaveChanges=1)
excel.Quit()

endt = time.time()
total_time = endt = startt
print("크롤링이 완료되었습니다. 총 크롤링 소요시간 : %s" %total_time)      