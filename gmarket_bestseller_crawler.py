from bs4 import BeautifulSoup
from selenium import webdriver
import time
import sys
import pandas as pd
import os
import urllib

print("=" *30)
print("     지마켓 Best Seller 상품 정보 추출하기     ")
print("=" *30)
print('\n')

cnt = int(input('크롤링 할 건수는 몇건입니까? (1-200건 사이로 입력하세요) : '))
while cnt > 200 : 
    cnt = int(input(" =============== 1-200건 사이로 입력하세요. : "))
    
f_dir = input("파일을 저장할 폴더명만 쓰세요(예: C:\\Users\\User\Desktop\PP\snspr) : ")

now = time.localtime()
s =  '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

os.makedirs(f_dir+s+'-Gmarket-')
os.chdir(f_dir+s+'-Gmarket-')

ff_dir=f_dir+s+'-Gmarket-'
ff_name=f_dir+s+'-Gmarket-' + '\\'+s+'-Gmarket-'+'.txt'
fc_name=f_dir+s+'-Gmarket-' + '\\'+s+'-Gmarket-'+'.csv'
fx_name=f_dir+s+'-Gmarket-' + '\\'+s+'-Gmarket-'+'.xls'

startt = time.time()

path = "C:\py_temp\chromedriver_90\chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get('http://corners.gmarket.co.kr/Bestsellers')
time.sleep(3)

height = driver.execute_script("return document.body.scrollHeight")

def scroll_down(driver):      
      driver.execute_script("window.scrollBy(0,%s);" %height)
      time.sleep(15)
      
scroll_down(driver)

bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

no = 0
no2 =[]

ranking2=[]
title2=[]
o_price2=[]
s_price2=[]
sale2=[]

img2 = []
img_dir = ff_dir+"-images" #이미지 저장할 폴더 이름
os.makedirs(img_dir)
os.chdir(img_dir) #이미지 저장하기 위해 폴더로 이동해있자
img_no = 0

#driver.switch_to.frame('best-list')[0]
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
bssection = soup.find_all('div', class_='best-list')[1].find_all('li')

for i in bssection: 
  
    f = open(ff_name, 'a', encoding = 'UTF-8')
    print("=" *30)
    no += 1
    print(" %d 번째 상품 정보입니다. " %no)
    f.write(str(no) + '번째 상품 정보입니다. ' + '\n')
        
    try: 
        ranking = i.find('p',class_='no%s' %no).get_text()
    except AttributeError: 
        ranking = ""
    print("1. 판매순위 : ", ranking)
    f.write('\n' + '1. 판매순위 : ' + ranking + '\n')
    ranking2.append(ranking)
    
    try:
         title = i.find('a',class_='itemname').get_text()
    except AttributeError:
         title = ""
    print("2. 제품소개 : ", title)
    f.write('2. 제품소개 : '+ title +'\n')         
    title2.append(title)
    
    try:
         o_price = i.find('div',class_='o-price').get_text().replace("원", "").replace(",","")
    except AttributeError:
         o_price = ""
    print("3. 원래가격 : ", o_price)
    f.write('3. 원래가격 : '+ o_price +'\n')         
    o_price2.append(o_price2)
        
    try:
        #s_price = i.find('div',class_='s-price').get_text().replace("원", "").replace(",","")
        s_price = i.select_one('div.s-price > strong').get_text().replace("원", "").replace(",","")
    except AttributeError:
        s_price = ''
    print("4. 판매가격 : ", s_price)
    f.write('4. 판매가격 : '+ s_price +'\n')         
    s_price2.append(s_price2)
    
    try:
        #sale = i.find('span',class_='sale').find('em').get_text().replace("%", "")
        sale = i.select_one('div.s-price > span > em').get_text().replace("%", "")
    except AttributeError:
        sale = ''
    print("5. 할인율 : ", sale)
    f.write('5. 할인율 : '+ sale +'\n')         
    sale2.append(sale2)
    print("=" *30)

    
    try:
        img = i.find('div','thumb').find('img')['src']
    except AttributeError :
             continue
 
    time.sleep(3)

    img_no += 1
    urllib.request.urlretrieve(img, str(img_no) + '.jpg') #이미지 저장
    time.sleep(3)
   
    if no == cnt : 
        time.sleep(3)
        print("이미지 수집 완료 ------ 기다려주세요. ")
        f.close()
        break
    
print(" ... ")    


#표로 저장
gmarket_bestseller = pd.DataFrame()
gmarket_bestseller['판매순위']=ranking2
gmarket_bestseller['제품소개']=pd.Series(title2)
gmarket_bestseller['원래가격']=pd.Series(o_price2)
gmarket_bestseller['판매가격']=pd.Series(s_price2)
gmarket_bestseller['할인율']=pd.Series(sale2)
        
gmarket_bestseller.to_csv(fc_name, encoding="utf-8-sig", index=True)
gmarket_bestseller.to_excel(fx_name, index=True)


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

driver.close()
