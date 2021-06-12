from bs4 import BeautifulSoup
from selenium import webdriver
import time
import math
import pandas as pd
import os

#검색어 입력 받기
print("=" *80)
print("네이버 영화 리뷰 및 평점을 수집합니다. ")
print("=" *80)
print("\n")

query_txt = input("1. 크롤링 할 영화의 제목을 입력하세요 : ")
cnt = int(input("2. 크롤링 할 리뷰 건수는 몇 건입니까? : "))
page_cnt = math.ceil(cnt / 10)
f_dir = input("3. 파일을 저장할 폴더명만 쓰세요(예: C:\\Users\\User\Desktop\PP\snspr) : ")

#파일 위치 및 이름 지정
now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

os.makedirs(f_dir+s+'-'+query_txt)
os.chdir(f_dir+s+'-'+query_txt)

ff_name = f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.txt'
fc_name = f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.csv'
fx_name = f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.xls'

#크롬 드라이버로 웹 브라우저 실행
startt = time.time()

path = "C:\py_temp\chromedriver_90\chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get('https://movie.naver.com/')
time.sleep(3)

#영화 제목으로 검색
query = driver.find_element_by_id('ipt_tx_srch')
query.send_keys(query_txt)

#영화 제목으로 검색했을 때 가장 상단에 뜨는 영화 클릭
driver.find_element_by_xpath('/html/body/div/div[2]/div/div/fieldset/div/button').click()
driver.find_element_by_xpath('/html/body/div/div[4]/div/div/div/div/div[1]/ul[2]/li/dl/dt/a').click()

#'평점' 클릭해서 들어가기
driver.find_element_by_link_text("평점").click()

driver.switch_to.frame('pointAfterListIframe')
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

#총 리뷰 건수 < 크롤링 요청 건수일 때 총 리뷰 건수로 동기화하기
'''total_cnt = soup.find('div', class_='score_total').find('strong','total').find_all('em')
total_cnt2 = total_cnt[1].get_text()
total_cnt3 = int(total_cnt2.replace(',',''))

if total_cnt3 < cnt : 
    cnt = total_cnt3
    print("총 리뷰 건수가 %d 건 이므로 %d 건만 수집하겠습니다. " %cnt , cnt)
    page_cnt = math.ceil(cnt / 10)
'''    
# 1 별점, 2 리뷰내용, 3 작성자, 4 작성일자, 5 공감횟수, 6 비공감횟수
no = 0
page_no = 1

score2 =[]
review2 = []
writer2 = []
writed2 = []
thumbs_up2 =[]
thumbs_down2 = []

while (no <= cnt):
    html = driver.page_source
    soup = BeautifulSoup(html,'html.parser')
    scorelist = soup.find('div', class_='score_result').find('ul').find_all('li')
    
    for i in scorelist:       
        f = open(ff_name, 'a', encoding = 'UTF-8')
        print("총 %d 개째입니다. " %no)
        f.write("총 " + str(no) + " 개째입니다. " + '\n')
        print('\n')
        
        score = i.find('div', class_='star_score').find('em').get_text()
        print('1. 별점 : ', score)
        f.write("1. 별점 : " + score + '\n')
        score2.append(score)
        
        content = i.find('div', class_='score_reple')
        review = content.find('p').get_text()
        print("2. 리뷰 내용 : ", review.strip())
        f.write("2. 리뷰 내용 : " + review.strip() + '\n')
        review2.append(review.strip())
        
        #/html/body/div/div/div[5]/ul/li[1]/div[2]/dl/dt/em[1]/a/span
        #/html/body/div/div/div[5]/ul/li[1]/div[2]/dl/dt/em[2]
        writer = content.find_all('em')
        writer3 = writer[0].find('span').get_text()
        print("3. 작성자 : ", writer3)
        f.write("3. 작성자 : " + writer3 + '\n')
        writer2.append(writer3)
        
        writed = content.find_all('em')
        writed3 = writer[1].text
        print("4. 작성일자 : ", writed3)
        f.write("4. 작성일자 : " + writed3 + '\n')
        writed2.append(writed3)
        
        #/html/body/div/div/div[5]/ul/li[1]/div[3]/a[1]/strong
        thumbs = i.find('div', class_='btn_area').find_all('strong')
        thumbs_up = thumbs[0].text
        thumbs_down = thumbs[1].text
        print("5. 공감 : ", thumbs_up, '\n', "6. 비공감 : ", thumbs_down)
        print('\n')
        f.write("5. 공감 : " + thumbs_up + '\n' + "6. 비공감 : " + thumbs_down + '\n' + '='*20 + '\n')
        thumbs_up2.append(thumbs_up)
        thumbs_down2.append(thumbs_down)
        
        time.sleep(5)
        
        no += 1
        if no == cnt : 
            time.sleep(3)
            driver.close()
            f.write("크롤링이 완료되었습니다. ")
            f.close()
            break
        else : 
            continue
        
    page_no += 1
    if page_no <= page_cnt : 
        driver.find_element_by_link_text('%s' %page_no).click()
        time.sleep(3)
    else : 
        break
    
# 표로 만들고 csv, xls로 저장
movie_rating = pd.DataFrame()
movie_rating['별점'] = score2
movie_rating['리뷰내용'] = review2
movie_rating['작성자'] = writer2
movie_rating['작성일자'] = writed3
movie_rating['공감'] = thumbs_up2
movie_rating['비공감'] = thumbs_down2

movie_rating.to_csv(fc_name, encoding='utf-8-sig', index=True)
movie_rating.to_excel(fx_name, index=True)

end = time.time()
total_time = endt = startt
print("크롤링이 완료되었습니다. 총 크롤링 소요시간 : %s" %total_time)