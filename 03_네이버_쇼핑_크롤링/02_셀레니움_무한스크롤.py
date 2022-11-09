# 실행전 주의사항
# 아래 Chrome Drive의 path를  본인이 다운로드한 곳에 맞게 수정

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
# 브라우저 생성
browser = webdriver.Chrome('C://chromedriver.exe')  #사전에 다운로드 해 둬야 함

# 웹사이트 열기
browser.get('https://www.naver.com')

# 
browser.implicitly_wait(10) # 10초 기다림. 로딩 끝나길 기다리기 위해

#쇼핑 메뉴 클릭
browser.find_element(By.CSS_SELECTOR, 'a.nav.shop').click()  #예전엔 find_element_by_css_selector

# 쇼핑 메뉴 클릭 후 기다려 줌
time.sleep(2)

#검색창 클릭 
search = browser.find_element(By.CSS_SELECTOR, 'input._searchInput_search_text_3CUDs')
search.click()
 
#검색어 입력
search.send_keys('아이폰 14')
search.send_keys(Keys.ENTER)

# 무한 스크롤
before_h = browser.execute_script("return window.scrollY")  # java script 실행
while True:
    # 맨아래로 스크롤
    browser.find_element(By.CSS_SELECTOR, 'body').send_keys(Keys.END)

    # 스크롤 사이 페이지 로딩 시간 고려
    time.sleep(1)

    # 스크롤 후 높이
    after_h = browser.execute_script("return window.scrollY")

    # 더이상 스크롤 안되면, 즉 끝까지 갔으면 브레이크
    if after_h == before_h:
        break

    before_h = after_h

# 상품정보 div
items = browser.find_elements(By.CSS_SELECTOR, '.basicList_info_area__TWvzp')

for item in items:
    name = item.find_element(By.CSS_SELECTOR, '.basicList_title__VfX3c').text

    # 가격 정보 없을 경우 처리 위해 try/except 활용
    try:
        price = item.find_element(By.CSS_SELECTOR, '.price_num__S2p_v').text
    except:
        price = "판매중단"
    link = item.find_element(By.CSS_SELECTOR, '.basicList_title__VfX3c > a').get_attribute('href')  # 클래스 바로 아래 있는 a tag
    print(name, price, link)