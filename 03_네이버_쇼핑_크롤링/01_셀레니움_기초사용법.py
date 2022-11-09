# 실행전 주의사항
# 아래 Chrome Drive의 path를  본인이 다운로드한 곳에 맞게 수정

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
# 브라우저 생성
browser = webdriver.Chrome('C:/chromedriver.exe')  #사전에 다운로드 해 둬야 함

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
