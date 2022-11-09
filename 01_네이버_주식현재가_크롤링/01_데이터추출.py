import requests
from bs4 import BeautifulSoup

# 종목코드리스트
codes = [
    '005930', #삼성전자
    '000660', #하이닉스
    '035720'  #카카오
]

for code in codes:
    url = f"https://finance.naver.com/item/sise.naver?code={code}"
    response = requests.get(url)
    html = response.text
    soup = BeautifulSoup(html, 'html.parser')
    price = soup.select_one("#_nowVal").text  # id 값은 앞에 # 붙여줌
    price = price.replace(',', '')  # , 제거
    print(price)
