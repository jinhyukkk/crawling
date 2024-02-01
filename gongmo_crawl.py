import tkinter as tk
from tkinter import Entry, Label, Button, StringVar
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import re
from datetime import datetime

# SSL 연결 설정 변경
requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += 'HIGH:!DH:!aNULL'

# URL 조회
def get_url():

    global url_list
    url = "https://www.38.co.kr/html/fund/?o=k"

    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # 모든 테이블 선택
        tables = soup.find_all('table')

        # "summary" 속성이 "공모주"인 테이블 선택
        for table in tables:
            summary_attr = table.get('summary')
            if summary_attr and '공모주 청약일정' in summary_attr:
                # 테이블 내의 모든 <tr> 태그 선택
                rows = table.find_all('tr')

                url_list = []
                # 각 <tr> 요소에서 첫 번째 <a> 태그 선택
                for row in rows:
                    first_link = row.find('a')
                    if first_link:
                        targetUrl = first_link.get('href')
                        name = first_link.text
                        url_list.append(targetUrl)
                        print(f"링크 URL: {targetUrl}, 텍스트: {name}")
                    else:
                        print("첫 번째 <a> 태그가 없습니다.")

        return url_list

    except Exception as e:
        print("대기 동안 오류 발생:", e)

# 상세 조회
def get_detail(url):
    # 기업개요  company_name
    # 공모정보  jugansa
    # 청약일정  subDate, refundDate, openDate, publicPrice, predictionRate
    global company_name, subDate, refundDate, openDate, publicPrice, predictionRate, jugansa

    response = requests.get('https://www.38.co.kr'+url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # 모든 테이블 선택
    tables = soup.find_all('table')
    company_name, subDate, refundDate, openDate, publicPrice, predictionRate, jugansa = '', '', '', '', '', '', ''

    # "summary" 속성으로 테이블 선택
    for table in tables:
        summary_attr = table.get('summary')
        if summary_attr:
            if summary_attr == '기업개요':
                company_name = find_data(table, '종목명')
                print(company_name)
            elif summary_attr == '공모정보':
                jugansa = find_data(table, '주간사')
            elif summary_attr == '공모청약일정':
                subDate = find_data(table, '공모청약일')
                # 문자열을 공백을 기준으로 분할
                dates = subDate.split()
                # 첫 번째 날짜 추출
                subDate = dates[0]
                refundDate = find_data(table, '환불일')
                openDate = find_data(table, '상장일')
                publicPrice = find_data(table, '확정공모가')
                numbers = re.findall(r'\d+', publicPrice)
                publicPrice = ''.join(numbers)
                predictionRate = find_data(table, '기관경쟁률')

    return [company_name, subDate, refundDate, openDate, publicPrice, predictionRate, jugansa]

def find_data(table, find_name):
    td_tags = table.find_all('td')

    is_next_td = False  # 플래그 변수
    result_data = ""    # 정보를 저장할 변수
    # 각 <td> 태그에 대해 처리
    for td_tag in td_tags:
        if is_next_td:
            # 정보 획득
            result_data = td_tag.text.strip()
            is_next_td = False

        if td_tag.text.strip() == find_name:
            # 다음 <td> 태그에 있는 정보를 획득하기 위해 플래그 설정
            is_next_td = True
    return result_data

# 엑셀 파일 생성 함수
def create_excel_file():
    data = []
    wb = Workbook()
    ws = wb.active
    ws.append(['종목명', '공모청약일정', '환불일', '상장일', '확정공모가', '기관경쟁률', '주간사'])

    url_list = get_url()
    for url in url_list:
        gongmo_list = get_detail(url)
        print(gongmo_list)
        # 현재 날짜 가져오기
        current_date = datetime.now()

        # 주어진 형식의 날짜 문자열을 날짜 객체로 변환
        given_date = datetime.strptime(gongmo_list[2], "%Y.%m.%d")
        if given_date < current_date:
            break
        ws.append(gongmo_list)

    wb.save("sample.xlsx")

# UI 생성 함수
def create_ui():
    window = tk.Tk()
    window.title("공모주 데이터 크롤러")

    # 윈도우 크기 조정 (너비 x 높이)
    window.geometry("400x200")  # 너비 400, 높이 200

    result_label = Label(window, text="")
    result_label.pack()

    def crawl_and_save():
        create_excel_file()
        result_label.config(text="데이터 크롤링 및 저장 완료")

    crawl_button = Button(window, text="실행", command=crawl_and_save)
    crawl_button.pack()

    window.mainloop()

if __name__ == "__main__":
    create_ui()
