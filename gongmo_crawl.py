import tkinter as tk
from tkinter import Entry, Label, Button, StringVar
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import re
from datetime import datetime
import os

# SSL 연결 설정 변경
requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += 'HIGH:!DH:!aNULL'

# URL 조회
def get_url():

    global urlList
    url = "https://www.38.co.kr/html/fund/?o=k"

    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # 모든 테이블 선택
        tables = soup.find_all('table')

        # "summary" 속성이 "공모주"인 테이블 선택
        for table in tables:
            summaryAttr = table.get('summary')
            if summaryAttr and '공모주 청약일정' in summaryAttr:
                # 테이블 내의 모든 <tr> 태그 선택
                rows = table.find_all('tr')
                # print(rows)
                urlList = []
                # 각 <tr> 요소에서 첫 번째 <a> 태그 선택
                for row in rows:
                    firstLink = row.find('a')

                    if firstLink:
                        tds = row.find_all('td')

                        gongmoDate = tds[1]
                        gongmoSplit = gongmoDate.text.strip().split('~')
                        gongmoYear = gongmoSplit[0].split('.')[0]
                        gongmoRndDate = gongmoYear + '.' + gongmoSplit[1]
                        gongmoRndDate = datetime.strptime(gongmoRndDate, "%Y.%m.%d")
                        # 현재 날짜 가져오기
                        currentDate = datetime.now()
                        if gongmoRndDate < currentDate:
                            break

                        name = firstLink.text
                        targetUrl = firstLink.get('href')
                        urlList.insert(0, targetUrl)
                        print(f"링크 URL: {targetUrl}, 텍스트: {name}")
                    else:
                        print("첫 번째 <a> 태그가 없습니다.")

        return urlList

    except Exception as e:
        print("대기 동안 오류 발생:", e)

# 상세 조회
def get_detail(url):
    # 기업개요  companyName
    # 공모정보  wishPrice, jugansa
    # 청약일정  subDate, refundDate, openDate, publicPrice, predictionRate
    global companyName, subDate, refundDate, openDate, publicPrice, predictionRate, wishPrice, jugansa

    response = requests.get('https://www.38.co.kr'+url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # 모든 테이블 선택
    tables = soup.find_all('table')
    companyName, subDate, refundDate, openDate, publicPrice, predictionRate, wishPrice, jugansa = '', '', '', '', '', '', '', ''

    # "summary" 속성으로 테이블 선택
    for table in tables:
        summaryAttr = table.get('summary')
        if summaryAttr:
            tdTags = table.find_all('td')
            if summaryAttr == '기업개요':
                # 각 <td> 태그에 대해 처리
                for i in range(len(tdTags) - 1):
                    if tdTags[i].text.strip() == '종목명':
                        companyName = tdTags[i + 1].text.strip()
                print(companyName)
            elif summaryAttr == '공모정보':
                # 각 <td> 태그에 대해 처리
                for i in range(len(tdTags) - 1):
                    if tdTags[i].text.strip() == '희망공모가액':
                        wishPrice = tdTags[i + 1].text.strip()
                    elif tdTags[i].text.strip() == '주간사':
                        jugansa = tdTags[i + 1].text.strip()
            elif summaryAttr == '공모청약일정':
                # 각 <td> 태그에 대해 처리
                for i in range(len(tdTags) - 1):
                    if tdTags[i].text.strip() == '공모청약일':
                        dates = tdTags[i + 1].text.strip().split()
                        # 첫 번째 날짜 추출
                        subDate = dates[0]
                        if subDate:
                            subDate = change_date_format(subDate)
                    elif tdTags[i].text.strip() == '환불일':
                        refundDate = tdTags[i + 1].text.strip()
                        if refundDate:
                            refundDate = change_date_format(refundDate)
                    elif tdTags[i].text.strip() == '상장일':
                        openDate = tdTags[i + 1].text.strip()
                        if openDate:
                            openDate = change_date_format(openDate)
                    elif tdTags[i].text.strip() == '확정공모가':
                        publicPrice = tdTags[i + 1].text.strip()
                        numbers = re.findall(r'\d+', publicPrice)
                        publicPrice = ''.join(numbers)
                        if publicPrice == '':
                            wishPriceSplit = wishPrice.strip().split('~')
                            wishPrice = re.findall(r'\d+', wishPriceSplit[1])
                            publicPrice = ''.join(wishPrice)
                    elif tdTags[i].text.strip() == '기관경쟁률':
                        predictionRate = tdTags[i + 1].text.strip()

    return [companyName, subDate, refundDate, openDate, publicPrice, predictionRate, jugansa]

def change_date_format(date):
    return datetime.strptime(date, "%Y.%m.%d").strftime("%Y. %m. %d")

def find_data(table, findName):
    tdTags = table.find_all('td')

    isNextTd = False  # 플래그 변수
    resultData = ""    # 정보를 저장할 변수
    # 각 <td> 태그에 대해 처리
    for tdTag in tdTags:
        if isNextTd:
            # 정보 획득
            resultData = tdTag.text.strip()
            isNextTd = False

        if tdTag.text.strip() == findName:
            # 다음 <td> 태그에 있는 정보를 획득하기 위해 플래그 설정
            isNextTd = True
    return resultData

# 엑셀 파일 생성 함수
def create_excel_file():
    data = []
    wb = Workbook()
    ws = wb.active
    ws.append(['종목명', '공모청약일정', '환불일', '상장일', '확정공모가', '기관경쟁률', '주간사'])

    urlList = get_url()
    for url in urlList:
        gongmoList = get_detail(url)
        print(gongmoList)
        # 현재 날짜 가져오기
        currentDate = datetime.now()

        # 주어진 형식의 날짜 문자열을 날짜 객체로 변환
        dateStr = gongmoList[2].replace(" ", "")
        givenDate = datetime.strptime(dateStr, "%Y.%m.%d")
        if givenDate < currentDate:
            break
        ws.append(gongmoList)

    wb.save("gongmo.xlsx")

def crawl_and_save():
    create_excel_file()
    # 저장된 엑셀 파일을 시스템의 기본 엑셀 프로그램으로 열기
    os.system('start gongmo.xlsx')

if __name__ == "__main__":
    crawl_and_save()
