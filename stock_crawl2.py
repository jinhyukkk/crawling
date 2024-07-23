import tkinter as tk
from tkinter import Entry, Label, Button, StringVar
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import os

# 크롤링 함수
def crawl_stock_data(stock_symbol):
    if stock_symbol.isnumeric():
        url = f"https://finance.naver.com/item/main.naver?code={stock_symbol}"
    else:
        url = f"https://finance.yahoo.com/quote/{stock_symbol}?p={stock_symbol}&.tsrc=fin-srch"
    # User-Agent 헤더 설정
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.0.0 Safari/537.36'
    }
    name = stock_symbol
    price = 0
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 404:
            print("404 오류: 페이지를 찾을 수 없습니다.")
        else:
            soup = BeautifulSoup(response.text, 'html.parser')
            # 웹 페이지를 크롤링하고 원하는 작업을 수행합니다.
            if stock_symbol.isnumeric():
                name = soup.select_one('.wrap_company a').text.strip()
                price = soup.select_one('.no_today .blind').text.strip()
            else:
                price_element = soup.select_one('[data-testid="qsp-price"]')
                price = price_element.get('data-value')

            print(f"{name}: {price}")
        return name, price

    except Exception as e:
        print(stock_symbol, " 크롤링 중 오류 발생:", e)
        return None, None

# 엑셀 파일 생성 함수
def create_excel_file(stock_symbols):
    data = []
    wb = Workbook()
    ws = wb.active

    for symbol in stock_symbols:
        name, price = crawl_stock_data(symbol)
        if name and price:
            ws.append([name, price])

    wb.save("sample.xlsx")

# UI 생성 함수
def create_ui():
    window = tk.Tk()
    window.title("주식 데이터 크롤러")

    # 윈도우 크기 조정 (너비 x 높이)
    window.geometry("400x200")  # 너비 400, 높이 200

    label = Label(window, text="주식 종목을 입력하세요 (콤마로 구분):")
    label.pack()

    entry = Entry(window)
    entry.pack()

    result_label = Label(window, text="")
    result_label.pack()

    def crawl_and_save():
        stock_symbols = entry.get().split(',')
        create_excel_file(stock_symbols)
        result_label.config(text="데이터 크롤링 및 저장 완료")
        # 저장된 엑셀 파일을 시스템의 기본 엑셀 프로그램으로 열기
        os.system('start sample.xlsx')

    crawl_button = Button(window, text="크롤링 및 저장", command=crawl_and_save)
    crawl_button.pack()

    window.mainloop()

if __name__ == "__main__":
    create_ui()
