import tkinter as tk
from tkinter import Entry, Label, Button, StringVar
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook

# 크롤링 함수
def crawl_stock_data(stock_symbol):
    if stock_symbol.isnumeric():
        url = f"https://finance.naver.com/item/main.naver?code={stock_symbol}"
    else:
        url = f"https://finance.yahoo.com/quote/{stock_symbol}?p={stock_symbol}&.tsrc=fin-srch"

    price = 0
    name = stock_symbol
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        if stock_symbol.isnumeric():
            name = soup.select_one('.wrap_company a').text.strip()
            price = soup.select_one('.no_today .blind').text.strip()
        else:
            price_element = soup.select_one('[data-test="qsp-price"]')
            price = price_element.get('value')

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

    crawl_button = Button(window, text="크롤링 및 저장", command=crawl_and_save)
    crawl_button.pack()

    window.mainloop()

if __name__ == "__main__":
    create_ui()
