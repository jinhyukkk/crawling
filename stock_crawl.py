import tkinter as tk
from tkinter import Entry, Label, Button, StringVar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook

# 웹 드라이버 초기화
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # GUI 없이 실행 (백그라운드 실행)
driver = webdriver.Chrome()  # chromedriver.exe의 경로로 설정

# 크롤링 함수
def crawl_stock_data(stock_symbol):

    if stock_symbol.isnumeric():
        url = f"https://finance.naver.com/item/main.naver?code={stock_symbol}"
    else:
        url = f"https://finance.yahoo.com/quote/{stock_symbol}?p={stock_symbol}&.tsrc=fin-srch"

    driver.get(url)

    # WebDriverWait 객체 생성
    wait = WebDriverWait(driver, 5)  # 5초 동안 대기
    price = 0
    name = stock_symbol
    try:
        if stock_symbol.isnumeric():
            # id가 "myElement"인 요소가 나타날 때까지 대기
            element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "rate_info")))
            # 클래스명이 "wrap_company"인 요소를 찾기
            name = driver.find_element(By.CLASS_NAME, "wrap_company").find_element(By.TAG_NAME, "a").text

            # 클래스명이 "no_down"인 <em> 태그 선택
            tags = driver.find_element(By.CLASS_NAME, "no_today").find_elements(By.TAG_NAME, "span")

            # <span> 태그 내의 텍스트 추출 및 조합
            price = ''.join(span.text for span in tags)

        else:
            # id가 "myElement"인 요소가 나타날 때까지 대기
            element = wait.until(EC.presence_of_element_located((By.ID, "YDC-Lead-Stack-Composite")))
            # 페이지에서 주식 가격 정보 가져오기
            price_element = driver.find_element(By.XPATH, "//fin-streamer[@data-test='qsp-price']")
            price = price_element.get_attribute("value")

        # 요소가 로드된 후에 할 작업 수행
        print("요소가 로드되었습니다.")
        print("요소의 텍스트:", element.text)
    except Exception as e:
        print("대기 동안 오류 발생:", e)

    return name, price

# 엑셀 파일 생성 함수
def create_excel_file(stock_symbols):
    data = []
    wb = Workbook()
    ws = wb.active

    for symbol in stock_symbols:
        name, price = crawl_stock_data(symbol)
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
    driver.quit()  # 모든 크롤링이 완료된 후 브라우저 종료
