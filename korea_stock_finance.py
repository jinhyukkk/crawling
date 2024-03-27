import requests
import pandas as pd
import json
import os


def get_header(tr_id):
    headers = {"content-type": "application/json",
               "appkey": api_key,
               "appsecret": secret_key,
               "authorization": f"Bearer {access_token}",
               "tr_id": tr_id,
               }
    return headers


api_key = 'PSpOIZNMGeAuBOZwl9FlKxzAW1zegemOer21'
secret_key = '8MSYELfVZqrIJmrwV9CZkpDJgG7zN3+FBHg8yK7X1gqJd2gYPlxRKTkkmNQzL7fCO13Xf6GPSRG8bSUzgwmIBMOfK6MuvbBDj1651sTb+6PWwvbtMwYHOS4eJ0sn6oBSROhqiYFbhCbkGO6wXdGml01//JspPpEQRdDWQOJX2gzZllnJ9vM='

base_url = 'https://openapi.koreainvestment.com:9443'

headers = {"content-type": "application/json"}
body = {
    "grant_type": "client_credentials",
    "appkey": api_key,
    "appsecret": secret_key,
}
url = base_url + '/oauth2/tokenP'
res = requests.post(url, headers=headers, data=json.dumps(body)).json()
print(res)
access_token = res['access_token']

# 파라미터 설정
ticker = '133690'  # 삼성전자 티커
api_starttime = '20170101'  # 0000년 1월 1일부터
api_endtime = '20211231'  # 9999년 12월 31일까지의 데이터
api_freq = 'M'  # 월봉
api_adj_price = True  # 수정주가로 받아오기

# headers, params 제작
headers = get_header('FHKST03010100')
params = {
    "fid_cond_mrkt_div_code": "J",
    "fid_input_iscd": ticker,
    "fid_input_date_1": api_starttime,
    "fid_input_date_2": api_endtime,
    "fid_period_div_code": api_freq,
    "fid_org_adj_prc": 0 if api_adj_price else 1,
}

# 데이터 수신 및 정제
url = base_url + '/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice'
temp = requests.get(url, headers=headers, params=params).json()['output2'][::-1]
# temp = pd.DataFrame(temp)
# temp = temp.rename(
#     columns={'stck_bsop_date': 'timestamp', 'stck_clpr': 'close', 'stck_oprc': 'open', 'stck_hgpr': 'high',
#              'stck_lwpr': 'low'})
# temp[['close', 'open', 'high', 'low']] = temp[['close', 'open', 'high', 'low']].astype(float)

totalPrice = 0  # 매수금액
totalQuantity = 0  # 매수수량
averagePrice = 0  # 평균단가
evaluationAmount = 0  # 평가금액
monthlyPrice = 700000  # 월납입금
result = []

for value in temp:
    lastPrice = int(value["stck_clpr"])
    quantity = round(round(monthlyPrice / lastPrice), 4)  # 매수할 수량
    price = lastPrice * quantity  # 매수할 금액

    totalQuantity += quantity  # 총수량
    totalPrice += price  # 총금액
    averagePrice = round(totalPrice / totalQuantity, 4)  # 평균단가
    rateReturn = round((lastPrice - averagePrice) / averagePrice * 100, 4)  # 수익률
    evaluationAmount = lastPrice * totalQuantity  # 평가금액
    dataSet = {'날짜': value["stck_bsop_date"], '현재금액': "{:,}".format(lastPrice), '평균단가': "{:,}".format(averagePrice),
               '매수수량': "{:,}".format(totalQuantity), '매수금액': "{:,}".format(totalPrice), '추가금액': "{:,}".format(price),
               '평가금액': "{:,}".format(evaluationAmount), '수익률': str(rateReturn) + '%'}
    result.append(dataSet)

result = pd.DataFrame(result)
# os.system('taskkill /im excel.exe')
result.to_excel("stock_view.xlsx", sheet_name="Sheet1", index=False)

# 저장된 엑셀 파일을 시스템의 기본 엑셀 프로그램으로 열기

os.system('start stock_view.xlsx')
