import requests
import pandas as pd
import json

def get_header(tr_id):
    headers = {"content-type":"application/json",
            "appkey":api_key,
            "appsecret":secret_key,
            "authorization":f"Bearer {access_token}",
            "tr_id":tr_id,
            }
    return headers

base_url = 'https://openapi.koreainvestment.com:9443'

# 파라미터 설정
ticker = '005930'  # 삼성전자 티커
api_starttime = '00000101'  # 0000년 1월 1일부터
api_endtime = '99991231'  # 9999년 12월 31일까지의 데이터
api_freq = 'D'  # 일봉
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
temp = pd.DataFrame(temp)
temp = temp.rename(
    columns={'stck_bsop_date': 'timestamp', 'stck_clpr': 'close', 'stck_oprc': 'open', 'stck_hgpr': 'high',
             'stck_lwpr': 'low', 'acml_vol': 'volume'})
temp[['close', 'open', 'high', 'low', 'volume']] = temp[['close', 'open', 'high', 'low', 'volume']].astype(float)