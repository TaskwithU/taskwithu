# 2025-01-28 오전 9:37 NAVER 금융에 조회 요청함
from io import StringIO
import pandas as pd
import winsound
import time
import requests
from bs4 import BeautifulSoup
import sys
import re

# 출력을 파일과 콘솔에 동시에 출력하기 위해 추가
class DualOutput:
    def __init__(self, file_path):
        self.terminal = sys.stdout
        self.log = open(file_path, 'w', encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(re.sub(r'\033\[\d+m', '', message))  # 칼라코드 제거

    def flush(self):
        self.terminal.flush()
        self.log.flush()

# 파일로 출력 시작
file_path = 'C:\\Result_Checker\\OLD\\Result_Checker_NAVER.txt'
sys.stdout = DualOutput(file_path)

# 시작 시간 기록
start_time = time.time()

# 엑셀 파일에서 데이터 읽기
#file_path = 'C:\\Result_Checker\\OLD\\KOSDAQ_LASTEST_ALL_1668EA.xlsx'
file_path = 'C:\\Result_Checker\\OLD\\0123 KOSDAQ CODING DIRECT RESULT.xlsx'
df = pd.read_excel(file_path)

# 조건 1: 제외할 조건 목록
exclude_conditions = ["정리매매", "관리", "투자위험", "투자경고", "투자주의", "거래정지", "환기", "불성실공시"]

# 필터링된 종목 코드 추출
df_filtered = df[~df['테마'].isin(exclude_conditions)]
tickers = df_filtered['코드'].apply(lambda x: str(x).zfill(6)).tolist()
names = df_filtered['종목명'].tolist()

# 네이버 금융에서 영업일 확인하는 함수 정의
def get_naver_business_days(years, month):
    business_days = []
    for year in years:
        url = f'https://finance.naver.com/sise/sise_index_day.nhn?code=KOSPI&year={year}&month={month}'
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        for day in soup.find_all('td', {'class': 'date'}):
            date = day.get_text(strip=True)
            date = pd.to_datetime(date, format='%Y.%m.%d')
            business_days.append(date)

    return business_days

# base_date가 영업일인지 확인하는 함수 정의
def is_business_day(date, business_days):
    return pd.to_datetime(date) in business_days


# 네이버 금융에서 주식 데이터를 여러 페이지에서 수집하는 함수 정의
def fetch_stock_data(ticker, start_date, end_date, max_pages=40):
    df_list = []
    start_date = pd.to_datetime(start_date)  # 문자열을 Timestamp로 변환
    end_date = pd.to_datetime(end_date)  # 문자열을 Timestamp로 변환
    for page in range(1, max_pages + 1):
        url = f'https://finance.naver.com/item/sise_day.naver?code={ticker}&page={page}'
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        table_html = str(soup.find("table"))
        df = pd.read_html(StringIO(table_html))[0].dropna()

        # '날짜' 열이 존재하는지 확인
        if '날짜' not in df.columns:
            print(f"오류: {ticker}의 데이터에서 '날짜' 열을 찾을 수 없습니다.")
            continue

        df_list.append(df)

        # 종료 날짜보다 이전의 데이터가 있는지 확인
        if pd.to_datetime(df['날짜']).min() <= start_date:
            break

    if df_list:
        df = pd.concat(df_list)
        # 데이터 전처리
        df['날짜'] = pd.to_datetime(df['날짜'])
        df = df[df['날짜'] <= end_date]

        # 날짜 기준으로 정렬
        df = df.sort_values(by='날짜')

        df.set_index('날짜', inplace=True)

        # 이동평균 계산
        df['MA5'] = df['종가'].rolling(window=5).mean()
        df['MA10'] = df['종가'].rolling(window=10).mean()
        df['MA20'] = df['종가'].rolling(window=20).mean()
        df['MA60'] = df['종가'].rolling(window=60).mean()
    else:
        print(f"오류: {ticker}의 데이터가 제대로 수집되지 않았습니다.")
        df = pd.DataFrame()

    return df

# base_date와 이전 영업일 찾는 함수 정의
def get_previous_business_days(base_date, business_days):
    previous_business_days = []
    date = pd.to_datetime(base_date)

    for _ in range(5):  # 최대 5일까지 이전 영업일 확인
        date -= pd.DateOffset(days=1)
        if date in business_days:
            previous_business_days.append(date)
        if len(previous_business_days) == 2:  # base_date의 전날과 그 전날을 찾으면 종료
            break

    if len(previous_business_days) < 2:
        print("오류: 유효한 previous_close_date를 찾지 못했습니다. 날짜를 확인해주세요.")
        return None

    return previous_business_days

# 변동율 계산 함수 정의 | 계산식 ===> 변동율 (%) = ((기간내 최고 종가 - 기간내 최저 종가) / 기간내 최저 종가) * 100
def calculate_volatility(data, data_long_ma, names, base_date):
    volatility = []
    for idx, (ticker, stock_data) in enumerate(data.items()):
        stock_data_long_ma = data_long_ma[ticker]
        if not stock_data.empty and base_date in stock_data.index and base_date in stock_data_long_ma.index:
            low_close = stock_data['종가'].min()
            high_close = stock_data['종가'].max()
            if not pd.isna(low_close) and not pd.isna(high_close):
                low_close = low_close if not isinstance(low_close, pd.Series) else low_close.iloc[0]
                high_close = high_close if not isinstance(high_close, pd.Series) else high_close.iloc[0]
                if low_close != 0 and low_close != high_close:
                    variability = round((high_close - low_close) / low_close * 100, 2)
                    current_price = int(stock_data['종가'].loc[base_date])
                    short_5ma = stock_data['MA5'].loc[base_date] if 'MA5' in stock_data else 'N/A'
                    mid_20ma = stock_data['MA20'].loc[base_date] if 'MA20' in stock_data else 'N/A'
                    long_60ma = stock_data_long_ma['MA60'].loc[base_date] if 'MA60' in stock_data_long_ma else 'N/A'
                    volatility.append((ticker.replace('.KQ', ''), variability, names[idx], current_price, short_5ma, mid_20ma, long_60ma))
    return volatility



# 시작 날짜와 종료 날짜 정의
start_date = '2025-01-17' #변동율 검색 시작 시점
end_date = '2025-01-23' #변동율 검색 종료 시점
base_date = '2025-01-23' #검색시점

#long_60ma_start_date = '2024-09-17' # 이동평균 계산 시작 시점 (최소 4개월 전)
long_60ma_start_date = '2024-07-01' # 이동평균 계산 시작 시점 (최소 4개월 전)
long_60ma_end_date = '2025-01-23'

# 네이버 금융에서 2000~2050년 영업일 확인
years = range(2000, 2051)
naver_business_days = get_naver_business_days(years, 1)

# base_date가 영업일인지 확인
if not is_business_day(base_date, naver_business_days):
    print(f"지정하신 날짜 {base_date}는 (네이버 조회 기준) 거래소 영업일이 아닙니다. 날짜를 확인해 주세요.")
    exit()

# 주식 데이터 가져오기 (이동평균 계산용)
data_long_ma = {}
for ticker in tickers:
    data_long_ma[ticker] = fetch_stock_data(ticker, long_60ma_start_date, long_60ma_end_date)

# 데이터 확인
for ticker in tickers:
    print(f"{ticker} 데이터 (처음 5개 행):")
    print(data_long_ma[ticker].head())
    print(f"{ticker} 데이터 (마지막 5개 행):")
    print(data_long_ma[ticker].tail())


# 주식 데이터 가져오기
data = {}
for ticker in tickers:
    data[ticker] = fetch_stock_data(ticker, start_date, end_date)


# 조건 4: 일봉상 5봉 동안 최저종가 대비 최고종가 변동폭 상위 150종목 계산
# 주식 데이터로 변동율 계산
volatility = calculate_volatility(data, data_long_ma, names, base_date)

# 변동폭 상위 150종목 필터링
sorted_volatility = sorted(volatility, key=lambda item: item[1], reverse=True)
top_150 = sorted_volatility[:150]

# 종목별 주식 데이터 가져오기 및 조건 확인
valid_stocks = []
seen_stock_names = set()
first_valid_stock = None

for idx, (ticker, variability, stock_name, current_price, short_5ma, mid_20ma, long_60ma) in enumerate(top_150):
    stock_data = data[ticker]
    stock_data_long_ma = data_long_ma[ticker]  # 여기서 stock_data_long_ma 변수를 정의합니다.

    if base_date not in stock_data.index:
        print(f"오류: base_date {base_date} 날짜에 주식 데이터가 존재하지 않습니다. 날짜를 확인해주세요.")
        continue

    previous_business_days = get_previous_business_days(base_date, naver_business_days)

    if previous_business_days is None:
        continue

    previous_close_date = previous_business_days[0].strftime('%Y-%m-%d')
    previous_close = stock_data['종가'].loc[previous_close_date]

    # 현재 등락율 계산
    change_rate = ((current_price - previous_close) / previous_close) * 100

    # 60MA 소수점 첫째 자리까지 반올림
    long_60ma_rounded = round(long_60ma, 1) if long_60ma != 'N/A' else 'N/A'

    print(
        f"{stock_name} ({ticker}): 현재가 {current_price}원 등락율 {change_rate:.2f}% 최근 5일간 변동율 {variability:.2f}% 5MA {short_5ma}원 20MA {mid_20ma}원 60MA {long_60ma_rounded}원"
    )

    if stock_name not in seen_stock_names:
        seen_stock_names.add(stock_name)
        valid_stocks.append(
            (stock_name, ticker, current_price, change_rate, previous_close_date, previous_close, variability))
        if not first_valid_stock:
            first_valid_stock = (
            stock_name, ticker, current_price, change_rate, previous_close_date, previous_close, variability)

    # 단순 이동평균 계산
    short_5ma = stock_data['MA5'].loc[base_date] if 'MA5' in stock_data else 'N/A'
    mid_20ma = stock_data['MA20'].loc[base_date] if 'MA20' in stock_data else 'N/A'
    long_60ma = stock_data_long_ma['MA60'].loc[base_date] if 'MA60' in stock_data_long_ma else 'N/A'


    # 조건 2: 현재가 1,200원 ~ 100,000원
    if 1200 <= current_price <= 100000:

        # 조건 3: 1~10봉 연속 양봉
        #positive_candles = (stock_data['종가'] > stock_data['시가']).sum()
        #if 1 <= positive_candles <= 10:

            # 조건 5: 현재가 기준 등락율 5.10% ~ 30.00%
            #if 5.10 <= change_rate <= 30.00:

                # 조건 6: 기타종목 [ETF, ETN종목] 제외
                #if "ETF" not in stock_name and "ETN" not in stock_name:

                    # 조건 7: 단순이평 단기[5] 중기[20] 장기[60] 이평이 역배열 제외
                    #if not (short_5ma < mid_20ma < long_60ma):

                        # 조건 8: 단순 60 이평 < 종가 1봉 지속
                        #if not pd.isna(long_60ma) and current_price > long_60ma:

                            # 조건 9: 단순 60 이평 90.00% ~ 192.00%
                            #if 0.90 * long_60ma <= current_price <= 1.92 * long_60ma:

                                # 조건 10: 0봉전 거래량 20,000주 ~ 90,000,000,000주
                                #if 20000 <= stock_data['거래량'].loc[base_date] <= 90000000000:

##################################### 아래 조건은 if문 마지막 문법이므로 항상 켜둘것 #####################################
                                    if stock_name not in seen_stock_names:
                                        seen_stock_names.add(stock_name)
                                        valid_stocks.append((stock_name, ticker, current_price, change_rate, previous_close_date, previous_close, variability))
                                        if not first_valid_stock:
                                            first_valid_stock = (stock_name, ticker, current_price, change_rate, previous_close_date, previous_close, variability)
##################################### 위의 조건은 if문 마지막 문법이므로 항상 켜둘것 #####################################

# 경과 시간 계산
end_time = time.time()
elapsed_time = end_time - start_time
formatted_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))

# 결과 출력
print(f" [ 결과 조회 합계: {len(valid_stocks)}개 ]")
print(f" [ 경과 시간: {formatted_time} ]")

# 추가된 출력 구문
if first_valid_stock:
    stock_name, ticker, current_price, change_rate, previous_close_date, previous_close, variability = first_valid_stock
    print(f"조건 2와 조건 5에 부합하는 첫 번째\033[35m 종목: {stock_name} \033[0m")
    print(f"조건 2: 검색시점 전날: {previous_close_date},\033[35m 검색시점: {base_date} \033[0m 검색시점 종가: {current_price}")
    print(f"조건 4: 일봉 5봉 최저종가 대비 최고종가 변동폭 상위 150종목\033[35m 날짜: {start_date} ~ {end_date} \033[0m")
    print(f"조건 5: 검색시점 전날: {previous_close_date},\033[35m 검색시점: {base_date} \033[0m 검색시점 종가: {current_price}, 변동율: {change_rate:.2f}%")
    #print(df)
# 조회한 URL 프린트
print("조회한 URL: https://finance.naver.com/item/sise_day.nhn?code=KOSPI&page=1")

# 사운드 파일 재생
winsound.PlaySound('C:\\Result_Checker\\OLD\\sound0.wav', winsound.SND_FILENAME)