# Python3 샘플 코드 #
import requests

myKey = 'MH5Jfb%2Brn%2F9xXXtUjvbxziwa1sAknbHIVKdX6T39yrrIiHq%2BN92FWRdsekTh425gtvOdxL6mEg1xtOyi4saUWg%3D%3D'
url = 'http://openapi.gimje.go.kr/rest/museum/getMuseumList'
params ={'serviceKey' : myKey, 'pageNo' : '1', 'numOfRows' : '10', 'museumNm' : '벽천미술관', 'museumType' : '미술관', 'roadAdd' : '벽골제로' }

response = requests.get(url, params=params)
print(response.content)
"""
# 라이브러리 import
import requests
import pprint
import json

# url 입력
url = 'http://api.data.go.kr/openapi/tn_pubr_public_cctv_api?serviceKey=개인인증키입력&pageNo=1&numOfRows=10&type=json'

# url 불러오기
response = requests.get(url)

#데이터 값 출력해보기
contents = response.text

# 데이터 결과값 예쁘게 출력해주는 코드
pp = pprint.PrettyPrinter(indent=4)
print(pp.pprint(contents))

#문자열을 json으로 변경
json_ob = json.loads(contents)
print(json_ob)
print(type(json_ob)) #json타입 확인

# 필요한 내용만 꺼내기
body = json_ob['response']['body']['items']
print(body)

# pandas import
import pandas as pd
from pandas.io.json import json_normalize
# Dataframe으로 만들기
dataframe = json_normalize(body)

print(dataframe)
"""
