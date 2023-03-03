import requests
import json
from bs4 import BeautifulSoup
import xlsxwriter
outWorkbook = xlsxwriter.Workbook("data.xlsx")
outSheet = outWorkbook.add_worksheet()
outSheet.write("A1","lokasi")
outSheet.write("B1","lat")
outSheet.write("C1","long")
outSheet.write("D1","jenis")
outSheet.write("E1","detail")
# outSheet.write("F1","timestamp")
def GetDetail(id,fas):
    url = 'https://sipsn.menlhk.go.id/sipsn/public/home/showdatatable'
    payload = {
        'sysgrup': fas,
        'id': id
    }
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,id-ID;q=0.8,id;q=0.7',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': 'ci_session=a7a8127d847c97efdb0a4ef33abfb0dc0abb7253',
        'dnt': '1',
        'Host': 'sipsn.menlhk.go.id',
        'Origin': 'https://sipsn.menlhk.go.id',
        'Referer': 'https://sipsn.menlhk.go.id/sipsn/public/home/peta',
        'sec-ch-ua': '"Chromium";v="110", "Not A(Brand";v="24", "Google Chrome";v="110"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    }

    response = requests.post(url, data=payload, headers=headers)
    data=json.loads(response.text)
    html_content = data["table"]
    soup = BeautifulSoup(html_content, 'html.parser')
    clean_text = soup.get_text()
    return clean_text
url = 'https://sipsn.menlhk.go.id/sipsn/public/home/getMarker'
payload = {
    'dd_propinsi': 'ALL',
    'dd_district': '',
    'dd_fasilitas': '',
    'exclude_fasilitas': ['rth', 'pengepul']
}
headers = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'en-US,en;q=0.9,id-ID;q=0.8,id;q=0.7',
    'Connection': 'keep-alive',
    'Content-Length': '80',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Cookie': 'ci_session=cbbd52473669b263450491930ca0d9f7cdbca8f3',
    'dnt': '1',
    'Host': 'sipsn.menlhk.go.id',
    'Origin': 'https://sipsn.menlhk.go.id',
    'Referer': 'https://sipsn.menlhk.go.id/sipsn/public/home/peta',
    'sec-ch-ua': '"Chromium";v="110", "Not A(Brand";v="24", "Google Chrome";v="110"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest'
}

response = requests.post(url, data=payload, headers=headers)
data = json.loads(response.text)
# print(data["markers"])
for y,x in enumerate(data["markers"]):
    # print(y)
    outSheet.write(y+1,0,x[0])
    outSheet.write(y+1,1,x[1])
    outSheet.write(y+1,2,x[2])
    outSheet.write(y+1,3,x[4])
    detail = GetDetail(x[5],x[4])
    outSheet.write(y+1,4,detail)
    print(detail)
outWorkbook.close()
