# naver 코스닥 시가총액
from datetime import datetime
import requests, bs4, openpyxl

# 기본 정보 setting
url = "https://finance.naver.com/item/main.nhn?code="
file_path = "D:\\crawling\\output\\코스닥재무제표_crawling_"

# 오늘 날짜 받아오기
today = datetime.now().strftime("%Y%m%d")

# excel 불러오기
wb = openpyxl.load_workbook(file_path + today + ".xlsx")
sheet = wb["Sheet"]

rowIndex = 2

for row in sheet.rows:
    rowIndex = rowIndex + 1
    co_code = sheet.cell(row=rowIndex,column=2).value
    cur_price = sheet.cell(row=rowIndex,column=3).value

    if co_code is None:
        break

    print(co_code)

    source_code = requests.get(url + co_code)
    plain_text = source_code.text
    bsObject = bs4.BeautifulSoup(plain_text, "html.parser")

    table = bsObject.find("table",{"summary":"시가총액 정보"})
    for th in table.find_all("th"):
        if th.text == "상장주식수":
            stock_count = int(th.find_next().text.replace(",",""))
            
            sheet.cell(row=rowIndex,column=4,value=stock_count)
            sheet.cell(row=rowIndex,column=6,value=(stock_count * cur_price))
            sheet.cell(row=rowIndex,column=6).number_format = "0"
            break
        
    table = bsObject.find("table",{"summary":"동일업종 PER 정보"})
    for a in table.find_all("a"):
        link = a.get("href")
        if link.find("upjong&no=") > 0:
            cate_code = link.split("upjong&no=")[1]
            
            sheet.cell(row=rowIndex,column=5,value=cate_code)
            break

wb.save(filename=file_path + today + ".xlsx");
