# naver 코스닥 시가총액
from datetime import datetime
import requests, bs4, openpyxl

# 기본 정보 setting
url = "https://finance.naver.com/sise/sise_market_sum.nhn?sosok=1&page="
file_path = "D:\\crawling\\output\\코스닥재무제표_crawling_"

# 오늘 날짜 받아오기
today = datetime.now().strftime("%Y%m%d")

# excel 생성
wb = openpyxl.Workbook()
sheet = wb["Sheet"]
sheet.merge_cells(start_row=1,start_column=12,end_row=1,end_column=15)
sheet.merge_cells(start_row=1,start_column=16,end_row=1,end_column=19)
sheet.cell(row=1,column=12,value="지수")
sheet.cell(row=1,column=16,value="rank")
sheet.cell(row=2,column=1,value="종목명")
sheet.cell(row=2,column=2,value="종목코드")
sheet.cell(row=2,column=3,value="시가")
sheet.cell(row=2,column=4,value="상장주식수")
sheet.cell(row=2,column=5,value="업종코드")
sheet.cell(row=2,column=6,value="시가총액")
sheet.cell(row=2,column=7,value="자산총계")
sheet.cell(row=2,column=8,value="자본총계")
sheet.cell(row=2,column=9,value="매출액")
sheet.cell(row=2,column=10,value="매출총이익")
sheet.cell(row=2,column=11,value="당기순이익")
sheet.cell(row=2,column=12,value="PBR")
sheet.cell(row=2,column=13,value="PSR")
sheet.cell(row=2,column=14,value="PER")
sheet.cell(row=2,column=15,value="GP/A")
sheet.cell(row=2,column=16,value="PBR")
sheet.cell(row=2,column=17,value="PSR")
sheet.cell(row=2,column=18,value="PER")
sheet.cell(row=2,column=19,value="GP/A")
sheet.cell(row=2,column=20,value="순위 합계")
sheet.cell(row=2,column=21,value="전체 순위")

page = 0
rowIndex = 2

while True:
    page = page + 1
    source_code = requests.get(url + str(page))
    plain_text = source_code.text
    bsObject = bs4.BeautifulSoup(plain_text, "html.parser")

    table = bsObject.find("table", {"class":"type_2"})

    #print(table.get("summary"))

    if len(table.find_all("a")) == 0:
        break

    for a in table.find_all("a"):
        link = a.get("href")
        if link.find("main") > 0:
            co_name = a.text
            co_code = link.split("=")[1]
            cur_price = int(a.parent.find_next().find_next().text.replace(",",""))
            #print(co_name, co_code, cur_price)

            # excel에 data 넣기
            rowIndex = rowIndex + 1
            sheet.cell(row=rowIndex,column=1,value=co_name)
            sheet.cell(row=rowIndex,column=2,value=co_code)
            sheet.cell(row=rowIndex,column=3,value=cur_price)
            sheet.cell(row=rowIndex,column=3).number_format = "0"

wb.save(filename=file_path + today + ".xlsx");
