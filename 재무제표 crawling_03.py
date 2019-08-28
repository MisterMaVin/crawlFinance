# DART에서 재무제표 crawling

from selenium import webdriver
from datetime import datetime
import requests, bs4, openpyxl

driver = webdriver.Chrome("D:\\crawling\\util\\chromedriver.exe")

file_path = "D:\\crawling\\output\\코스닥재무제표_crawling_"
list_url = "http://dart.fss.or.kr/dsab001/search.ax?publicType=A001&publicType=A002&publicType=A003&textCrpNm="
report_url = "http://dart.fss.or.kr/dsaf001/main.do?rcpNo="

# 오늘 날짜 받아오기
today = datetime.now().strftime("%Y%m%d")

# excel 불러오기
wb = openpyxl.load_workbook(file_path + today + ".xlsx")
sheet = wb["Sheet"]

# 공통으로 쓰이는 부분 함수화
def getValue(co_code, co_name, exprList):
    value = 1
    for expr in exprList:
        try:
            value = getValueByXPath(expr)
            print(co_code, co_name, expr, "equal", value)
            break
        except Exception as e:
            print(co_code, co_name, expr, "equal", e)
            pass

    if value == 1 or value == 0:
        for expr in exprList:
            try:
                value = getValueByXPathMatch(expr)
                print(co_code, co_name, expr, "match", value)
                break
            except Exception as e:
                print(co_code, co_name, expr, "match", e)
                pass

    return value

def getValueByXPath(expr):
    value = driver.find_element_by_xpath("//p[normalize-space(text()) =  '" + expr + "']//parent::td//following-sibling::td").text
    #print(value, value.find("("))
    if value.find("(") > -1:
        value = value.replace("(","")
        value = value.replace(")","")
        value = "-" + value
    return int(value.replace(",",""))

def getValueByXPathMatch(expr):
    value = driver.find_element_by_xpath("//p[contains(text(), '" + expr + "')]//parent::td//following-sibling::td").text
    #print(value, value.find("("))
    if value.find("(") > -1:
        value = value.replace("(","")
        value = value.replace(")","")
        value = "-" + value
    return int(value.replace(",",""))
            

rowIndex = 1006
#rowIndex = 412
sRow_count = str(sheet.max_row)
#print(sRow_count)

for row in sheet.rows:
    rowIndex = rowIndex + 1
    #if rowIndex == 1009:
    #if rowIndex == 414:
    #    break
    co_code = sheet.cell(row=rowIndex,column=2).value
    co_name = sheet.cell(row=rowIndex,column=1).value

    if co_code is None:
        break

    #print(co_code)
    print(rowIndex)

    source_code = requests.get(list_url + co_code)
    plain_text = source_code.text
    bsObject = bs4.BeautifulSoup(plain_text, "html.parser")

    try:
        table = bsObject.find("table", {"summary":"공시서류검색에 대한 번호, 공시대상회사, 보고서명, 제출인, 접수일자, 비고 등을 알리는 표입니다."})
        
        for a in table.find_all("a"):
            link = a.get("href")
            if link.find("rcpNo=") > 0:
                report_no = link.split("rcpNo=")[1]

                # webdriver로 화면 접속
                driver.get(report_url + report_no)
                #print(report_url + report_no)

                try:
                    driver.find_element_by_xpath("//span[contains(text(), ' 재무제표')]").click()

                    driver.switch_to.frame(0)

                    total_assets = getValue(co_code, co_name, ["자산총계"])
                    total_capital = getValue(co_code, co_name, ["자본총계"])
                    sales = getValue(co_code, co_name, ["매출액","영업수익","매출","수익(매출액)","매출액 (주","수익(매출액) (주"])
                    gross_profit = getValue(co_code, co_name, ["매출총이익","영업이익","매출총이익 (주","영업이익(손실)","매출총손익","영업손익"])
                    net_income = getValue(co_code, co_name, ["당기순이익","반기순이익","당기순이익(손실)","반기순이익(손실)","분기순손익","당기순손익"])
                    #print(total_assets, total_capital, sales, gross_profit, net_income)

                    sheet.cell(row=rowIndex,column=7,value=total_assets)
                    sheet.cell(row=rowIndex,column=7).number_format = "0"
                    sheet.cell(row=rowIndex,column=8,value=total_capital)
                    sheet.cell(row=rowIndex,column=8).number_format = "0"
                    sheet.cell(row=rowIndex,column=9,value=sales)
                    sheet.cell(row=rowIndex,column=9).number_format = "0"
                    sheet.cell(row=rowIndex,column=10,value=gross_profit)
                    sheet.cell(row=rowIndex,column=10).number_format = "0"
                    sheet.cell(row=rowIndex,column=11,value=net_income)
                    sheet.cell(row=rowIndex,column=11).number_format = "0"

                    # 지표 계산
                    sRowIndex = str(rowIndex)
                    
                    sheet.cell(row=rowIndex,column=12,value="=F" + sRowIndex + "/H" + sRowIndex) # PBR
                    sheet.cell(row=rowIndex,column=13,value="=F" + sRowIndex + "/I" + sRowIndex) # PSR
                    sheet.cell(row=rowIndex,column=14,value="=F" + sRowIndex + "/K" + sRowIndex) # PER
                    sheet.cell(row=rowIndex,column=15,value="=J" + sRowIndex + "/G" + sRowIndex) # GP/A

                    # 순위 매기기
                    sheet.cell(row=rowIndex,column=16,value="=rank(L" + sRowIndex + ",$L$3:$L$" + sRow_count + ", 1)") # PBR
                    sheet.cell(row=rowIndex,column=17,value="=rank(M" + sRowIndex + ",$M$3:$M$" + sRow_count + ", 1)") # PSR
                    sheet.cell(row=rowIndex,column=18,value="=rank(N" + sRowIndex + ",$N$3:$N$" + sRow_count + ", 1)") # PER
                    sheet.cell(row=rowIndex,column=19,value="=rank(O" + sRowIndex + ",$O$3:$O$" + sRow_count + ")") # GP/A

                    # summary
                    sheet.cell(row=rowIndex,column=20,value="=sum(P" + sRowIndex + ":S" + sRowIndex + ")") # sum rank
                    sheet.cell(row=rowIndex,column=21,value="=rank(T" + sRowIndex + ",$T$3:$T$" + sRow_count + ", 1)") # rank sum rank

                    # url
                    sheet.cell(row=rowIndex,column=22,value=report_url + report_no)
                    
                    break
                except Exception as e:
                    print(e)
                    pass
            
    except Exception as e:
        print(e)
        pass
      
    #break

wb.save(file_path + today + ".xlsx")
