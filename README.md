# 주식 투자 정보 
## 성공적인 투자를 위한 재무정보 및 주가 정보를 parsing이나 crawling을 통해 가져오자.
### 1. 필요한 정보
- 회사명
- 종목코드
- 거래정지 여부 : 거래정지된 주식은 취급할 필요가 없다.
- 업종명
- 업종코드
- 현재가
- 상장주식수
- 시가총액
- 자산총계
- 자본총계
- 매출액
- 매출총이익
- 당기순이익
- PBR
- PSR
- PER
- GP/A
### 2. 필요한 정보를 얻을 곳
- 네이버
    - 현재가
    - 상장주식수
- 다트
    - 재무상태표
        - 회사명
        - 종목코드
        - 업종명
        - 업종코드
        - 자산총계 : ifrs_Assets
        - 자본총계 : ifgs_Equity
    - 손익계산서
        - 매출액 : ifrs_Revenue
        - 매출총이익 : ifrs_GrossProfit
        - 당기순이익 : ifrs_ProfitLoss
- formula
    - PBR : 자본총계 / 시가총액
    - PSR : 매출액 / 시가총액
    - PER : 당기순이익 / 시가총액
    - GP/A : 매출총이익 / 자산총계
### 3. 구현방법
    1. 네이버 금융에서 시가총액 기준으로 종목코드, 현재가를 가져온다.
    2. 네이버 금융에서 거래정지 목록에서 종목코드, 회사명을 가져온다.
    3. 다트에서 공시정보활용마당에서 필요한 보고서를 다운받는다.
    4. 재무상태표를 파싱해서 위의 데이터를 가져온다.
    5. 손익계산서를 파싱해서 위의 데이터를 가져온다.
    6. 위의 데이터들을 필요에 맞게 각 sheet에 넣는다.
        1. sheet1 : 증시현황 : sheet2 기준으로 **종목코드**, 현재가를 가져와서 나머지 sheet에서 데이터를 vlookup으로 가져온다.
        2. sheet2 : 시가총액 : **종목코드**, 현재가
        3. sheet3 : 재무상태표 : **종목코드**, 회사명, 업종코드, 업종명, 자산총계, 자본총계
        4. sheet4 : 손익계산서 : **종목코드**, 매출액, 매출총이익, 당기순이익
        5. sheet5 : 회사 개별 : **종목코드**, 상장주식수
        6. sheet6 : 거래 정지 : **종목코드**, 회사명
    7. **종목코드**를 key로 해서 vlookup을 하면 엑셀을 쉽게 컨트롤할 수 있을 것이다.
### 4. Rule
    - 엑셀 파일명 : KOSDAQ_증시현황_20190905.xlsx
### 5. 요청사항
    - 아직 만들지도 않았지만, 사용은 최대한 간단히 사람의 손을 적게 타고 싶다.
    - 따라서 다트에서 필요 정보를 다운받아서 압축을 풀고, 재무상태표 파일과 손익계산서 파일의 path만 입력하면 아웃풋이 나오도록 개발 필요.
