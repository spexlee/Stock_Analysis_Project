import win32com.client

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 종목코드 리스트 구하기
# CpUtil.CpCodeMgr: 각종 코드 정보 및 코드 리스트를 얻을 수 있음
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
KOSPI_codeList = objCpCodeMgr.GetStockListByMarket(1) # 1 = KOSPI
KOSDAQ_codeList = objCpCodeMgr.GetStockListByMarket(2) # 2 = KOSDAQ

print("KOSPI 종목코드 :", len(KOSPI_codeList))
for i, code in enumerate(KOSPI_codeList):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)

print("코스닥 종목코드", len(KOSDAQ_codeList))
for i, code in enumerate(KOSDAQ_codeList):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)

print(" 거래소 + 코스닥 종목코드 ", len(KOSPI_codeList) + len(KOSDAQ_codeList))