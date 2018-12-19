from hana import HanaUser


h = HanaUser()  # HanaUser 객체 생성 (자동 로그인)
h.get_stock_tig_data()  # Kodex 코스닥 150 주가 틱 데이터 수신

print(h.tig_data)  # 저장된 틱 데이터 확인
h.tig_data_dump()  # 저장된 틱 데이터를 엑셀로 가져오는 로직

