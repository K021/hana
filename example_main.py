"""
이 문서는, hana 모둘을 효과적으로 사용하기 위한 예시 코드가 담긴 파일입니다.
참고용으로만 사용해야 하며, 이 모듈을 실행할 경우, 다양한 에러가 날 수 있습니다.
"""

# hana 모듈 내부의 클래스 임포트
# Hana 클래스: 매번 아이디와 비밀번호를 입력해야 하고, 직접 로그인 메서드를 호출해주어야 한다.
# HanaUser 클래스: config_secret.py 의 유저 정보를 사용하여 자동으로 로그인한다.
from hana import Hana, HanaUser


# Hana 클래스를 이용하는 방식
user_id = '아이디'
user_pw = '패스워드'
cert_pw = '공인인증서 비밀번호'

ha1 = Hana(user_id, user_pw, cert_pw)
ha1.login()
ha1.get_stock_tig_data()  # 주식의 틱데이터를 가져옴. 기본값은 Kodex 코스닥 150

# 수신하려는 틱 데이터의 종류가 2개라면, 인스턴스도 2개여야 한다.
ha2 = Hana(user_id, user_pw, cert_pw)
ha2.login()
ha2.get_futures_tig_data()  # 지수선물의 틱데이터를 가져옴. 기본값은 코스피 선물 19년 3월물


# HanaUser 클래스를 이용하는 방식
# 이 클래스를 이용하기 위해선, config_secret.py 안에 USER_INFO 가 정의되어 있어야 한다.
# USER_INFO = {
#     'user_id': b'alskfjslkfj',  # 이곳에 암호화된 user_id가 들어간다.
#     'user_pw': b'asdfsdafasdf',  # 이곳에 암호화된 user_pw가 들어간다.
#     'cert_pw': b'fasdfasfasdfasdfasdfsad',  # 이곳에 암호화된 cert_pw 가 들어간다.
# }
h = HanaUser()
h.get_stock_tig_data()  # 주식의 틱데이터를 가져옴. 기본값은 Kodex 코스닥 150

# 수신하려는 틱 데이터의 종류가 2개라면, 인스턴스도 2개여야 한다.
h2 = HanaUser()
h2.get_futures_tig_data()  # 지수선물의 틱데이터를 가져옴. 기본값은 코스피 선물 19년 3월물


# 저장된 틱 데이터 확인
print(h.tig_data)
# 저장된 틱 데이터를 엑셀로 가져오는 로직
h.tig_data_dump()

# 유저 정보 가져오기.
# 이 정보를 바탕으로 config_secret.py 의 USER_INFO 를 작성한다.
user_id_encrypted = h.user_id
user_pw_encrypted = h.user_pw
cert_pw_encrypted = h.cert_pw

