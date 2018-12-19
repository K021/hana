# hana

이 프로젝트는 하나대투증권의 1Q Api를 파이썬으로 이용하기 위한 패키지이다.

## 프로젝트 폴더 구조
```
Inv/
    - hana.py: 데이터 수신에 필요한 메서드와 속성을 지닌 Hana 클래스가 정의된 모듈
    - functions.py: 비밀번호 암호화, 복호화 함수가 내장된 모듈
    - config_secret.py: 암호화 키, 비밀번호 등 보안에 관련된 정보가 담긴 모듈
```

## Requirements
python==3.7.0  
openpyxl==2.5.6  
pyqt==5.9.2  

> 현재 임시로 프로젝트를 진행하고 있어서, 파이썬 환경이 전역 환경밖에 없다. 따라서 pip requirements 파일을 만들기 곤란해서 직접 적는다.