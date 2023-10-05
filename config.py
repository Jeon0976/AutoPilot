### kakao coordinate 상수 값 
# EXCEL_FILENAME = 엑셀 파일 이름 (test.xlsx)
# WORKSHEET = 첫 번째 Sheet일 경우: 0(기본값)
# START_ROW = 시작 행: 1(기본값)
# ADDRESS_COLUMN = 주소 값이 들어간 열: A(기본값)
# LATITUDE_COLUMN = Lat 좌표 값이 들어갈 열: B(기본값)
# LONGITUDE_COLUMN = LNG 좌표 값이 들어갈 열: C(기본값)

class KakaoMapAddressToCoordinate: 
    EXCEL_FILENAME = "seoul1_dodam-2.xlsx"
    WORKSHEET = 1
    START_ROW = 9
    ADDRESS_COLUMN = "C"
    LATITUDE_COLUMN = "D"
    LONGITUDE_COLUMN = "E"