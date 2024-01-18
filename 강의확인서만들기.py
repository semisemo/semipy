import pprint # dictionary 를 보기좋게 출력하려고 쓰는 라이브러리
import openpyxl
from datetime import datetime

raw_data = '교육파일.xlsm'
save_data = '강의확인서.xlsx'

MONTH = 3
EDU = "꾸러기수사대"
# 0. 데이터 통합시트 먼저 읽기
wb = openpyxl.load_workbook(raw_data)
lec_raw_sheet = wb['통합시트']

first_row_number = 2
last_row_number = lec_raw_sheet.max_row


"""
MONTH 데이터만 추출될 거고

{
    'lec_name'{ 
        강사A : [날짜, 시간, 신청기관, 반, 인원],
                [날짜, 시간, 신청기관, 반, 인원]}
        강사b : 
}


"""

PROGRAM_DATA = {}

for row in lec_raw_sheet.iter_rows(min_row=first_row_number, max_row=last_row_number):
    date = row[1].value

    if date != None:
        if date.month != MONTH:
            continue

        lec_name = row[0].value # 강의명
        date = row[1].value.strftime("%Y-%m-%d")  #날짜
        timetable = row[3].value # 시간
        applicant = row[4].value # 신청기관
        division = row[5].value # 반
        amount = row[6].value # 인원
        teacher1 = row[8].value # 강사1
        teacher2 = row[9].value # 강사2

        data = [date, timetable, applicant, division] #월별 데이터 나오는것까지는 성공



        if not lec_name in PROGRAM_DATA:
            PROGRAM_DATA[lec_name] = {}  

        
        for teacher in (teacher1, teacher2):
            if teacher is None:  
                continue
            if teacher is str('-'):
                continue
          
            if not teacher in PROGRAM_DATA[lec_name]:
                PROGRAM_DATA[lec_name][teacher] = []
                
            PROGRAM_DATA[lec_name][teacher].append(data) #data를 붙이면 1개만 나오구 ㅜㅜ

wb.close()  #이거 들어갈 자리도 물어봐야지...

print("-------------------- PROGRAM_DATA: 출 력 ---------------------------\n")
pprint.pprint(PROGRAM_DATA, width=1, indent=5)

#위에서 추출된 PROGRAM_DATA를 가지고 강사별로 강의확인서 인적사항 적기


ws_personal = wb["인적사항"]
PERSONAL_DICT = {}

for row in ws_personal.iter_rows(min_row=2):

    name = row[0].value # 첫 칸 값

    if not name:
        break

    social_id = row[1].value # 생년월일
    address = row[2].value # 주소
    phone = row[3].value # 연락처
    bank = row[4].value # 은행
    banknum = row[5].value # 계좌번호

    # print(name, social_id, address,phone, bank, banknum)

    PERSONAL_DICT[name] = [social_id, address,phone, bank, banknum]

    #직원같은경우는 목록에 안쓰는데 어떻게 할건지....
    #데이터는 뽑앗는데,,,, 어떻게 매칭시켜서 써야할지를 모르겠음




            


    

