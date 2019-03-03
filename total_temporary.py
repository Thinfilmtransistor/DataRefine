import pandas as pd
import win32com.client
import os
import re


def get_csv_name(filepath):
    filenames = os.listdir(filepath)

    csvlist = []
    for filename in filenames:
        name = filename.split(".")
        if len(name) == 1:
            pass
        elif name[1] == "csv":
            csvlist.append(filename)
    return csvlist

def calculating(rawdatapath, calpath):
    # ====================================================================
    # rawdata 경로명에서 file name 추출
    filename = os.path.basename(rawdatapath)
    splitname, ext = filename.split(".")
    p = re.compile("\d+")

    # 정리 csv 이름
    refine_name = splitname[:-3]

    # 정리 csv이름 중 숫자 추출
    m = p.search(refine_name)

    # 추출된 숫자는 sheet 이름에 쓰임
    sheet_name = m.group()[-2:]

    # ====================================================================


    with open(rawdatapath, 'r') as f:
        data = f.readlines()
        with open("{}.{}".format(refine_name, ext), 'w') as f2:
            for a, i in enumerate(data[20:]):
                f2.write(i)

    # =================================================================================
    # read csv like text에서 정리한 data를 읽는다.
    df = pd.read_csv('{}.csv'.format(refine_name))

    # df중 CH1열이 0보다 큰 행의 index를 rownumber에 list로 저장한다
    rownumber = df[df['CH1'] >= 0].index

    # rowindecator는 rownumber가 연속해서 존재하는경우 같은 list에 묶어서 보관한다
    rowindecator = []

    # rowindecator2는 rownumber가 보관하는 list의 크기를 보여준다.
    rowindecator2 = []

    # temlist는 rowindecator에 list로 묶어서 넣기위한 임시 list다
    temlist = []

    # ticker는 rownumber와 같고 ticker가 1증가할때 rownumber도 1 증가하면 같은 list로 묶어준다
    ticker = rownumber[0]
    for i in rownumber:
        if i == ticker:
            temlist.append(i)
            ticker += 1
        else:
            rowindecator.append(temlist)
            rowindecator2.append(len(temlist))
            temlist = []
            ticker = i
            temlist.append(i)
            ticker += 1

    # for 문이 if 참으로 끝나면 temlist가 rowindecator에 들어가지 못한채 끝나므로 다음 절차가 필요하다
    if temlist:
        rowindecator.append(temlist)
        rowindecator2.append(len(temlist))
        temlist = []

    # rowindecator2를 내림차순으로 정렬하여 가장 큰 묶음 2개의 크기를 얻는다
    rowindecator2.sort(reverse=True)

    rowname = []
    # 가장 큰 묶음 list의 첫번째 row 번호와 마지막 row 번호를 얻는다
    for i in rowindecator:
        if len(i) == rowindecator2[0] or len(i) == rowindecator2[1]:
            rowname.append(i[0] + 2)
            rowname.append(i[-1] + 2)
            print(i[0], i[-1])
    # ===================================================================================================

    # Excel 실행
    excel = win32com.client.Dispatch("Excel.Application")

    # Excel 경고문 끄기
    excel.DisplayAlerts = False

    # Excel 보이게 하기
    excel.Visible = True

    # 해당 경로 excel workbook 열기
    wb = excel.Workbooks.Open(calpath)

    ws = wb.Worksheets

    # workbook의 worksheet 갯수를 보여준다
    sheetNb = wb.Sheets.count
    print(sheetNb)

    # 첫번째 sheet 가 계산시트고 이름을 claculator로 저장한다
    calculator = wb.sheets[0].Name
    lastsheet = wb.sheets[sheetNb - 1]

    # data 입력될 계산sheet 복사 생성
    ws(calculator).Copy(Before=lastsheet)

    # sheet가 추가되었으므로 sheetNb 갱신
    sheetNb = wb.Sheets.count
    print(sheetNb)

    sheetnamelist = []
    # workbook의 worksheet 이름을 보여준다
    for i in range(sheetNb):
        print(wb.Sheets[i].Name)
        sheetnamelist.append(wb.Sheets[i].Name)

    # workbook의 worksheet 이름을 변경한다
    if "{}번".format(sheet_name) not in sheetnamelist:
        wb.Sheets[sheetNb - 2].Name = "{}번".format(sheet_name)
    else:
        wb.Sheets[sheetNb - 2].Name = "{}번 (2)".format(sheet_name)

    # 현재 활성화된 sheet는 새로 만들어진 sheet이고 여기에 data 입력하기 위해 ws로 변수 설정
    ws = wb.Sheets[sheetNb - 2]

    # raw data 부르기
    wb2 = excel.Workbooks.Open(r'C:\Users\RyuTaeHyun\PycharmProjects\studyXL\{}.csv'.format(refine_name))
    ws2 = wb2.ActiveSheet

    def call_raw_data(start, end):
        wb2.Activate()
        ws2.Range("A{}:C{}".format(start, end)).Select()
        ws2.Range("A{}:C{}".format(start, end)).Copy()

    # raw data 붙여넣기
    call_raw_data("1", "10001")

    # 계산기 부르기
    wb.Activate()

    # 복사해 붙여넣기
    ws3 = wb.Sheets[sheetNb - 2]
    ws3.Range("B2:D2").Select()
    ws.Paste()

    # raw data 부르기
    call_raw_data(rowname[0], rowname[1])

    # 계산기 부르기
    wb.Activate()

    # 복사해 붙여넣기
    ws3.Range("I3:K3").Select()
    ws.Paste()

    # raw data 부르기
    call_raw_data(rowname[2], rowname[3])

    # 계산기 부르기
    wb.Activate()

    # 복사해 붙여넣기
    ws3.Range("M3:O3").Select()
    ws.Paste()

    # raw data 닫기
    wb2.Close(0)

    # workbook을 저장한다
    wb.Save()

    # workbook을 닫는다
    wb.Close(0)
    """
    # excel을 닫는다
    excel.Quit()
    """


