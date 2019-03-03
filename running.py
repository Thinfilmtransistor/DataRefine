from total_temporary import calculating, get_csv_name
import time


start_time = time.time()

# 계산기 경로
calpath = r"C:\Users\RyuTaeHyun\Documents\2Pr.xlsx"

# rawdata 디렉토리 경로
filepath = r"C:\Users\RyuTaeHyun\Documents"

# radata file 이름
filenamelist = get_csv_name(filepath)

for i in filenamelist:
    # rawdata file 절대경로
    rawdatapath = "{}\{}".format(filepath, i)
    calculating(rawdatapath, calpath)

# 프로그램 실행시간
print("--- %s seconds ---" % (time.time() - start_time))

