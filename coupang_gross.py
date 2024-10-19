import pandas as pd
import glob
import xlwings as xw
import re  # 정규식을 사용하기 위해 import

# 엑셀 파일 불러오기 (첫 번째 행을 데이터로 처리)
gross_file = glob.glob("Statistics*.xlsx")
df = pd.read_excel(gross_file[0], header=None)  # header=None으로 첫 번째 줄도 데이터로 취급

# 첫 번째 줄을 삭제
editdf = df.drop(index=0)

# 필요한 열 삭제
editdf = editdf.drop(columns=[0, 3, 4, 5])  # 열 번호로 삭제할 열 선택 (필요에 맞게 수정)

# 마지막 행의 첫 번째 열에 '합계' 삽입
editdf.iloc[-1, 0] = '합계'

# glob에서 가져온 파일명에서 날짜 추출 (정규식을 이용)
file_name = gross_file[0]  # 파일명 가져오기
date_match = re.search(r"Statistics-(\d{8})", file_name)  # Statistics- 뒤 8자리 숫자 추출
if date_match:
    file_date = date_match.group(1)  # 'YYYYMMDD' 형식의 날짜 추출
    file_date_formatted = f"{file_date[:4]}-{file_date[4:6]}-{file_date[6:]}"  # 'YYYY-MM-DD' 형식으로 변환

# xlwings로 엑셀 파일 열기
file_path = 'NEW판매관리.xlsm'
sheet_name = '그로스판매'

# Excel 애플리케이션
app = xw.App(visible=False)
try:
    # 엑셀 파일 열기
    wb = xw.Book(file_path)
    ws = wb.sheets[sheet_name]

    # 마지막 행 찾기 (B열 기준으로 마지막 행 찾기, A열은 비어있을 수 있으므로)
    last_row = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row

    # 새 데이터의 크기를 확인하여 범위를 지정
    num_rows, num_cols = editdf.shape

    # 추가할 범위 지정 (기존 데이터 다음 행에서 시작, B열부터 데이터를 넣음)
    target_range = ws.range(f'B{last_row + 1}').resize(num_rows, num_cols)

    # 데이터를 엑셀 시트에 추가 (B열부터)
    target_range.value = editdf.values

    # 새로 추가된 데이터 중 첫 번째 행의 A열에만 날짜 입력
    ws.range(f'A{last_row + 1}').value = file_date_formatted

    # 파일 저장 및 종료
    wb.save()

finally:
    wb.close()
    app.quit()
