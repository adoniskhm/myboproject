import win32com.client as win32
import os
from datetime import datetime

# 1. 현재 파이썬 파일이 실행되는 디렉토리에서 엑셀 파일을 엽니다.
current_dir = os.path.dirname(os.path.abspath(__file__))  # 현재 실행 중인 파이썬 파일의 경로
excel_file = os.path.join(current_dir, 'NEW수입철.xlsm')  # 엑셀 파일의 상대 경로

excel = win32.Dispatch('Excel.Application')
excel.Visible = True  # 엑셀 창을 사용자에게 보이게 함
wb = excel.Workbooks.Open(excel_file)  # 파일 경로 수정

# 2. '수입리스트' 시트에서 데이터 복사
sheet_import_list = wb.Sheets('수입리스트')
used_range = sheet_import_list.UsedRange

# 3. 새로운 시트 '수입품목' 생성
new_sheet = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
new_sheet.Name = '수입품목'

# 4. 값만 붙여넣기
used_range.Copy()
new_sheet.Range("A1").PasteSpecial(Paste=win32.constants.xlPasteValues)  # 값만 붙여넣기
excel.CutCopyMode = False  # 복사 모드 해제

# 5. 열 너비와 행 높이 복사
for col in range(1, used_range.Columns.Count + 1):
    new_sheet.Columns(col).ColumnWidth = sheet_import_list.Columns(col).ColumnWidth  # 열 너비 복사

for row in range(1, used_range.Rows.Count + 1):
    new_sheet.Rows(row).RowHeight = sheet_import_list.Rows(row).RowHeight  # 행 높이 복사

# 6. 이미지 복사 및 원래 위치에 붙여넣기
for shape in sheet_import_list.Shapes:
    try:
        shape.Copy()  # 이미지 복사
        new_sheet.Paste()  # 이미지 붙여넣기
        
        # 새로 붙여넣은 마지막 shape를 가져오기
        last_shape_index = new_sheet.Shapes.Count  # 현재 새 시트의 shape 개수
        new_shape = new_sheet.Shapes(last_shape_index)  # 마지막 shape 객체 가져오기
        
        # 원래 이미지의 위치를 가져와서 새 이미지에 적용
        new_shape.Left = shape.Left  # 원래 위치에 맞게 좌표 설정
        new_shape.Top = shape.Top  # 원래 위치에 맞게 좌표 설정
        
    except Exception:
        continue  # 오류 발생 시 해당 이미지를 건너뜀

# 7. A, E, I 열 삭제 (오른쪽에서 왼쪽으로)
for col in ['I', 'E', 'A']:
    new_sheet.Columns(col).Delete()

# 8. 첫 행의 데이터가 있는 부분 배경색 변경, 글꼴 굵게, 가운데 정렬
header_range = new_sheet.Rows(1)
header_range.Font.Bold = True  # 첫 행 글꼴 굵게
header_range.HorizontalAlignment = win32.constants.xlCenter  # 첫 행 가운데 정렬

for col in range(1, new_sheet.UsedRange.Columns.Count + 1):
    if new_sheet.Cells(1, col).Value is not None:  # 첫 행의 데이터가 있는 경우
        new_sheet.Cells(1, col).Interior.Color = 0xE2EFDA  # 배경색 변경
    else:
        new_sheet.Cells(1, col).Interior.Pattern = win32.constants.xlNone  # 배경색 제거

# 9. I2 셀에 'MBK' 추가 및 가운데 정렬
new_sheet.Cells(2, 9).Value = "MBK"  # I2 셀에 'MBK' 추가
# new_sheet.Cells(2, 9).HorizontalAlignment = win32.constants.xlCenter  # 가운데 정렬

# 9-1. 모든 셀을 가운데 정렬
new_sheet.UsedRange.HorizontalAlignment = win32.constants.xlCenter

# 10. 글꼴 및 글꼴 크기 설정
new_sheet.Cells.Font.Name = '굴림'  # 글꼴 변경
new_sheet.Cells.Font.Size = 9  # 글꼴 크기 변경

# 11. 모든 실수 반올림하여 소수점 둘째 자리까지 표시
for row in range(1, new_sheet.UsedRange.Rows.Count + 1):
    for col in range(1, new_sheet.UsedRange.Columns.Count + 1):
        cell_value = new_sheet.Cells(row, col).Value
        if isinstance(cell_value, float):  # 셀 값이 실수인지 확인
            # H열의 값이 날짜일 경우
            if col == 8:  # H열 (8번째 열)이면 날짜 형식으로 설정
                new_sheet.Cells(row, col).NumberFormat = 'yyyy-mm-dd'  # 날짜 형식 지정
            else:
                new_sheet.Cells(row, col).Value = round(cell_value, 2)  # 반올림하여 두 자리 표시

# 12. '수입품목' 시트를 제외한 모든 시트 삭제
excel.DisplayAlerts = False  # 삭제 확인 알림 비활성화
for sheet in wb.Sheets:
    if sheet.Name != '수입품목':
        wb.Sheets(sheet.Name).Delete()
excel.DisplayAlerts = True  # 알림 다시 활성화

# 13. 현재 날짜를 기반으로 파일을 저장
today = datetime.today().strftime('%Y%m%d')  # 오늘 날짜를 YYYYMMDD 형식으로 저장
save_path = os.path.join(current_dir, f'수입품목리스트_{today}.xlsx')  # 파일명 생성

wb.SaveAs(save_path, FileFormat=51)  # 새로운 이름으로 .xlsx 형식으로 저장 (51 = xlOpenXMLWorkbook)
print(f'파일이 {save_path}에 저장되었습니다.')

# 엑셀을 종료하지 않고 작업을 끝냄
